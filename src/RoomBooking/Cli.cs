using System.Globalization;
using System.Reflection;
using System.Text.Json.Nodes;
using M365Graph;
using MeetingRooms;
using WinBackground;

namespace RoomBooking;

/// <summary>Everything the command line touches outside itself, so tests can give it fakes.</summary>
public sealed class CliHost
{
    public TextWriter Out { get; init; } = Console.Out;
    public TextWriter Err { get; init; } = Console.Error;
    public Func<string, string?> Env { get; init; } = Config.Environment;
    public Func<DateTime> Now { get; init; } = () => DateTime.Now;
    public Func<Config, ICalendar> Calendar { get; init; } = config => new Calendar(RoomBooking.Calendar.Client(config), config.TimeZone, config.LoginHint);
    public Action<Config> SignIn { get; init; } = config =>
        RoomBooking.Calendar.Client(config).TokenAsync(RoomBooking.Calendar.Scopes, true).GetAwaiter().GetResult();
    public Action<string, IEnumerable<string>> Notify { get; init; } = Toast.Notify;
    public Action<TimeSpan> Sleep { get; init; } = Thread.Sleep;
    public Func<ITask> Task { get; init; } = () => new ScheduledTask(Cli.TaskName,
        "Books meeting rooms for your Teams meetings on your office days, by your rules (calendar-room-booking run).");
    public string ProgramDir { get; init; } = AppContext.BaseDirectory;
    public bool Lock { get; init; } = true;
}

/// <summary>calendar-room-booking: meeting rooms for the Teams meetings you're going to on your office days, chosen by
/// rules. Exit codes: 0 done; 1 failed; 2 settings, arguments or sign-in.</summary>
public static class Cli
{
    public const string Prog = "calendar-room-booking";
    public const string TaskName = "calendar-room-booking run";
    public const string BackgroundExe = "calendar-room-booking-background.exe";

    private const string Usage = "usage: calendar-room-booking [-h] [--version] [--config CONFIG] {plan,run,book,skip,list,rooms,autostart,sign-in,config} ...";

    // Explicit "\n" throughout (a raw literal would take the .cs file's line endings).
    public const string Help =
        Usage + "\n\n" +
        "calendar-room-booking: meeting rooms for the Teams meetings you're going to on your office days, chosen by rules.\n" +
        "A room is added to a meeting you organise (everyone gets an update) and the invitation is forwarded to the room for\n" +
        "one you attend (as Outlook's Add room does). A meeting has a room only when the room's own schedule shows it.\n\n" +
        "    calendar-room-booking plan                    each office day: the meetings, the rule, the room, and why\n" +
        "    calendar-room-booking run                     ask mode: review items only; act mode: book the \"yes\" meetings\n" +
        "    calendar-room-booking book MEETING [--room R] [--yes]   book one (a preview without --yes)\n" +
        "    calendar-room-booking skip MEETING            no room for that one (at that time)\n" +
        "    calendar-room-booking list                    the rooms it has asked for, and the meetings skipped\n" +
        "    calendar-room-booking rooms                   the rooms in the order they're tried, per rule\n" +
        "    calendar-room-booking autostart on            run on the [schedule] (off, restart, status)\n" +
        "    calendar-room-booking sign-in                 sign in to Microsoft (once; opens a browser)\n" +
        "    calendar-room-booking config                  create the settings file (if missing) and say where it is\n\n" +
        "MEETING is an event id, or \"YYYY-MM-DD HH:MM\" (its start).\n\n" +
        "options:\n" +
        "  -h, --help       show this help message and exit\n" +
        "  --version        show program's version number and exit\n" +
        "  --config CONFIG  settings file (default %APPDATA%\\calendar-room-booking\\config.toml)\n";

    private static readonly string[] Commands = ["plan", "run", "book", "skip", "list", "rooms", "autostart", "sign-in", "config"];
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    public static string Version =>
        typeof(Cli).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";

    internal sealed class Options
    {
        public string? Config;
        public string Command = "";
        public List<string> Args = [];
        public string? Room;
        public bool Yes;
    }

    private sealed class ArgumentsException(string message) : Exception(message);

    /// <summary>A run that stopped itself rather than write something that looks wrong.</summary>
    public sealed class SafetyStop(string message) : Exception(message);

    public static int Run(IReadOnlyList<string> argv, CliHost? host = null)
    {
        host ??= new CliHost();
        Options options;
        try
        {
            var parsed = Parse(argv, host);
            if (parsed is null)
                return 0;
            options = parsed;
        }
        catch (ArgumentsException error)
        {
            host.Err.WriteLine(Usage);
            host.Err.WriteLine($"{Prog}: error: {error.Message}");
            return 2;
        }

        Config? config = null;
        try
        {
            if (options.Command == "config")
            {
                host.Out.WriteLine(RoomBooking.Config.CreateDefault(options.Config ?? host.Env(RoomBooking.Config.EnvConfig), host.Env));
                return 0;
            }
            config = RoomBooking.Config.Load(options.Config, host.Env);
            switch (options.Command)
            {
                case "sign-in":
                    host.SignIn(config);
                    host.Out.WriteLine("Signed in.");
                    return 0;
                case "autostart":
                    return Autostart(config, options.Args[0], host);
                case "rooms":
                    Rooms(config, host);
                    return 0;
                case "list":
                    List(config, host);
                    return 0;
            }
            using var held = host.Lock ? Locks.Exclusive("calendar-room-booking-" + Locks.Key(config.Path), 300) : null;
            var code = options.Command switch
            {
                "plan" => Apply(config, host, write: false),
                "run" => Apply(config, host, write: true),
                "book" => BookOne(config, host, options),
                _ => SkipOne(config, host, options),
            };
            if (options.Command == "run")
                ForgetFailure(config);
            return code;
        }
        catch (Exception error) when (error is UsageException or SignInNeededException)
        {
            host.Err.WriteLine(error.Message);
            if (options.Command == "run" && config is not null)
                Failed(config, host, error);
            return 2;
        }
        catch (Exception error)
        {
            host.Err.WriteLine(error is SafetyStop ? error.Message : $"{error.GetType().Name}: {error.Message}");
            if (options.Command == "run" && config is not null)
                Failed(config, host, error);
            return 1;
        }
    }

    private static (DateOnly First, DateOnly Last) Period(Config config, CliHost host)
    {
        var today = DateOnly.FromDateTime(host.Now());
        return (today, today.AddDays(config.DaysAhead));
    }

    /// <summary>Reads the calendar, the work plan and the rooms' schedules, and judges every office day.</summary>
    private static (List<DayPlan> Plans, ICalendar Calendar) Judge(Config config, CliHost host, State state, DateOnly first, DateOnly last)
    {
        var calendar = host.Calendar(config);
        var now = host.Now();
        var me = calendar.Me();
        var plans = Planner.Days(config, me, now, calendar.WorkPlan(first, last), calendar.Events(first, last), first, last);
        var addresses = Planner.RoomsToRead(config, plans.Where(p => p.Office));
        var schedules = plans.Any(p => p.Office && p.Candidates.Count > 0) && addresses.Count > 0
            ? calendar.RoomSchedules(addresses, first.ToDateTime(TimeOnly.MinValue), last.AddDays(1).ToDateTime(TimeOnly.MinValue))
            : new Dictionary<string, IReadOnlyList<ScheduleItem>?>();
        Planner.Judge(config, plans, schedules, state, now);
        return (plans, calendar);
    }

    private static int Apply(Config config, CliHost host, bool write)
    {
        var log = write ? Logger.Create(config.LogFile, null, clock: host.Now) : (_ => { });
        if (!config.Enabled)
        {
            host.Out.WriteLine("Switched off (enabled = false in the settings): nothing done.");
            return 0;
        }
        if (write && !config.Schedule.IsActive(host.Now(), out var why))
        {
            host.Out.WriteLine($"Outside the [schedule] ({why}): nothing done.");
            return 0;
        }
        var state = State.Load(config.StateFile);
        var (first, last) = Period(config, host);
        var (plans, calendar) = Judge(config, host, state, first, last);
        var book = config.Act && write ? plans.SelectMany(p => p.Candidates).Where(c => c.Status == Status.Book).ToList() : [];
        if (book.Count > config.MaxBookings)
            throw new SafetyStop($"A run would book {book.Count} rooms, more than max_bookings ({config.MaxBookings}), so nothing was "
                                 + "booked. Check `calendar-room-booking plan`; raise max_bookings if it's right.");

        if (!write)
        {
            host.Out.WriteLine(config.Act
                ? "mode = act: a run books the meetings marked \"book\" and asks about the rest in the review queue."
                : "mode = ask: a run books nothing; it asks about these in the review queue.");
            foreach (var p in plans)
            {
                host.Out.WriteLine($"  {Planner.DayName(p.Day)}  {(p.Office ? "office" : "-"),-6}  ({p.Why})");
                foreach (var c in p.Candidates)
                    host.Out.WriteLine(Planner.Line(c, config.Act));
                foreach (var o in p.Others.Where(o => !o.Meeting.AllDay))
                    host.Out.WriteLine($"      - {o.Meeting.Start:HH:mm}-{o.Meeting.End:HH:mm}  {Planner.Clip(o.Meeting.Subject, 40)}: {o.Why}");
            }
        }

        var outcomes = new Dictionary<Outcome, int>();
        foreach (var c in book.OrderBy(c => c.Meeting.Start))
        {
            var (outcome, message) = Booker.Book(config, calendar, state, c.Meeting, c.Choice!, log, host.Sleep, host.Now(), ActivityJournal.By.Auto, c.WantWhy);
            if (outcome == Outcome.Declined && state.Attempts(c.Meeting.Key).Count(a => a.State == "done") < config.MaxAttempts
                && c.Order.FirstOrDefault(o => o.Room != c.Choice && !state.Attempts(c.Meeting.Key).Any(a => string.Equals(a.Room, o.Room.Address, StringComparison.OrdinalIgnoreCase)))
                    is { } next && NextFree(calendar, next.Room, c.Meeting))
            {
                log($"{c.Choice!.Name} declined {c.Meeting.Start:dd/MM HH:mm} \"{c.Meeting.Subject}\": trying {next.Room.Name}");
                (outcome, message) = Booker.Book(config, calendar, state, c.Meeting, next.Room, log, host.Sleep, host.Now(), ActivityJournal.By.Auto,
                    $"{c.WantWhy}; {c.Choice!.Name} declined");
            }
            outcomes[outcome] = outcomes.GetValueOrDefault(outcome) + 1;
            host.Out.WriteLine($"  {c.Meeting.Start:ddd dd/MM HH:mm} {c.Meeting.Subject}: {message}");
            if (outcome is Outcome.MayHaveGone or Outcome.Declined)
            {
                c.Status = outcome == Outcome.MayHaveGone ? Status.MayHaveGone : Status.GaveUp;
                c.StatusWhy = message;
            }
        }

        var raised = 0;
        if (write)
        {
            raised = Booker.Raise(config, plans, state, host.Now(), log);
            state.Prune(first);
            state.Save(config.StateFile);
        }

        var all = plans.SelectMany(p => p.Candidates).ToList();
        int Count(Status s) => all.Count(c => c.Status == s);
        var summary = $"{first:dd/MM} to {last:dd/MM}: {plans.Count(p => p.Office)} office days, {all.Count} meetings: "
                      + $"{Count(Status.Booked)} with a room, {Count(Status.Book)} {(config.Act ? "to book" : "the rules would book")}, "
                      + $"{Count(Status.Ask)} to ask, {Count(Status.No) + Count(Status.Skipped)} no"
                      + (Count(Status.NoRoom) + Count(Status.GaveUp) > 0 ? $", {Count(Status.NoRoom) + Count(Status.GaveUp)} with no room free" : "")
                      + (Count(Status.Waiting) > 0 ? $", {Count(Status.Waiting)} waiting for a room's answer" : "")
                      + (Count(Status.MayHaveGone) > 0 ? $", {Count(Status.MayHaveGone)} to check (may have gone)" : "");
        if (!write)
            host.Out.WriteLine(summary + " (nothing changed)");
        else
        {
            var did = book.Count == 0 ? "nothing booked"
                : string.Join(", ", outcomes.Select(kv => $"{kv.Value} {kv.Key.ToString().ToLowerInvariant()}"));
            host.Out.WriteLine($"{summary}; {did}; {raised} item(s) raised");
            log($"{summary}; {did}; {raised} item(s) raised");
        }
        return outcomes.ContainsKey(Outcome.MayHaveGone) ? 1 : 0;
    }

    private static bool NextFree(ICalendar calendar, Room room, Meeting m)
    {
        var s = calendar.RoomSchedules([room.Address], m.Start.Date, m.Start.Date.AddDays(1));
        return s.TryGetValue(room.Address, out var items) && items is not null && RoomSchedule.Free(items, m.Start, m.End);
    }

    /// <summary>The meeting MEETING names, and its candidate (null when it isn't one, with the reason).</summary>
    private static (Meeting Meeting, Candidate? Candidate, string? NotWhy, ICalendar Calendar, State State) Find(Config config, CliHost host, string target, bool unskip = false)
    {
        var state = State.Load(config.StateFile);
        var calendar = host.Calendar(config);
        Meeting meeting;
        if (DateTime.TryParseExact(target.Trim(), ["yyyy-MM-dd HH:mm", "yyyy-MM-ddTHH:mm"], En, DateTimeStyles.None, out var at))
        {
            var day = DateOnly.FromDateTime(at);
            var found = calendar.Events(day, day).Where(e => !e.Cancelled && e.Start == at).ToList();
            if (found.Count == 0)
                throw new UsageException($"No meeting starts at {at:ddd dd MMM HH:mm} in your calendar.");
            if (found.Count > 1)
                throw new UsageException($"{found.Count} meetings start at {at:ddd dd MMM HH:mm}: give the event id instead.\n"
                                         + string.Join("\n", found.Select(e => $"  {e.Subject}: {e.Id}")));
            meeting = found[0];
        }
        else
            meeting = calendar.Event(target.Trim()) ?? throw new UsageException($"No event with id {target.Trim()} in your calendar.");

        if (unskip)
            state.Skipped.Remove(meeting.Key);   // booking it is changing your mind (saved only if it's booked)
        var first = DateOnly.FromDateTime(meeting.Start);
        var now = host.Now();
        var me = calendar.Me();
        var plans = Planner.Days(config, me, now, calendar.WorkPlan(first, first), calendar.Events(first, first), first, first);
        var schedules = calendar.RoomSchedules(Planner.RoomsToRead(config, plans).Concat(meeting.ListedRooms()).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            first.ToDateTime(TimeOnly.MinValue), first.AddDays(1).ToDateTime(TimeOnly.MinValue));
        if (plans.FirstOrDefault() is { } plan && !plan.Office && plan.Candidates.Count == 0)
            plans[0] = plan with { Office = true };   // judged even on a day that isn't an office day: you asked for it
        var candidate = plans.SelectMany(p => p.Candidates).FirstOrDefault(c => c.Meeting.Id == meeting.Id);
        string? notWhy = null;
        if (candidate is null)
        {
            notWhy = plans.SelectMany(p => p.Others).FirstOrDefault(o => o.Meeting.Id == meeting.Id)?.Why
                     ?? (plans.FirstOrDefault() is { Office: false } p2 ? p2.Why : "not on an office day");
            var people = Planner.People(config, me, meeting);
            var (want, rule, wantWhy) = Planner.Decide(config, meeting, people);
            candidate = new Candidate
            {
                Meeting = meeting, People = people.Count, Want = want, WantWhy = wantWhy, Rule = rule,
                Order = Planner.Order(config, rule, people.Count),
            };
            plans.Add(new DayPlan(first, true, "", [], [candidate], []));
        }
        Planner.Judge(config, plans, schedules, state, now);
        return (meeting, candidate, notWhy, calendar, state);
    }

    private static int BookOne(Config config, CliHost host, Options options)
    {
        var (m, c, notWhy, calendar, state) = Find(config, host, options.Args[0], unskip: true);
        host.Out.WriteLine($"{m.Start:ddd dd MMM HH:mm}-{m.End:HH:mm}  {m.Subject}  ({c!.People} people; organiser {m.OrganiserName})");
        if (notWhy is not null)
            host.Out.WriteLine($"  Not one it would pick ({notWhy}); booked because you asked.");
        if (c.Status == Status.Booked)
        {
            host.Err.WriteLine($"It already has a room: {c.BookedRoom} ({c.StatusWhy}). Nothing sent.");
            return 2;
        }
        if (c.Status == Status.MayHaveGone)
        {
            host.Err.WriteLine($"{c.StatusWhy} Check the room in Outlook.");
            return 2;
        }
        Room room;
        string why;
        if (options.Room is { } wanted)
        {
            room = RoomList.Find(config.Rooms, wanted) ?? throw new UsageException(
                $"{wanted} isn't one of the [[rooms]] ({string.Join(", ", config.Rooms.Select(r => r.Name))}): only a listed room is booked. Nothing sent.");
            if (!NextFree(calendar, room, m))
            {
                host.Err.WriteLine($"{room.Name} isn't free {m.Start:HH:mm}-{m.End:HH:mm}. Nothing sent.");
                return 2;
            }
            why = "your choice";
        }
        else if (c.Choice is { } choice)
            (room, why) = (choice, c.ChoiceWhy);
        else
        {
            host.Err.WriteLine($"No room is free that seats {c.People} ({c.StatusWhy}). Nothing sent.");
            return 2;
        }
        host.Out.WriteLine($"  room   {room.Name} <{room.Address}> ({why})");
        host.Out.WriteLine(m.IsOrganiser
            ? "  how    added to your meeting as a room and a location: everyone on it gets an update"
            : "  how    the invitation forwarded to the room, as Outlook's Add room does (Exchange tells the organiser)");
        if (!options.Yes)
        {
            host.Out.WriteLine("Nothing sent. Run again with --yes to book it.");
            return 0;
        }
        var log = Logger.Create(config.LogFile, null, clock: host.Now);
        state.Skipped.Remove(m.Key);
        var (outcome, message) = Booker.Book(config, calendar, state, m, room, log, host.Sleep, host.Now(), ActivityJournal.By.Hand, null);
        state.Save(config.StateFile);
        host.Out.WriteLine(message);
        return outcome is Outcome.Accepted or Outcome.NoAnswer ? 0 : 1;
    }

    private static int SkipOne(Config config, CliHost host, Options options)
    {
        var (m, _, _, _, state) = Find(config, host, options.Args[0]);
        state.Skip(m, host.Now());
        state.Save(config.StateFile);
        var log = Logger.Create(config.LogFile, null, clock: host.Now);
        log($"skip: {m.Start:dd/MM HH:mm} \"{m.Subject}\"");
        Activity.Skipped(config, host.Now(), log, m);
        host.Out.WriteLine($"No room for {m.Start:ddd dd MMM HH:mm} {m.Subject} (at that time). `book` it to change your mind.");
        return 0;
    }

    private static void List(Config config, CliHost host)
    {
        var state = State.Load(config.StateFile);
        var rows = state.Bookings.Select(kv => (Key: kv.Key, Entry: kv.Value as JsonObject)).Where(x => x.Entry is not null)
            .OrderBy(x => State.Str(x.Entry, "start"), StringComparer.Ordinal).ToList();
        host.Out.WriteLine(rows.Count == 0 ? "No rooms asked for." : "Rooms asked for:");
        foreach (var (key, entry) in rows)
        {
            var start = DateTime.TryParse(State.Str(entry, "start"), En, DateTimeStyles.None, out var s) ? s : default;
            var end = DateTime.TryParse(State.Str(entry, "end"), En, DateTimeStyles.None, out var e) ? e : default;
            var attempts = state.Attempts(key).Select(a =>
                $"{a.Name} ({(a.How == "add" ? "added" : "forwarded")} {a.When:dd/MM HH:mm}; {(a.State == "sending" ? "MAY HAVE GONE" : a.Answer ?? "no answer read")})");
            host.Out.WriteLine($"  {s:ddd dd MMM HH:mm}-{end:HH:mm}  {State.Str(entry, "subject")}: {string.Join("; then ", attempts)}");
        }
        var skipped = state.Skipped.Select(kv => kv.Value as JsonObject).OfType<JsonObject>().OrderBy(o => State.Str(o, "start"), StringComparer.Ordinal).ToList();
        if (skipped.Count > 0)
        {
            host.Out.WriteLine("Skipped (no room):");
            foreach (var o in skipped)
                host.Out.WriteLine($"  {State.Str(o, "start").Replace('T', ' ')}  {State.Str(o, "subject")}");
        }
    }

    private static void Rooms(Config config, CliHost host)
    {
        if (config.Rooms.Count == 0)
        {
            host.Out.WriteLine("No [[rooms]] in the settings: nothing can be booked.");
            return;
        }
        void Show(string title, IReadOnlyList<string> preferred)
        {
            host.Out.WriteLine(title);
            foreach (var r in RoomOrder.Rank(config.Rooms, preferred, config.RoomFloors, 1))
                host.Out.WriteLine($"    {r.Room.Name,-40} {r.Room.Address,-36} {r.Why}");
        }
        Show("Default order (default_rooms, then room_floors, then smallest first; a room too small for a meeting is left out):", config.DefaultRooms);
        foreach (var rule in config.Rules.Where(r => r.Rooms is not null))
            Show($"{rule.Label}:", [.. rule.Rooms!, .. config.DefaultRooms]);
    }

    /// <summary>A run that failed: logged, and notified once (on the third failure in a row, or at once for a sign-in
    /// or a safety stop) until a run succeeds again.</summary>
    private static void Failed(Config config, CliHost host, Exception error)
    {
        try
        {
            Logger.Create(config.LogFile, null, clock: host.Now)($"run failed: {error.GetType().Name}: {error.Message}");
            var state = LoadFailures(config.StateFile);
            var failures = (state["failures"]?.GetValue<int>() ?? 0) + 1;
            state["failures"] = failures;
            if (state["notified"] is null && (failures >= 3 || error is SignInNeededException or SafetyStop))
            {
                if (config.Notify)
                    host.Notify("calendar-room-booking couldn't run", [error.Message, $"Log: {config.LogFile}"]);
                state["notified"] = host.Now().ToString("yyyy-MM-ddTHH:mm:ss", En);
            }
            new State(state).Save(config.StateFile);
        }
        catch (Exception)
        {
            // nothing more can be done: the exit code says it failed
        }
    }

    private static JsonObject LoadFailures(string path)
    {
        try
        {
            return File.Exists(path) ? JsonNode.Parse(File.ReadAllText(path)) as JsonObject ?? [] : [];
        }
        catch (Exception error) when (error is IOException or System.Text.Json.JsonException)
        {
            return [];
        }
    }

    private static void ForgetFailure(Config config)
    {
        var state = LoadFailures(config.StateFile);
        if (state.Remove("failures") | state.Remove("notified"))
            new State(state).Save(config.StateFile);
    }

    private static int Autostart(Config config, string action, CliHost host)
    {
        var task = host.Task();
        switch (action)
        {
            case "on":
                var background = Path.Combine(host.ProgramDir, BackgroundExe);
                var exe = File.Exists(background) ? background : Path.Combine(host.ProgramDir, "calendar-room-booking.exe");
                var schedule = config.Schedule.ToTaskSchedule(TaskKind.Job, TimeSpan.FromMinutes(10));
                var line = task.Enable(exe, ["--config", config.Path, "run"], Path.GetDirectoryName(config.Path), schedule);
                host.Out.WriteLine($"Autostart is on: '{TaskName}' runs {Describe(config.Schedule)}, and has run now.\n  {line}");
                return 0;
            case "off":
                host.Out.WriteLine(task.Disable() ? "Autostart is off." : "Autostart was already off.");
                return 0;
            case "restart":
                try
                {
                    host.Out.WriteLine(task.Restart() ? $"Started '{TaskName}' now." : "Autostart is off, so there's nothing to restart.");
                }
                catch (InvalidOperationException error)
                {
                    host.Err.WriteLine(error.Message);
                    return 1;
                }
                return 0;
            default:
                var info = task.Status();
                if (info is null)
                    host.Out.WriteLine("Autostart is off. `calendar-room-booking autostart on` turns it on.");
                else
                {
                    host.Out.WriteLine($"Autostart is on ('{TaskName}').");
                    foreach (var name in new[] { "state", "last_run", "last_result", "next_run", "command" })
                        host.Out.WriteLine($"  {name.Replace('_', ' '),-12} {info.GetValueOrDefault(name, "")}");
                    if (ScheduleDrift.Advice(ScheduleDrift.Check(info, config.Schedule.Fingerprint(TaskKind.Job)), "calendar-room-booking") is { } advice)
                        host.Out.WriteLine($"  {advice}");
                }
                return 0;
        }
    }

    /// <summary>The task's triggers in words: "every 15 minutes, at logon".</summary>
    internal static string Describe(ScheduleSettings s)
    {
        var parts = new List<string>();
        if (s.EveryMinutes is > 0 and var m)
            parts.Add($"every {m} minutes");
        else if (s.DailyAt is { } at)
            parts.Add($"daily at {at}");
        if (s.AtLogon)
            parts.Add("at logon");
        if (s.AfterUnlockMinutes > 0)
            parts.Add($"{s.AfterUnlockMinutes} minute(s) after each unlock");
        return string.Join(", ", parts);
    }

    private static Options? Parse(IReadOnlyList<string> argv, CliHost host)
    {
        var options = new Options();
        var i = 0;
        for (; i < argv.Count && argv[i].StartsWith('-'); i++)
        {
            var arg = argv[i];
            if (arg is "-h" or "--help")
            {
                host.Out.Write(Help);
                return null;
            }
            if (arg == "--version")
            {
                host.Out.WriteLine(Version);
                return null;
            }
            if (arg == "--config")
            {
                if (++i >= argv.Count)
                    throw new ArgumentsException("argument --config: expected one argument");
                options.Config = argv[i];
            }
            else if (arg.StartsWith("--config=", StringComparison.Ordinal))
                options.Config = arg["--config=".Length..];
            else
                throw new ArgumentsException($"unrecognized arguments: {arg}");
        }
        if (i >= argv.Count)
            throw new ArgumentsException("the following arguments are required: command");
        options.Command = argv[i++];
        if (!Commands.Contains(options.Command))
            throw new ArgumentsException($"argument command: invalid choice: '{options.Command}' (choose from {string.Join(", ", Commands.Select(c => $"'{c}'"))})");
        for (; i < argv.Count; i++)
        {
            var arg = argv[i];
            switch (options.Command, arg)
            {
                case (_, "-h" or "--help"):
                    host.Out.Write(Help);
                    return null;
                case ("book", "--yes"):
                    options.Yes = true;
                    break;
                case ("book", "--room"):
                    if (++i >= argv.Count)
                        throw new ArgumentsException("argument --room: expected one argument");
                    options.Room = argv[i];
                    break;
                default:
                    if (arg.StartsWith('-') && arg != "-")
                        throw new ArgumentsException($"unrecognized arguments: {arg}");
                    options.Args.Add(arg);
                    break;
            }
        }
        if (options.Command is "book" or "skip")
        {
            // "2026-11-26 10:00" as one argument, or as two
            if (options.Args.Count == 2 && DateOnly.TryParseExact(options.Args[0], "yyyy-MM-dd", En, DateTimeStyles.None, out _))
                options.Args = [$"{options.Args[0]} {options.Args[1]}"];
            if (options.Args.Count == 0)
                throw new ArgumentsException("the following arguments are required: meeting (an event id, or \"YYYY-MM-DD HH:MM\")");
        }
        var max = options.Command is "autostart" or "book" or "skip" ? 1 : 0;
        if (options.Command == "autostart" && options.Args.Count == 0)
            throw new ArgumentsException("the following arguments are required: action");
        if (options.Args.Count > max)
            throw new ArgumentsException($"unrecognized arguments: {string.Join(" ", options.Args.Skip(max))}");
        if (options.Command == "autostart" && options.Args[0] is not ("on" or "off" or "restart" or "status"))
            throw new ArgumentsException($"argument action: invalid choice: '{options.Args[0]}' (choose from 'on', 'off', 'restart', 'status')");
        return options;
    }
}
