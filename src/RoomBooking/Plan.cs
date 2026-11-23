using System.Globalization;
using CalendarDays;
using MeetingRooms;

namespace RoomBooking;

/// <summary>Where a candidate meeting stands.</summary>
public enum Status
{
    /// <summary>A room's own schedule shows this meeting.</summary>
    Booked,
    /// <summary>Wants a room ("yes"): a run in act mode books <see cref="Candidate.Choice"/>.</summary>
    Book,
    /// <summary>Its rule says ask: raised in the review queue.</summary>
    Ask,
    /// <summary>Its rule (or category) says no.</summary>
    No,
    /// <summary>Wants a room, but none is free that fits.</summary>
    NoRoom,
    /// <summary>You said no to this one (`skip`).</summary>
    Skipped,
    /// <summary>A booking was sent and its answer lost: never sent again.</summary>
    MayHaveGone,
    /// <summary>A room was asked a short while ago and hasn't shown the booking yet.</summary>
    Waiting,
    /// <summary>max_attempts rooms were asked and none took it.</summary>
    GaveUp,
}

/// <summary>A meeting on an office day that could have a room: how many are at it, what the rules say, the rooms in the
/// order they'd be tried, and (once judged) where it stands and the room chosen.</summary>
public sealed class Candidate
{
    public required Meeting Meeting { get; init; }
    public required int People { get; init; }
    public required Want Want { get; init; }
    public required string WantWhy { get; init; }
    public Rule? Rule { get; init; }
    public required IReadOnlyList<RankedRoom> Order { get; init; }

    public Status Status { get; set; }
    public string? BookedRoom { get; set; }
    public string StatusWhy { get; set; } = "";
    public Room? Choice { get; set; }
    public string ChoiceWhy { get; set; } = "";
    public IReadOnlyList<Attempt> Attempts { get; set; } = [];

    /// <summary>How a room gets onto it: added to a meeting you organise, forwarded to the room otherwise.</summary>
    public string How => Meeting.IsOrganiser ? "add" : "forward";
}

/// <summary>A meeting on an office day that isn't a candidate, and why.</summary>
public sealed record Excluded(Meeting Meeting, string Why);

/// <summary>A day: office or not and why, its office hours, its candidates and the meetings left out.</summary>
public sealed record DayPlan(DateOnly Day, bool Office, string Why, IReadOnlyList<(DateTime Start, DateTime End)> Hours,
    List<Candidate> Candidates, List<Excluded> Others);

public static class Planner
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    /// <summary>Each weekday from first to last: whether it's an office day (work plan office time, and not leave or a
    /// bank holiday), and its candidates with what the rules say. Pure.</summary>
    public static List<DayPlan> Days(Config config, string me, DateTime now, IReadOnlyList<Segment> workPlan,
        IReadOnlyList<Meeting> events, DateOnly first, DateOnly last)
    {
        var dayEvents = events.Where(e => !e.Cancelled && e.Response != "declined")
            .Select(e => new DayEvent(e.Subject, e.Start, e.End, e.AllDay, e.ShowAs, e.Categories, e.IsOrganiser,
                e.Attendees.Count(a => a.Type != "resource" && !Same(a.Address, me))))
            .ToList();
        var plans = new List<DayPlan>();
        for (var day = first; day <= last; day = day.AddDays(1))
        {
            if (DayRules.Weekend(day))
                continue;
            var start = day.ToDateTime(TimeOnly.MinValue);
            var end = start.AddDays(1);
            var off = DayOffRules.Classify(config.Leave, dayEvents, day);
            if (off.Kind is DayOffKind.BankHoliday or DayOffKind.Leave)
            {
                plans.Add(new DayPlan(day, false, $"{(off.Kind == DayOffKind.BankHoliday ? "bank holiday" : "leave")}: {off.Entry?.Subject}", [], [], []));
                continue;
            }
            var hours = Merge(workPlan.Where(s => s.Location == "office" && s.Start < end && s.End > start)
                .Select(s => (s.Start < start ? start : s.Start, s.End > end ? end : s.End)));
            if (hours.Count == 0)
            {
                var kinds = workPlan.Where(s => s.Start < end && s.End > start).Select(s => s.Location).Distinct().ToList();
                plans.Add(new DayPlan(day, false, kinds.Count == 0 ? "no work plan" : $"work plan: {string.Join(", ", kinds)}", [], [], []));
                continue;
            }
            var why = "work plan: office " + string.Join(", ", hours.Select(h => $"{h.Start:HH:mm}-{h.End:HH:mm}"));
            if (off.Kind == DayOffKind.HalfDay)
                why += $"; half day off ({off.Half?.ToString() ?? "?"}): {off.Entry?.Subject}";
            var plan = new DayPlan(day, true, why, hours, [], []);
            foreach (var m in events.Where(e => !e.Cancelled && e.Start < end && e.End > start).OrderBy(e => e.Start).ThenBy(e => e.Subject))
            {
                if (Exclude(config, me, now, m, hours, off) is { } reason)
                {
                    plan.Others.Add(new Excluded(m, reason));
                    continue;
                }
                var people = People(config, me, m);
                var (want, rule, wantWhy) = Decide(config, m, people);
                plan.Candidates.Add(new Candidate
                {
                    Meeting = m, People = people.Count, Want = want, WantWhy = wantWhy, Rule = rule,
                    Order = Order(config, rule, people.Count),
                });
            }
            plans.Add(plan);
        }
        return plans;
    }

    /// <summary>The rooms a meeting would try, in order: its rule's rooms, then default_rooms, then room_floors, then
    /// the smallest that seats everyone (background-tools-lib's RoomOrder).</summary>
    public static IReadOnlyList<RankedRoom> Order(Config config, Rule? rule, int seats) =>
        RoomOrder.Rank(config.Rooms, [.. rule?.Rooms ?? [], .. config.DefaultRooms], config.RoomFloors, seats);

    /// <summary>Why a meeting on an office day isn't a candidate; null when it is.</summary>
    internal static string? Exclude(Config config, string me, DateTime now, Meeting m, IReadOnlyList<(DateTime Start, DateTime End)> hours, DayOff off)
    {
        if (m.AllDay)
            return "all day";
        if (!m.IsOrganiser && !config.CandidateResponses.Contains(m.Response))
            return m.Response == "declined" ? "declined" : $"not accepted ({m.Response})";
        if (!m.Online)
            return "not a Teams meeting";
        var others = People(config, me, m).Count - 1;
        if (others == 0)
            return "your own entry (no one else)";
        if (string.Equals(m.ShowAs, "free", StringComparison.OrdinalIgnoreCase))
            return "shown as free";
        foreach (var s in config.Skip)
            if ((s.Subject is null || s.Subject.IsMatch(m.Subject))
                && (s.Category is null || m.Categories.Contains(s.Category, StringComparer.OrdinalIgnoreCase)))
                return $"skip: {s.Why}";
        if (!hours.Any(h => h.Start <= m.Start && m.End <= h.End))
            return "outside office hours";
        if ((m.End - m.Start).TotalMinutes > config.MaxMinutes)
            return $"longer than {config.MaxMinutes} minutes";
        if (off.Kind == DayOffKind.HalfDay)
        {
            var midday = m.Start.Date + DayOffRules.Midday.ToTimeSpan();
            var during = off.Half switch
            {
                HalfOfDay.Am => m.Start < midday,
                HalfOfDay.Pm => m.End > midday,
                _ => off.Leave?.Entries.Any(e => !e.AllDay && e.Start < m.End && e.End > m.Start) ?? false,
            };
            if (during)
                return "during your half day off";
        }
        if (m.Start <= now)
            return "already started";
        return null;
    }

    /// <summary>Everyone at a meeting, by address (lower case): the organiser, the people invited (not rooms or
    /// resources), and you.</summary>
    public static HashSet<string> People(Config config, string me, Meeting m)
    {
        var people = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (m.OrganiserAddress.Length > 0 && !IsRoom(config, m.OrganiserAddress))
            people.Add(m.OrganiserAddress.ToLowerInvariant());
        foreach (var a in m.Attendees)
            if (a.Type != "resource" && a.Address.Length > 0 && !IsRoom(config, a.Address))
                people.Add(a.Address.ToLowerInvariant());
        if (me.Length > 0)
            people.Add(me.ToLowerInvariant());
        return people;
    }

    private static bool IsRoom(Config config, string address) => config.Rooms.Any(r => Same(r.Address, address));

    private static bool Same(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    /// <summary>Whether a meeting wants a room: the "No room" and "Room" categories first, then the first matching
    /// [[rules]] entry, else ask.</summary>
    public static (Want Want, Rule? Rule, string Why) Decide(Config config, Meeting m, IReadOnlySet<string> people)
    {
        if (m.Categories.Contains(config.NoRoomCategory, StringComparer.OrdinalIgnoreCase))
            return (Want.No, null, $"category \"{config.NoRoomCategory}\"");
        if (m.Categories.Contains(config.RoomCategory, StringComparer.OrdinalIgnoreCase))
            return (Want.Yes, null, $"category \"{config.RoomCategory}\"");
        foreach (var rule in config.Rules)
            if (Matches(rule, m, people))
                return (rule.Room, rule, rule.Label);
        return (Want.Ask, null, "no rule matched");
    }

    internal static bool Matches(Rule rule, Meeting m, IReadOnlySet<string> people)
    {
        if (rule.Subject is { } subject && !subject.IsMatch(m.Subject))
            return false;
        if (rule.Organiser is { } organisers && !organisers.Any(o =>
                Same(o, "me") ? m.IsOrganiser : Same(o, m.OrganiserAddress) || Same(o, m.OrganiserName)))
            return false;
        if (rule.OneToOne is { } oneToOne && (people.Count == 2) != oneToOne)
            return false;
        if (rule.MinAttendees is { } min && people.Count < min)
            return false;
        if (rule.MaxAttendees is { } max && people.Count > max)
            return false;
        if (rule.Response is { } responses && !responses.Contains(m.IsOrganiser ? "organizer" : m.Response))
            return false;
        if (rule.Series is { } series && !Same(series, m.SeriesMasterId ?? ""))
            return false;
        if (rule.Category is { } category && !m.Categories.Contains(category, StringComparer.OrdinalIgnoreCase))
            return false;
        if (rule.With is { } with && !with.Any(people.Contains))
            return false;
        return true;
    }

    /// <summary>The rooms whose schedules a judgement needs: every listed room, and every room named on a candidate.</summary>
    public static IReadOnlyList<string> RoomsToRead(Config config, IEnumerable<DayPlan> plans) =>
        config.Rooms.Select(r => r.Address)
            .Concat(plans.SelectMany(p => p.Candidates).SelectMany(c => c.Meeting.ListedRooms()))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    /// <summary>Where each candidate stands, and the room each that wants one would get. "Booked" is judged from the
    /// rooms' own schedules: a room counts only if its schedule shows this meeting (at its times, under its organiser's
    /// name or its subject), never just because your copy lists it. Back-to-back meetings (gaps of join_gap_minutes or
    /// less) get the same room where one is free for all of them. Pure.</summary>
    public static void Judge(Config config, List<DayPlan> plans, IReadOnlyDictionary<string, IReadOnlyList<ScheduleItem>?> schedules,
        State state, DateTime now)
    {
        var taken = new Dictionary<string, List<(DateTime Start, DateTime End)>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyList<ScheduleItem>? Schedule(string address) => schedules.TryGetValue(address, out var s) ? s : null;
        bool Free(Room room, Meeting m) =>
            Schedule(room.Address) is { } booked && RoomSchedule.Free(booked, m.Start, m.End)
            && !(taken.TryGetValue(room.Address, out var mine) && mine.Any(t => t.Start < m.End && t.End > m.Start));
        void Take(Room room, Meeting m)
        {
            if (!taken.TryGetValue(room.Address, out var list))
                taken[room.Address] = list = [];
            list.Add((m.Start, m.End));
        }
        bool Tried(Candidate c, Room room) => c.Attempts.Any(a => Same(a.Room, room.Address));
        string NameOf(string address, Meeting m) =>
            config.Rooms.FirstOrDefault(r => Same(r.Address, address))?.Name
            ?? (m.Attendees.FirstOrDefault(a => Same(a.Address, address))?.Name is { Length: > 0 } name ? name : address);

        foreach (var plan in plans.Where(p => p.Office))
        {
            foreach (var c in plan.Candidates)
            {
                var m = c.Meeting;
                c.Attempts = state.Attempts(m.Key);
                var listed = m.ListedRooms().ToList();
                var tried = c.Attempts.Select(a => a.Room).ToList();
                string? booked = null, bookedWhy = null;
                foreach (var address in listed.Concat(tried).Concat(config.Rooms.Select(r => r.Address)).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    if (Schedule(address) is not { } items)
                    {
                        if (listed.Contains(address, StringComparer.OrdinalIgnoreCase))
                            (booked, bookedWhy) = (address, "listed on the meeting; its schedule can't be read");
                        continue;
                    }
                    var (_, match) = RoomSchedule.Showing(items, m.Start, m.End, m.OrganiserName, m.OrganiserAddress, m.Subject);
                    if (match is ShowingMatch.Organiser or ShowingMatch.Subject
                        || (match == ShowingMatch.Hidden && (listed.Contains(address, StringComparer.OrdinalIgnoreCase) || tried.Contains(address, StringComparer.OrdinalIgnoreCase))))
                    {
                        (booked, bookedWhy) = (address, match switch
                        {
                            ShowingMatch.Organiser => "on the room's schedule under the organiser",
                            ShowingMatch.Subject => "on the room's schedule under its subject",
                            _ => "on the room's schedule at its times",
                        });
                        break;
                    }
                }
                if (booked is not null)
                {
                    c.Status = Status.Booked;
                    c.BookedRoom = NameOf(booked, m);
                    c.StatusWhy = bookedWhy!;
                    continue;
                }
                var done = c.Attempts.Where(a => a.State == "done").ToList();
                if (state.IsSkipped(m.Key))
                    (c.Status, c.StatusWhy) = (Status.Skipped, "you said no room (skip)");
                else if (c.Want == Want.No)
                    (c.Status, c.StatusWhy) = (Status.No, c.WantWhy);
                else if (c.Attempts.Any(a => a.State == "sending"))
                    (c.Status, c.StatusWhy) = (Status.MayHaveGone, $"a booking of {c.Attempts.Last(a => a.State == "sending").Name} may have gone: never sent again");
                else if (done.Count >= config.MaxAttempts)
                    (c.Status, c.StatusWhy) = (Status.GaveUp, $"{string.Join(", ", done.Select(a => a.Name))} asked; none shows it");
                else if (done.Count > 0 && now - done[^1].When < TimeSpan.FromMinutes(config.AnswerWaitMinutes))
                {
                    (c.Status, c.StatusWhy) = (Status.Waiting, $"{done[^1].Name} asked at {done[^1].When:HH:mm}; waiting for its answer");
                    if (config.Rooms.FirstOrDefault(r => Same(r.Address, done[^1].Room)) is { } waitingFor)
                        Take(waitingFor, m);
                }
                else
                    (c.Status, c.StatusWhy) = c.Want == Want.Yes ? (Status.Book, c.WantWhy) : (Status.Ask, c.WantWhy);
            }

            // Rooms: back-to-back runs first (one room for the run where one is free throughout), then one at a time.
            var ordered = plan.Candidates.OrderBy(c => c.Meeting.Start).ThenBy(c => c.Meeting.End).ToList();
            var runs = new List<List<Candidate>>();
            DateTime runEnd = default;
            foreach (var c in ordered.Where(c => c.Status is Status.Book or Status.Ask or Status.Booked or Status.Waiting))
            {
                if (runs.Count > 0 && c.Meeting.Start <= runEnd.AddMinutes(config.JoinGapMinutes))
                    runs[^1].Add(c);
                else
                    runs.Add([c]);
                runEnd = runs[^1].Count == 1 || c.Meeting.End > runEnd ? c.Meeting.End : runEnd;
            }
            foreach (var run in runs.Where(r => r.Count > 1))
            {
                var need = run.Where(c => c.Status is Status.Book or Status.Ask).ToList();
                if (need.Count == 0)
                    continue;
                var held = run.Where(c => c.Status is Status.Booked or Status.Waiting)
                    .Select(c => c.Status == Status.Booked ? config.Rooms.FirstOrDefault(r => r.Name == c.BookedRoom)
                        : config.Rooms.FirstOrDefault(r => c.Attempts.Count > 0 && Same(r.Address, c.Attempts[^1].Room)))
                    .OfType<Room>();
                // The room already held in the run, then the biggest meeting's order (its preferences lead), then the rest.
                var options = held.Concat(need.OrderByDescending(c => c.People).SelectMany(c => c.Order.Select(o => o.Room))).Distinct().ToList();
                var shared = options.FirstOrDefault(room => need.All(c => c.Order.Any(o => o.Room == room) && !Tried(c, room) && Free(room, c.Meeting)));
                if (shared is null)
                    continue;
                var span = $"{run[0].Meeting.Start:HH:mm}-{run.Max(c => c.Meeting.End):HH:mm}";
                foreach (var c in need)
                {
                    c.Choice = shared;
                    c.ChoiceWhy = $"one room for the back-to-back meetings {span}";
                    Take(shared, c.Meeting);
                }
            }
            foreach (var c in ordered.Where(c => c.Status is Status.Book or Status.Ask && c.Choice is null))
            {
                var pick = c.Order.FirstOrDefault(o => !Tried(c, o.Room) && Free(o.Room, c.Meeting));
                if (pick is null)
                {
                    var busy = c.Order.Count(o => Schedule(o.Room.Address) is not null && !Free(o.Room, c.Meeting));
                    var unread = c.Order.Count(o => Schedule(o.Room.Address) is null);
                    var small = config.Rooms.Count - c.Order.Count;
                    var parts = new List<string>();
                    if (busy > 0)
                        parts.Add($"{busy} busy");
                    if (c.Attempts.Count > 0)
                        parts.Add($"{c.Attempts.Count} already asked");
                    if (small > 0)
                        parts.Add($"{small} too small");
                    if (unread > 0)
                        parts.Add($"{unread} unreadable");
                    c.StatusWhy = config.Rooms.Count == 0 ? "no [[rooms]] in the settings" : $"no room free that seats {c.People}: {string.Join(", ", parts)}";
                    if (c.Status == Status.Book)
                        c.Status = Status.NoRoom;
                    continue;
                }
                c.Choice = pick.Room;
                c.ChoiceWhy = pick.Why.StartsWith("preference", StringComparison.Ordinal)
                    ? $"{pick.Why} ({(c.Rule?.Rooms?.Any(r => RoomList.Find([pick.Room], r) is not null) == true ? $"rule {c.Rule.Index}'s rooms" : "default_rooms")})"
                    : pick.Why;
                if (c.Attempts.Count > 0)
                    c.ChoiceWhy += "; after no answer from " + string.Join(", ", c.Attempts.Select(a => a.Name));
                Take(pick.Room, c.Meeting);
            }
        }
    }

    private static List<(DateTime Start, DateTime End)> Merge(IEnumerable<(DateTime Start, DateTime End)> spans)
    {
        var merged = new List<(DateTime Start, DateTime End)>();
        foreach (var s in spans.OrderBy(s => s.Start))
        {
            if (merged.Count > 0 && s.Start <= merged[^1].End)
                merged[^1] = (merged[^1].Start, s.End > merged[^1].End ? s.End : merged[^1].End);
            else
                merged.Add(s);
        }
        return merged;
    }

    /// <summary>One line for a candidate: its time, subject, verdict, and the room.</summary>
    public static string Line(Candidate c, bool act)
    {
        var m = c.Meeting;
        var what = c.Status switch
        {
            Status.Booked => $"booked: {c.BookedRoom} ({c.StatusWhy})",
            Status.Book => $"{(act ? "book" : "would book")} {c.Choice!.Name} by {(c.How == "add" ? "adding it to your meeting" : "forwarding to the room")} ({c.ChoiceWhy}; {c.WantWhy})",
            Status.Ask => c.Choice is null ? $"ask; {c.StatusWhy} ({c.WantWhy})" :$"ask: {c.Choice.Name} ({c.ChoiceWhy}; {c.WantWhy})",
            Status.No => $"no room ({c.StatusWhy})",
            Status.NoRoom => $"wants a room, none free ({c.StatusWhy}; {c.WantWhy})",
            Status.Skipped => $"no room ({c.StatusWhy})",
            Status.MayHaveGone => $"CHECK: {c.StatusWhy}",
            Status.Waiting => $"waiting: {c.StatusWhy}",
            Status.GaveUp => $"no room: {c.StatusWhy}",
            _ => "",
        };
        return $"    {m.Start:HH:mm}-{m.End:HH:mm}  {Clip(m.Subject, 40),-40}  {c.People,3}  {what}";
    }

    public static string Clip(string text, int width) => text.Length <= width ? text : text[..(width - 1)] + "…";

    public static string DayName(DateOnly day) => day.ToString("ddd dd MMM", En);
}
