using System.Text.RegularExpressions;
using CalendarDays;
using MeetingRooms;
using PyCompat;
using ReviewQueue;
using Tomlyn;
using Tomlyn.Model;
using WinBackground;

namespace RoomBooking;

/// <summary>Settings that can't be used, or a request that can't be met: the command line reports the message and
/// exits 2.</summary>
public class UsageException(string message) : Exception(message);

/// <summary>Whether a meeting wants a room.</summary>
public enum Want { Yes, No, Ask }

/// <summary>A meeting that's never a candidate: its subject matches, or it carries the category.</summary>
public sealed record SkipRule(Regex? Subject, string? Category, string Why);

/// <summary>One of the [[rules]], tried in order; the first whose every given test passes decides. Index is its place
/// (1-based), for plan's "rule 3".</summary>
public sealed record Rule(int Index, Want Room, string Why)
{
    public Regex? Subject { get; init; }
    /// <summary>Addresses or names (any case); "me" is a meeting you organise.</summary>
    public IReadOnlyList<string>? Organiser { get; init; }
    public bool? OneToOne { get; init; }
    public int? MinAttendees { get; init; }
    public int? MaxAttendees { get; init; }
    /// <summary>Your response: organizer, accepted, tentativelyAccepted.</summary>
    public IReadOnlyList<string>? Response { get; init; }
    public string? Series { get; init; }
    public string? Category { get; init; }
    /// <summary>Any of these addresses is at the meeting.</summary>
    public IReadOnlyList<string>? With { get; init; }
    /// <summary>The rooms to try first for this rule, in order (names or addresses); null: default_rooms.</summary>
    public IReadOnlyList<string>? Rooms { get; init; }

    public string Label => $"rule {Index} ({Why})";
}

/// <summary>The settings file. See the README and <see cref="Template"/>. Read strictly: an unknown key, or a value of
/// the wrong type, is an error that names it.</summary>
public sealed class Config
{
    public const string EnvConfig = "ROOM_BOOKING_CONFIG";

    public static readonly string[] Responses = ["organizer", "accepted", "tentativelyAccepted", "declined", "notResponded", "none"];

    // Explicit "\n" (a raw literal would take the .cs file's line endings), written with CRLF.
    public const string Template =
        "# calendar-room-booking settings. See the README.\n" +
        "# base = \"base-config.toml\"        # settings shared with other tools\n" +
        "# login_hint = \"you@example.com\"\n" +
        "# tenant = \"organizations\"          # or your organisation's tenant id\n" +
        "# time_zone = \"GMT Standard Time\"   # default: this PC's\n" +
        "\n" +
        "mode = \"ask\"                       # ask: books nothing, asks in the review queue; act: books the \"yes\" meetings\n" +
        "days_ahead = 14\n" +
        "default_rooms = []                 # rooms to try first when a rule names none\n" +
        "room_floors = []                   # then rooms on these floors, in this order\n" +
        "\n" +
        "[[rooms]]\n" +
        "name = \"Room 1\"\n" +
        "address = \"room1@example.com\"\n" +
        "# floor = \"1\"\n" +
        "# capacity = 6\n" +
        "\n" +
        "[[skip]]                           # never a candidate\n" +
        "subject = '^(lunch|travel|focus)'\n" +
        "why = \"your own blocks\"\n" +
        "\n" +
        "[[rules]]                          # first match decides\n" +
        "one_to_one = true\n" +
        "room = \"yes\"\n" +
        "why = \"1:1s\"\n" +
        "\n" +
        "[[rules]]\n" +
        "room = \"ask\"\n" +
        "why = \"everything else\"\n" +
        "\n" +
        "[leave]                            # what counts as a day off (background-tools-lib's shared keys)\n" +
        "\n" +
        "[schedule]\n" +
        "every_minutes = 15\n" +
        "hours = \"07:00-19:00\"\n" +
        "days = \"Mon-Fri\"\n";

    public static readonly string[] KnownKeys =
    [
        "login_hint", "tenant", "msal_cache", "log_file", "state_file", "time_zone", "enabled", "mode", "days_ahead",
        "max_minutes", "join_gap_minutes", "max_bookings", "max_attempts", "answer_wait_minutes", "notify",
        "room_category", "no_room_category", "responses", "forward_comment", "default_rooms", "room_floors", "rooms",
        "skip", "rules", "leave", "schedule", "review", "journal_file",
    ];

    private static readonly string[] RuleKeys =
        ["subject", "organiser", "one_to_one", "min_attendees", "max_attendees", "response", "series", "category", "with", "room", "rooms", "why"];

    private static readonly string[] SkipKeys = ["subject", "category", "why"];

    /// <summary>The task's schedule when the settings don't give one: every 15 minutes, 07:00 to 19:00, Monday to
    /// Friday, and at logon.</summary>
    public static readonly ScheduleSettings ScheduleDefaults = new()
    {
        EveryMinutes = 15, AtLogon = true, AfterUnlockMinutes = 0,
        Hours = ScheduleSettings.ParseHours("07:00-19:00"), Days = ScheduleSettings.ParseDays("Mon-Fri"),
    };

    public required string Path { get; init; }
    public string? LoginHint { get; init; }
    public string Tenant { get; init; } = "organizations";
    public string? MsalCache { get; init; }
    public required string LogFile { get; init; }
    public required string StateFile { get; init; }
    /// <summary>The shared activity journal (0.2.0; journal_file, usually from the base): one line for each room asked for
    /// and each meeting skipped. Null writes none.</summary>
    public string? JournalFile { get; init; }
    public required TimeZoneInfo TimeZone { get; init; }
    public bool Enabled { get; init; } = true;
    /// <summary>"ask": no writes, the meetings it would book go to the review queue; "act": it books them.</summary>
    public string Mode { get; init; } = "ask";
    public bool Act => Mode == "act";
    public int DaysAhead { get; init; } = 14;
    public int MaxMinutes { get; init; } = 240;
    public int JoinGapMinutes { get; init; } = 10;
    public int MaxBookings { get; init; } = 10;
    /// <summary>Rooms tried for one meeting at one time before it's raised as "no room": the first and one more after
    /// a decline.</summary>
    public int MaxAttempts { get; init; } = 2;
    /// <summary>How long a room is given to answer before its silence counts as a decline.</summary>
    public int AnswerWaitMinutes { get; init; } = 10;
    public bool Notify { get; init; } = true;
    public string RoomCategory { get; init; } = "Room";
    public string NoRoomCategory { get; init; } = "No room";
    /// <summary>Your answers that make a meeting a candidate, besides organising it.</summary>
    public IReadOnlyList<string> CandidateResponses { get; init; } = ["accepted", "tentativelyAccepted"];
    public string ForwardComment { get; init; } = "";
    public IReadOnlyList<string> DefaultRooms { get; init; } = [];
    public IReadOnlyList<string> RoomFloors { get; init; } = [];
    public IReadOnlyList<Room> Rooms { get; init; } = [];
    public IReadOnlyList<SkipRule> Skip { get; init; } = [];
    public IReadOnlyList<Rule> Rules { get; init; } = [];
    public LeaveSettings Leave { get; init; } = new();
    public ScheduleSettings Schedule { get; init; } = ScheduleDefaults;
    public ReviewSettings? Review { get; init; }

    public static string? Environment(string name) =>
        string.IsNullOrEmpty(name) || name.Contains('\0') ? null : System.Environment.GetEnvironmentVariable(name);

    public static string Local(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var local = env("LOCALAPPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(local) ? PyPath.Home(env) : local, "calendar-room-booking"));
    }

    public static string DefaultPath(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var appData = env("APPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(appData) ? PyPath.Home(env) : appData, "calendar-room-booking", "config.toml"));
    }

    /// <summary>The settings as text, read with their base (tests use this directly).</summary>
    public static Config Load(string? path = null, Func<string, string?>? env = null)
    {
        env ??= Environment;
        var given = !string.IsNullOrEmpty(path) ? path : env(EnvConfig) is { Length: > 0 } fromEnv ? fromEnv : DefaultPath(env);
        var full = PyPath.Resolve(given);
        if (!File.Exists(full))
            throw new UsageException($"No settings file at {full}. Run `calendar-room-booking config` to create one.");
        TomlTable raw;
        try
        {
            var text = PyStr.ReadText(full);
            if (text.StartsWith('﻿'))
                text = text[1..];
            var read = BaseSettings.Read(text, full, env, KnownKeys);
            raw = read.Raw;
            env = read.Env;
        }
        catch (TomlException error)
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message.TrimEnd());
        }
        catch (Exception error) when (error is not (IOException or UnauthorizedAccessException or UsageException))
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }
        foreach (var key in raw.Keys)
            if (!KnownKeys.Contains(key))
                throw new UsageException($"{full}: unknown setting '{key}' (known: {string.Join(", ", KnownKeys)}).");

        var folder = System.IO.Path.GetDirectoryName(full)!;
        var expandWith = env;
        string PathSetting(string value)
        {
            var expanded = PyPath.ExpandVars(PyPath.ExpandUser(value, expandWith), expandWith);
            return PyPath.IsAbsolute(expanded) ? PyPath.Pure(expanded) : PyPath.ResolveAgainst(folder, expanded);
        }
        UsageException Bad(string message) => new($"{full}: {message}");
        string? Text(TomlTable table, string key, string where) =>
            table.TryGetValue(key, out var v) ? v as string is { Length: > 0 } s ? s : throw Bad($"{where} should be text, and not empty.") : null;
        bool Flag(string key, bool fallback) =>
            raw.TryGetValue(key, out var v) ? v as bool? ?? throw Bad($"{key} should be true or false.") : fallback;
        int Whole(TomlTable table, string key, string where, int fallback, int min, int max) =>
            table.TryGetValue(key, out var v) ? v is long n && n >= min && n <= max ? (int)n : throw Bad($"{where} should be a whole number from {min} to {max}.") : fallback;
        IReadOnlyList<string>? Strings(TomlTable table, string key, string where, bool single = false) =>
            !table.TryGetValue(key, out var v) ? null
            : single && v is string one ? [one]
            : v is TomlArray a && a.All(x => x is string) ? a.Select(x => (string)x!).ToList()
            : throw Bad($"{where} should be a list of text{(single ? ", or text" : "")}.");
        Regex Pattern(string text, string where)
        {
            try
            {
                return new Regex(text, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            }
            catch (ArgumentException error)
            {
                throw Bad($"{where} isn't a valid pattern: {error.Message}");
            }
        }
        IEnumerable<TomlTable> Tables(string key)
        {
            if (!raw.TryGetValue(key, out var v))
                return [];
            return v is TomlTableArray array ? array : throw Bad($"{key} should be a list of [[{key}]] tables.");
        }

        TimeZoneInfo zone;
        var zoneName = Text(raw, "time_zone", "time_zone");
        try
        {
            zone = zoneName is null ? TimeZoneInfo.Local : TimeZoneInfo.FindSystemTimeZoneById(zoneName);
        }
        catch (Exception error) when (error is TimeZoneNotFoundException or InvalidTimeZoneException)
        {
            throw Bad($"time_zone '{zoneName}' isn't a time zone this PC knows (a Windows name, such as \"GMT Standard Time\").");
        }

        var mode = Text(raw, "mode", "mode") ?? "ask";
        if (mode is not ("ask" or "act"))
            throw Bad("mode should be \"ask\" (book nothing; ask in the review queue) or \"act\" (book the \"yes\" meetings).");

        IReadOnlyList<Room> rooms;
        try
        {
            rooms = RoomList.Parse(raw.TryGetValue("rooms", out var rv) ? rv : null, full);
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message);
        }
        void KnownRooms(IReadOnlyList<string> names, string where)
        {
            foreach (var name in names)
                if (RoomList.Find(rooms, name) is null)
                    throw Bad($"{where}: '{name}' isn't one of the [[rooms]] (by name or address).");
        }

        var defaultRooms = Strings(raw, "default_rooms", "default_rooms") ?? [];
        KnownRooms(defaultRooms, "default_rooms");

        var responses = Strings(raw, "responses", "responses") ?? ["accepted", "tentativelyAccepted"];
        foreach (var r in responses)
            if (!Responses.Contains(r))
                throw Bad($"responses: '{r}' isn't one of {string.Join(", ", Responses)}.");

        var skip = new List<SkipRule>();
        var n = 0;
        foreach (var t in Tables("skip"))
        {
            n++;
            var where = $"[[skip]] {n}";
            foreach (var key in t.Keys)
                if (!SkipKeys.Contains(key))
                    throw Bad($"{where} has an unknown setting '{key}' (known: {string.Join(", ", SkipKeys)}).");
            var subject = Text(t, "subject", $"{where} subject");
            var category = Text(t, "category", $"{where} category");
            if (subject is null && category is null)
                throw Bad($"{where} needs a subject or a category.");
            skip.Add(new SkipRule(subject is null ? null : Pattern(subject, $"{where} subject"), category, Text(t, "why", $"{where} why") ?? "a [[skip]] rule"));
        }

        var rules = new List<Rule>();
        n = 0;
        foreach (var t in Tables("rules"))
        {
            n++;
            var where = $"[[rules]] {n}";
            foreach (var key in t.Keys)
                if (!RuleKeys.Contains(key))
                    throw Bad($"{where} has an unknown setting '{key}' (known: {string.Join(", ", RuleKeys)}).");
            var room = Text(t, "room", $"{where} room") ?? throw Bad($"{where} needs room = \"yes\", \"no\" or \"ask\".");
            var want = room switch
            {
                "yes" => Want.Yes,
                "no" => Want.No,
                "ask" => Want.Ask,
                _ => throw Bad($"{where} room should be \"yes\", \"no\" or \"ask\"."),
            };
            var ruleRooms = Strings(t, "rooms", $"{where} rooms");
            if (ruleRooms is not null)
                KnownRooms(ruleRooms, $"{where} rooms");
            var response = Strings(t, "response", $"{where} response", single: true);
            foreach (var r in response ?? [])
                if (!Responses.Contains(r))
                    throw Bad($"{where} response: '{r}' isn't one of {string.Join(", ", Responses)}.");
            bool? oneToOne = t.TryGetValue("one_to_one", out var o) ? o as bool? ?? throw Bad($"{where} one_to_one should be true or false.") : null;
            var subject = Text(t, "subject", $"{where} subject");
            rules.Add(new Rule(n, want, Text(t, "why", $"{where} why") ?? $"room = \"{room}\"")
            {
                Subject = subject is null ? null : Pattern(subject, $"{where} subject"),
                Organiser = Strings(t, "organiser", $"{where} organiser", single: true),
                OneToOne = oneToOne,
                MinAttendees = t.ContainsKey("min_attendees") ? Whole(t, "min_attendees", $"{where} min_attendees", 0, 1, 10000) : null,
                MaxAttendees = t.ContainsKey("max_attendees") ? Whole(t, "max_attendees", $"{where} max_attendees", 0, 1, 10000) : null,
                Response = response,
                Series = Text(t, "series", $"{where} series"),
                Category = Text(t, "category", $"{where} category"),
                With = Strings(t, "with", $"{where} with", single: true),
                Rooms = ruleRooms,
            });
        }

        LeaveSettings leave;
        ScheduleSettings schedule;
        ReviewSettings? review = null;
        try
        {
            leave = LeaveSettings.Parse(raw.TryGetValue("leave", out var lv) ? lv as TomlTable ?? throw Bad("leave should be a [leave] section.") : null, full);
            schedule = ScheduleSettings.Parse(raw.TryGetValue("schedule", out var sv) ? sv as TomlTable ?? throw Bad("schedule should be a [schedule] section.") : null,
                ScheduleDefaults, TaskKind.Job, $"{full}: [schedule]");
            if (raw.TryGetValue("review", out var review_))
                review = ReviewSettings.Read(review_, full, PathSetting);
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message);
        }

        return new Config
        {
            Path = full,
            LoginHint = Text(raw, "login_hint", "login_hint"),
            Tenant = Text(raw, "tenant", "tenant") ?? "organizations",
            MsalCache = Text(raw, "msal_cache", "msal_cache") is { } cache ? PathSetting(cache) : null,
            LogFile = Text(raw, "log_file", "log_file") is { } log ? PathSetting(log) : PyPath.Pure(System.IO.Path.Combine(Local(env), "calendar-room-booking.log")),
            JournalFile = Text(raw, "journal_file", "journal_file") is { } journal ? PathSetting(journal) : null,
            StateFile = Text(raw, "state_file", "state_file") is { } state ? PathSetting(state) : PyPath.Pure(System.IO.Path.Combine(Local(env), "state.json")),
            TimeZone = zone,
            Enabled = Flag("enabled", true),
            Mode = mode,
            DaysAhead = Whole(raw, "days_ahead", "days_ahead", 14, 0, 62),
            MaxMinutes = Whole(raw, "max_minutes", "max_minutes", 240, 5, 1440),
            JoinGapMinutes = Whole(raw, "join_gap_minutes", "join_gap_minutes", 10, 0, 120),
            MaxBookings = Whole(raw, "max_bookings", "max_bookings", 10, 0, 1000),
            MaxAttempts = Whole(raw, "max_attempts", "max_attempts", 2, 1, 10),
            AnswerWaitMinutes = Whole(raw, "answer_wait_minutes", "answer_wait_minutes", 10, 0, 1440),
            Notify = Flag("notify", true),
            RoomCategory = Text(raw, "room_category", "room_category") ?? "Room",
            NoRoomCategory = Text(raw, "no_room_category", "no_room_category") ?? "No room",
            CandidateResponses = responses,
            ForwardComment = raw.TryGetValue("forward_comment", out var fc) ? fc as string ?? throw Bad("forward_comment should be text.") : "",
            DefaultRooms = defaultRooms,
            RoomFloors = Strings(raw, "room_floors", "room_floors") ?? [],
            Rooms = rooms,
            Skip = skip,
            Rules = rules,
            Leave = leave,
            Schedule = schedule,
            Review = review,
        };
    }

    public static string CreateDefault(string? path, Func<string, string?>? env = null)
    {
        var full = PyPath.Resolve(!string.IsNullOrEmpty(path) ? path : DefaultPath(env));
        if (!File.Exists(full))
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(full)!);
            File.WriteAllText(full, Template.Replace("\n", "\r\n"), new System.Text.UTF8Encoding(false));
        }
        return full;
    }
}
