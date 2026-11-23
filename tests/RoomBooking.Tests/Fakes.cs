using System.Text.Json.Nodes;
using MeetingRooms;

namespace RoomBooking.Tests;

/// <summary>A calendar in memory: a work plan, meetings, rooms' schedules, and the writes made.</summary>
internal sealed class FakeCalendar : ICalendar
{
    public const string MeAddress = "pat.example@example.com";
    public List<Segment> Plan { get; } = [];
    public List<Meeting> Items { get; } = [];
    public Dictionary<string, List<ScheduleItem>> Schedules { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> Writes { get; } = [];
    public List<JsonObject> Bodies { get; } = [];

    /// <summary>Rooms that take a booking (it shows on their schedule straight away); the others say nothing.</summary>
    public HashSet<string> Accepting { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The next write throws this, once.</summary>
    public Exception? WriteFails { get; set; }

    public string Me() => MeAddress;

    public IReadOnlyList<Segment> WorkPlan(DateOnly first, DateOnly last) => Plan;

    public IReadOnlyList<Meeting> Events(DateOnly first, DateOnly last) =>
        Items.Where(e => e.End > first.ToDateTime(TimeOnly.MinValue) && e.Start < last.AddDays(1).ToDateTime(TimeOnly.MinValue)).ToList();

    public Meeting? Event(string id) => Items.FirstOrDefault(e => e.Id == id);

    public IReadOnlyDictionary<string, IReadOnlyList<ScheduleItem>?> RoomSchedules(IReadOnlyList<string> addresses, DateTime from, DateTime to) =>
        addresses.Distinct(StringComparer.OrdinalIgnoreCase).ToDictionary(a => a,
            a => Schedules.TryGetValue(a, out var s) ? (IReadOnlyList<ScheduleItem>?)s.ToList() : null, StringComparer.OrdinalIgnoreCase);

    public void AddRoom(Meeting meeting, Room room, IReadOnlyCollection<string> drop)
    {
        Writes.Add($"add {room.Name} to {meeting.Subject}");
        Bodies.Add(Calendar.AddRoomBody(meeting, room, drop));
        Landed(meeting, room);
    }

    public void Forward(Meeting meeting, Room room, string comment)
    {
        Writes.Add($"forward {meeting.Subject} to {room.Name}");
        Bodies.Add(Calendar.ForwardBody(room, comment));
        Landed(meeting, room);
    }

    private void Landed(Meeting meeting, Room room)
    {
        if (WriteFails is { } fail)
        {
            WriteFails = null;
            throw fail;
        }
        if (Accepting.Contains(room.Address))
            Schedules[room.Address].Add(new ScheduleItem(meeting.Start, meeting.End, "busy", meeting.OrganiserName));
    }
}

internal static class Make
{
    public static readonly DateTime Now = new(2026, 11, 23, 9, 0, 0);   // a Monday
    public static readonly DateOnly Thu = new(2026, 11, 26);

    public static DateTime At(DateOnly day, int hour, int minute = 0) => day.ToDateTime(new TimeOnly(hour, minute));

    private static int next;

    /// <summary>A Teams meeting: organised by <paramref name="organiser"/> ("me" for you), with these other people.</summary>
    public static Meeting Meeting(string subject, DateTime start, int minutes, string organiser = "me", params string[] others)
    {
        var mine = organiser == "me";
        var organiserAddress = mine ? FakeCalendar.MeAddress : $"{organiser}@example.com";
        var attendees = others.Select(o => new Attendee($"{o}@example.com", Title(o), "required", "accepted")).ToList();
        if (!mine)
            attendees.Add(new Attendee(FakeCalendar.MeAddress, "Pat Example", "required", "accepted"));
        var id = $"ev{++next}";
        return new Meeting(id, $"uid-{id}", subject, start, start.AddMinutes(minutes))
        {
            IsOrganiser = mine,
            OrganiserAddress = organiserAddress,
            OrganiserName = mine ? "Pat Example" : Title(organiser),
            Attendees = attendees,
            Response = mine ? "organizer" : "accepted",
            Online = true,
            Locations = [new JsonObject { ["displayName"] = "Microsoft Teams Meeting", ["locationType"] = "default" }],
            Location = "Microsoft Teams Meeting",
        };
    }

    private static string Title(string who) => string.Concat(who[..1].ToUpperInvariant(), who.AsSpan(1)) + " Example";

    public const string Settings = """
        mode = "ask"
        days_ahead = 14
        default_rooms = ["Room 3.1"]
        room_floors = ["3", "2"]
        max_bookings = 5

        [[rooms]]
        name = "Room 3.1"
        address = "room3.1@example.com"
        floor = "3"
        capacity = 8
        [[rooms]]
        name = "Room 3.2"
        address = "room3.2@example.com"
        floor = "3"
        [[rooms]]
        name = "Booth 3.9"
        address = "booth3.9@example.com"
        floor = "3"
        capacity = 2
        [[rooms]]
        name = "Room 2.1"
        address = "room2.1@example.com"
        floor = "2"

        [[skip]]
        subject = '^(lunch|focus)'
        why = "own blocks"

        [[rules]]
        one_to_one = true
        organiser = ["me", "boss@example.com"]
        room = "yes"
        rooms = ["Booth 3.9"]
        why = "1:1s"
        [[rules]]
        min_attendees = 20
        room = "no"
        why = "big calls"
        [[rules]]
        organiser = "me"
        room = "yes"
        why = "calls I chair"
        [[rules]]
        room = "ask"
        why = "everything else"

        [leave]

        [review]
        file = "review.json"
        source = "calendar-room-booking"
        """;

    /// <summary>A settings file in a new temporary folder (with its review queue and state beside it), and a fake
    /// calendar with an office day on Thursday 26 November, 08:00-17:00, and every room empty.</summary>
    public static (Config Config, FakeCalendar Calendar, string Folder) Setup(string? settings = null, string extra = "")
    {
        var folder = Path.Combine(Path.GetTempPath(), "room-booking-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "settings.toml");
        // extra replaces the mode line: other top-level keys, such as mode = "act".
        File.WriteAllText(path, "state_file = \"state.json\"\nlog_file = \"log.txt\"\n" + (settings ?? Settings).Replace("mode = \"ask\"", extra.Length > 0 ? extra : "mode = \"ask\""));
        var config = Config.Load(path, _ => null);
        var calendar = new FakeCalendar();
        calendar.Plan.Add(new Segment(At(Thu, 8), At(Thu, 17), "office"));
        calendar.Plan.Add(new Segment(At(Thu.AddDays(-1), 8), At(Thu.AddDays(-1), 17), "remote"));
        foreach (var room in config.Rooms)
            calendar.Schedules[room.Address] = [];
        return (config, calendar, folder);
    }

    public static CliHost Host(FakeCalendar calendar, StringWriter output, DateTime? now = null) => new()
    {
        Out = output, Err = output, Env = _ => null, Now = () => now ?? Now, Calendar = _ => calendar, Sleep = _ => { },
        Notify = (_, _) => { }, Lock = false,
    };
}
