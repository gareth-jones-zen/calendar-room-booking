using System.Globalization;
using System.Text.Json.Nodes;
using M365Graph;
using MeetingRooms;
using WinBackground;

namespace RoomBooking;

/// <summary>One stretch of a day in your work plan (Graph's workLocationType: office, remote, timeOff), local time.</summary>
public sealed record Segment(DateTime Start, DateTime End, string Location);

/// <summary>Someone on a meeting: address (as given), name, type (required, optional, resource) and their response.</summary>
public sealed record Attendee(string Address, string Name, string Type, string Response);

/// <summary>A meeting in your calendar as read, local time. <see cref="Locations"/> are the event's locations as Graph
/// gave them (kept to write back when a room is added to a meeting you organise).</summary>
public sealed record Meeting(string Id, string ICalUId, string Subject, DateTime Start, DateTime End)
{
    public bool AllDay { get; init; }
    public string ShowAs { get; init; } = "busy";
    public IReadOnlyList<string> Categories { get; init; } = [];
    public bool IsOrganiser { get; init; }
    public string OrganiserAddress { get; init; } = "";
    public string OrganiserName { get; init; } = "";
    public IReadOnlyList<Attendee> Attendees { get; init; } = [];
    public bool Cancelled { get; init; }
    /// <summary>Your response: organizer, accepted, tentativelyAccepted, declined, notResponded, none.</summary>
    public string Response { get; init; } = "none";
    public bool Online { get; init; }
    public string? SeriesMasterId { get; init; }
    public string Location { get; init; } = "";
    public IReadOnlyList<JsonObject> Locations { get; init; } = [];

    /// <summary>The meeting at this time: what the send-once record is keyed by, so a moved meeting is a new one.</summary>
    public string Key => $"{ICalUId}|{Start.ToString("yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture)}";

    /// <summary>Rooms named on this copy: resource attendees and locations with an address.</summary>
    public IEnumerable<string> ListedRooms() =>
        Attendees.Where(a => a.Type == "resource" && a.Address.Length > 0).Select(a => a.Address)
            .Concat(Locations.Select(l => Calendar.Text(l, "locationEmailAddress") ?? "").Where(a => a.Contains('@')))
            .Distinct(StringComparer.OrdinalIgnoreCase);
}

/// <summary>What a run reads and writes. GraphCalendar is the real one; tests give a fake.</summary>
public interface ICalendar
{
    /// <summary>Your own address.</summary>
    string Me();

    /// <summary>Your work plan from first to last (inclusive).</summary>
    IReadOnlyList<Segment> WorkPlan(DateOnly first, DateOnly last);

    /// <summary>Every event from first to last (inclusive).</summary>
    IReadOnlyList<Meeting> Events(DateOnly first, DateOnly last);

    /// <summary>One event, read again (null if it's gone).</summary>
    Meeting? Event(string id);

    /// <summary>Rooms' schedules (free/busy), by address; null for one that can't be read.</summary>
    IReadOnlyDictionary<string, IReadOnlyList<ScheduleItem>?> RoomSchedules(IReadOnlyList<string> addresses, DateTime from, DateTime to);

    /// <summary>A meeting you organise: the room added as a resource attendee and a location (an update goes to every
    /// attendee). <paramref name="drop"/>: rooms this tool added before that declined, taken off. Sent once.</summary>
    void AddRoom(Meeting meeting, Room room, IReadOnlyCollection<string> drop);

    /// <summary>A meeting you attend: the invitation forwarded to the room, as Outlook's "Add room" does. Sent once.</summary>
    void Forward(Meeting meeting, Room room, string comment);
}

/// <summary>Your Outlook calendar, work plan and rooms' free/busy, through Microsoft Graph (Calendars.ReadWrite).</summary>
public sealed class Calendar(GraphClient graph, TimeZoneInfo zone, string? loginHint) : ICalendar
{
    public static readonly string[] Scopes = ["Calendars.ReadWrite"];

    private const string Select = "id,iCalUId,subject,start,end,isAllDay,showAs,categories,isOrganizer,organizer,attendees,isCancelled,"
                                  + "responseStatus,isOnlineMeeting,onlineMeeting,seriesMasterId,location,locations";

    private Dictionary<string, string> Headers => new() { ["Prefer"] = $"outlook.timezone=\"{zone.Id}\"" };

    private string Utc(DateOnly day) =>
        TimeZoneInfo.ConvertTimeToUtc(day.ToDateTime(TimeOnly.MinValue), zone).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

    private string? me;

    public string Me()
    {
        if (me is not null)
            return me;
        if (loginHint is { Length: > 0 } hint && hint.Contains('@'))
            return me = hint;
        var reply = graph.RequestAsync("/me?$select=mail,userPrincipalName", HttpMethod.Get, null, ["User.Read"]).GetAwaiter().GetResult();
        return me = Text(reply, "mail") ?? Text(reply, "userPrincipalName") ?? "";
    }

    /// <summary>Graph refuses more than 31 days of the work plan at once; 30 keeps a margin for a clock change.</summary>
    public const int WorkPlanChunkDays = 30;

    public IReadOnlyList<Segment> WorkPlan(DateOnly first, DateOnly last)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var segments = new List<Segment>();
        for (var from = first; from <= last; from = from.AddDays(WorkPlanChunkDays))
        {
            var end = from.AddDays(WorkPlanChunkDays - 1);
            var to = end < last ? end : last;
            var url = $"/me/settings/workHoursAndLocations/occurrencesView(startDateTime='{Utc(from)}',endDateTime='{Utc(to.AddDays(1))}')";
            foreach (var o in graph.PagesAsync(url, Scopes, Headers).ToBlockingEnumerable().OfType<JsonObject>())
            {
                var segment = new Segment(Time(o["start"]), Time(o["end"]), Text(o, "workLocationType") ?? "unknown");
                if (seen.Add(Text(o, "id") ?? $"{segment.Start:O}|{segment.End:O}|{segment.Location}"))
                    segments.Add(segment);
            }
        }
        return segments;
    }

    public IReadOnlyList<Meeting> Events(DateOnly first, DateOnly last)
    {
        var url = $"/me/calendarView?startDateTime={Utc(first)}&endDateTime={Utc(last.AddDays(1))}&$select={Select}&$top=250";
        return graph.PagesAsync(url, Scopes, Headers).ToBlockingEnumerable().OfType<JsonObject>().Select(Read).ToList();
    }

    public Meeting? Event(string id)
    {
        try
        {
            return Read(graph.RequestAsync($"/me/events/{Uri.EscapeDataString(id)}?$select={Select}", HttpMethod.Get, null, Scopes, Headers)
                .GetAwaiter().GetResult());
        }
        catch (GraphRequestException error) when ((int)error.Status is 404 or 410)
        {
            return null;
        }
    }

    public IReadOnlyDictionary<string, IReadOnlyList<ScheduleItem>?> RoomSchedules(IReadOnlyList<string> addresses, DateTime from, DateTime to) =>
        new GraphRoomSchedules(graph, Scopes, zone).Read(addresses, from, to);

    /// <summary>A PATCH of the meeting's attendees and locations: sent once (the library's default for a PATCH), since
    /// it sends every attendee an update.</summary>
    public void AddRoom(Meeting meeting, Room room, IReadOnlyCollection<string> drop) =>
        graph.RequestAsync($"/me/events/{Uri.EscapeDataString(meeting.Id)}", HttpMethod.Patch, AddRoomBody(meeting, room, drop), Scopes, Headers)
            .GetAwaiter().GetResult();

    /// <summary>Graph's event: forward (Calendars.Read is enough; 202 and no body): sent once (a POST).</summary>
    public void Forward(Meeting meeting, Room room, string comment) =>
        graph.RequestAsync($"/me/events/{Uri.EscapeDataString(meeting.Id)}/forward", HttpMethod.Post, ForwardBody(room, comment), Scopes, Headers)
            .GetAwaiter().GetResult();

    /// <summary>The meeting's attendees (each as address, name and type: responses aren't written) with the room added
    /// as a resource, and its locations with the room added as a conference room; rooms in <paramref name="drop"/> are
    /// taken off both. A Teams meeting's "Microsoft Teams Meeting" location is kept.</summary>
    public static JsonObject AddRoomBody(Meeting meeting, Room room, IReadOnlyCollection<string> drop)
    {
        bool Dropped(string? address) => address is not null && (drop.Contains(address, StringComparer.OrdinalIgnoreCase)
                                                                  || string.Equals(address, room.Address, StringComparison.OrdinalIgnoreCase));
        var attendees = new JsonArray();
        foreach (var a in meeting.Attendees.Where(a => a.Address.Length > 0 && !Dropped(a.Address)))
            attendees.Add((JsonNode)new JsonObject
            {
                ["emailAddress"] = new JsonObject { ["address"] = a.Address, ["name"] = a.Name },
                ["type"] = a.Type,
            });
        attendees.Add((JsonNode)new JsonObject
        {
            ["emailAddress"] = new JsonObject { ["address"] = room.Address, ["name"] = room.Name },
            ["type"] = "resource",
        });
        var locations = new JsonArray();
        var existing = meeting.Locations.Count > 0 ? meeting.Locations
            : meeting.Location.Length > 0 ? [new JsonObject { ["displayName"] = meeting.Location }] : [];
        foreach (var l in existing)
        {
            var address = Text(l, "locationEmailAddress");
            if (Dropped(address))
                continue;
            var copy = new JsonObject();
            foreach (var key in new[] { "displayName", "locationEmailAddress", "locationType", "locationUri" })
                if (Text(l, key) is { Length: > 0 } value)
                    copy[key] = value;
            if (copy.Count > 0)
                locations.Add((JsonNode)copy);
        }
        locations.Add((JsonNode)new JsonObject
        {
            ["displayName"] = room.Name, ["locationEmailAddress"] = room.Address, ["locationType"] = "conferenceRoom",
        });
        return new JsonObject { ["attendees"] = attendees, ["locations"] = locations };
    }

    public static JsonObject ForwardBody(Room room, string comment) => new()
    {
        ["ToRecipients"] = new JsonArray(new JsonObject
        {
            ["EmailAddress"] = new JsonObject { ["Address"] = room.Address, ["Name"] = room.Name },
        }),
        ["Comment"] = comment,
    };

    internal static Meeting Read(JsonObject o)
    {
        var organiser = (o["organizer"] as JsonObject)?["emailAddress"] as JsonObject;
        var attendees = (o["attendees"] as JsonArray)?.OfType<JsonObject>().Select(a =>
        {
            var email = a["emailAddress"] as JsonObject;
            return new Attendee(email is null ? "" : Text(email, "address") ?? "", email is null ? "" : Text(email, "name") ?? "",
                Text(a, "type") ?? "required", (a["status"] as JsonObject) is { } s ? Text(s, "response") ?? "none" : "none");
        }).ToList() ?? [];
        var online = Flag(o, "isOnlineMeeting")
                     || (o["onlineMeeting"] as JsonObject) is { } om && Text(om, "joinUrl") is { Length: > 0 };
        return new Meeting(Text(o, "id") ?? "", Text(o, "iCalUId") ?? Text(o, "id") ?? "", Text(o, "subject") ?? "", Time(o["start"]), Time(o["end"]))
        {
            AllDay = Flag(o, "isAllDay"),
            ShowAs = Text(o, "showAs") ?? "busy",
            Categories = (o["categories"] as JsonArray)?.Select(c => c?.GetValue<string>() ?? "").Where(c => c.Length > 0).ToList() ?? [],
            IsOrganiser = Flag(o, "isOrganizer"),
            OrganiserAddress = organiser is null ? "" : Text(organiser, "address") ?? "",
            OrganiserName = organiser is null ? "" : Text(organiser, "name") ?? "",
            Attendees = attendees,
            Cancelled = Flag(o, "isCancelled"),
            Response = (o["responseStatus"] as JsonObject) is { } r ? Text(r, "response") ?? "none" : "none",
            Online = online,
            SeriesMasterId = Text(o, "seriesMasterId"),
            Location = (o["location"] as JsonObject) is { } loc ? Text(loc, "displayName") ?? "" : "",
            Locations = (o["locations"] as JsonArray)?.OfType<JsonObject>().Select(l => (JsonObject)l.DeepClone()).ToList() ?? [],
        };
    }

    internal static DateTime Time(JsonNode? node)
    {
        var text = (node as JsonObject) is { } o ? Text(o, "dateTime") : null;
        return text is null ? default : DateTime.Parse(text[..Math.Min(19, text.Length)], CultureInfo.InvariantCulture, DateTimeStyles.None);
    }

    internal static string? Text(JsonObject o, string key) =>
        o[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;

    private static bool Flag(JsonObject o, string key) =>
        o[key] is JsonValue v && v.TryGetValue<bool>(out var b) && b;

    public static GraphClient Client(Config config) => new(new GraphClientOptions
    {
        CachePath = config.MsalCache ?? M365Graph.Graph.DefaultCachePath,
        Tenant = config.Tenant,
        LoginHint = config.LoginHint,
        Scopes = Scopes,
        SignInHint = "At a console, run: calendar-room-booking sign-in",
        Log = Logger.Create(config.LogFile),
    });
}
