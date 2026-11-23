using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace RoomBooking;

/// <summary>One room asked for one meeting at one time: the room, how ("add" or "forward"), "sending" (written before
/// the write: if it's still "sending", the answer was lost and it's never sent again) or "done", when, and the room's
/// answer as last read ("accepted", "declined", or null: not seen yet).</summary>
public sealed record Attempt(string Room, string Name, string How, string State, DateTime When, string? Answer);

/// <summary>The state file (state_file): what this tool has asked rooms for (the send-once record), the meetings you
/// said no to, the review items it has raised, and whether runs are failing. Runtime state, outside any repo.
/// <code>
/// {"bookings": {"&lt;iCalUId&gt;|&lt;start&gt;": {"subject", "start", "end", "event_id", "attempts": [...], "raised": "..."}},
///  "skipped": {"&lt;iCalUId&gt;|&lt;start&gt;": {"subject", "start", "when"}},
///  "raised": {"2026-11-26": {"hash", "id"}},
///  "failures": 1, "notified": "..."}
/// </code></summary>
public sealed class State
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    public JsonObject Data { get; }

    public State(JsonObject? data = null) => Data = data ?? [];

    public JsonObject Bookings => Section("bookings");
    public JsonObject Skipped => Section("skipped");
    public JsonObject Raised => Section("raised");

    private JsonObject Section(string name)
    {
        if (Data[name] is JsonObject o)
            return o;
        var made = new JsonObject();
        Data[name] = made;
        return made;
    }

    public static State Load(string path)
    {
        try
        {
            return new State(File.Exists(path) ? JsonNode.Parse(File.ReadAllText(path)) as JsonObject : null);
        }
        catch (Exception error) when (error is IOException or JsonException)
        {
            // A state file that can't be read loses the record of bookings sent: refuse rather than risk sending twice.
            throw new UsageException($"The state file {path} can't be read ({error.Message}). It holds the record of rooms already asked "
                                     + "for, so nothing is booked until it's fixed or moved aside.");
        }
    }

    public void Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        var temp = path + ".tmp";
        File.WriteAllText(temp, PyCompat.PyJson.Dump(Data, 2, false));
        File.Move(temp, path, true);
    }

    public IReadOnlyList<Attempt> Attempts(string key) =>
        (Bookings[key] as JsonObject)?["attempts"] is JsonArray list
            ? list.OfType<JsonObject>().Select(a => new Attempt(Str(a, "room"), Str(a, "name"), Str(a, "how"), Str(a, "state"),
                DateTime.TryParse(Str(a, "when"), En, DateTimeStyles.None, out var when) ? when : default,
                a["answer"] is JsonValue v && v.TryGetValue<string>(out var answer) ? answer : null)).ToList()
            : [];

    /// <summary>The booking entry for a meeting at its time, made if it isn't there.</summary>
    public JsonObject Entry(Meeting m)
    {
        if (Bookings[m.Key] is JsonObject entry)
            return entry;
        entry = new JsonObject
        {
            ["subject"] = m.Subject,
            ["start"] = m.Start.ToString("yyyy-MM-ddTHH:mm", En),
            ["end"] = m.End.ToString("yyyy-MM-ddTHH:mm", En),
            ["event_id"] = m.Id,
            ["organiser"] = m.IsOrganiser,
            ["attempts"] = new JsonArray(),
        };
        Bookings[m.Key] = entry;
        return entry;
    }

    public bool IsSkipped(string key) => Skipped.ContainsKey(key);

    public void Skip(Meeting m, DateTime now) => Skipped[m.Key] = new JsonObject
    {
        ["subject"] = m.Subject, ["start"] = m.Start.ToString("yyyy-MM-ddTHH:mm", En), ["when"] = now.ToString("yyyy-MM-ddTHH:mm:ss", En),
    };

    /// <summary>Forget what's past: bookings and skips that started before yesterday, and raised days before today.</summary>
    public void Prune(DateOnly today)
    {
        var cutoff = today.AddDays(-1).ToString("yyyy-MM-dd", En);
        foreach (var section in new[] { Bookings, Skipped })
            foreach (var key in section.Where(kv => (Str(kv.Value as JsonObject, "start") is { Length: >= 10 } s ? s[..10] : "") is var d && string.CompareOrdinal(d, cutoff) < 0)
                         .Select(kv => kv.Key).ToList())
                section.Remove(key);
        var todayText = today.ToString("yyyy-MM-dd", En);
        foreach (var key in Raised.Select(kv => kv.Key).Where(k => string.CompareOrdinal(k, todayText) < 0).ToList())
            Raised.Remove(key);
    }

    internal static string Str(JsonObject? o, string key) =>
        o?[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : "";
}
