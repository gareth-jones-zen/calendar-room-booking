using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using M365Graph;
using MeetingRooms;
using ReviewQueue;

namespace RoomBooking;

/// <summary>How asking a room went.</summary>
public enum Outcome
{
    /// <summary>The room's schedule shows the meeting.</summary>
    Accepted,
    /// <summary>The room declined (read from your meeting's attendees, for a meeting you organise).</summary>
    Declined,
    /// <summary>Sent; the room hadn't answered by the last read. The next run judges it from the room's schedule.</summary>
    NoAnswer,
    /// <summary>The answer to the write was lost: it may have gone. Recorded, never sent again.</summary>
    MayHaveGone,
    /// <summary>Microsoft refused it (a 4xx): nothing was sent.</summary>
    Refused,
}

/// <summary>Asking a room for a meeting, once, and the review items a run raises.</summary>
public static class Booker
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    /// <summary>How many times the room's answer is read after asking, and the wait between reads.</summary>
    public const int Reads = 4;
    public static readonly TimeSpan ReadWait = TimeSpan.FromSeconds(5);

    /// <summary>Asks <paramref name="room"/> for <paramref name="m"/>: added to the meeting if you organise it,
    /// forwarded to the room if not. <b>Never twice:</b> the attempt is recorded as "sending" in the state file before
    /// the write and "done" after, so an answer lost in between leaves "sending", and nothing for that meeting at that
    /// time is ever sent again. Then the room's answer is read back a few times, a few seconds apart. What was sent goes
    /// in the shared activity journal (<see cref="Activity"/>), <paramref name="by"/> auto or hand, with the rule that
    /// chose it.</summary>
    public static (Outcome Outcome, string Message) Book(Config config, ICalendar calendar, State state, Meeting m, Room room,
        Action<string> log, Action<TimeSpan> sleep, DateTime now, string by, string? rule)
    {
        var before = state.Attempts(m.Key);
        if (before.FirstOrDefault(a => a.State == "sending") is { } lost)
            return (Outcome.MayHaveGone, $"{lost.Name} may already have been asked for this meeting ({lost.When:dd/MM HH:mm}; the answer was lost): never sent again.");
        var how = m.IsOrganiser ? "add" : "forward";
        var entry = state.Entry(m);
        var attempt = new JsonObject
        {
            ["room"] = room.Address, ["name"] = room.Name, ["how"] = how, ["state"] = "sending",
            ["when"] = now.ToString("yyyy-MM-ddTHH:mm:ss", En),
        };
        ((JsonArray)entry["attempts"]!).Add((JsonNode)attempt);
        state.Save(config.StateFile);   // before the write: if its answer is lost, it's never sent again
        var what = $"{m.Start:ddd dd/MM HH:mm} \"{m.Subject}\": {room.Name}";
        try
        {
            if (m.IsOrganiser)
                calendar.AddRoom(m, room, before.Select(a => a.Room).ToList());
            else
                calendar.Forward(m, room, config.ForwardComment);
        }
        catch (GraphRequestException error) when ((int)error.Status is >= 400 and < 500)
        {
            ((JsonArray)entry["attempts"]!).Remove(attempt);
            state.Save(config.StateFile);
            log($"refused ({(int)error.Status}): {what}: {Head(error.Message)}");
            return (Outcome.Refused, $"Microsoft refused it ({(int)error.Status}): nothing was sent. {Head(error.Message)}");
        }
        catch (SignInNeededException)
        {
            ((JsonArray)entry["attempts"]!).Remove(attempt);
            state.Save(config.StateFile);
            throw;
        }
        catch (Exception error)
        {
            // A 5xx, a timeout or a dropped connection after sending (GraphDeliveryUnknownException), or anything else
            // once the request may have left: it may have gone, so it stays "sending" and is never sent again.
            log($"couldn't confirm: {what} may have gone ({error.GetType().Name}: {Head(error.Message)}); never sent again");
            Activity.Booked(config, now, log, m, room, by, rule, Outcome.MayHaveGone);
            return (Outcome.MayHaveGone, $"The answer was lost ({Head(error.Message)}): it may have gone. Check the room in Outlook; it's never sent again from here.");
        }
        attempt["state"] = "done";
        state.Save(config.StateFile);
        log($"{(how == "add" ? "added to the meeting" : "forwarded to the room")}: {what}");

        var answer = ReadAnswer(calendar, m, room, sleep);
        if (answer is not null)
        {
            attempt["answer"] = answer;
            state.Save(config.StateFile);
        }
        if (!m.IsOrganiser)
            try
            {
                var copy = calendar.Event(m.Id);
                log($"after the forward, your copy {(copy?.ListedRooms().Contains(room.Address, StringComparer.OrdinalIgnoreCase) == true ? "lists" : "doesn't list")} {room.Name}");
            }
            catch (Exception error) when (error is not SignInNeededException)
            {
                log($"couldn't read your copy back: {Head(error.Message)}");
            }
        log($"room's answer: {what}: {answer ?? "none yet"}");
        var result = answer switch
        {
            "accepted" => (Outcome.Accepted, $"{room.Name} accepted: it's booked."),
            "declined" => (Outcome.Declined, $"{room.Name} declined."),
            _ => (Outcome.NoAnswer, $"{room.Name} hasn't answered yet: the next run checks its schedule."),
        };
        Activity.Booked(config, now, log, m, room, by, rule, result.Item1);
        return result;
    }

    /// <summary>The room's answer, read back up to <see cref="Reads"/> times: "accepted" when its schedule shows the
    /// meeting, "declined" when your meeting's attendees say so (a meeting you organise), null while neither. Only
    /// reads; the write is never repeated, whatever comes back.</summary>
    internal static string? ReadAnswer(ICalendar calendar, Meeting m, Room room, Action<TimeSpan> sleep)
    {
        for (var n = 0; n < Reads; n++)
        {
            if (n > 0)
                sleep(ReadWait);
            try
            {
                var schedule = calendar.RoomSchedules([room.Address], m.Start.Date, m.Start.Date.AddDays(1));
                if (schedule.TryGetValue(room.Address, out var items) && items is not null
                    && RoomSchedule.Showing(items, m.Start, m.End, m.OrganiserName, m.OrganiserAddress, m.Subject).Match != ShowingMatch.None)
                    return "accepted";
                if (m.IsOrganiser && calendar.Event(m.Id)?.Attendees.FirstOrDefault(a => string.Equals(a.Address, room.Address, StringComparison.OrdinalIgnoreCase)) is { } back)
                {
                    if (back.Response == "declined")
                        return "declined";
                    if (back.Response == "accepted")
                        return "accepted";
                }
            }
            catch (Exception error) when (error is not SignInNeededException)
            {
                // a read that failed: try again, or leave it to the next run
            }
        }
        return null;
    }

    private static string Head(string text) => text.Length <= 200 ? text : text[..200] + "…";

    /// <summary>The line an ask item gives a meeting.</summary>
    internal static string ItemLine(Candidate c, bool act) =>
        $"{c.Meeting.Start:HH:mm}-{c.Meeting.End:HH:mm} {c.Meeting.Subject} ({c.People} people; {c.WantWhy}): " + c.Status switch
        {
            Status.Book when !act => $"book {c.Choice!.Name}?",
            Status.Ask => c.Choice is null ? "no room free" : $"book {c.Choice.Name}?",
            Status.NoRoom => "wants a room; none free",
            _ => c.Status.ToString(),
        };

    /// <summary>The review items for a run: one per office day listing the meetings to decide (in ask mode, those the
    /// rules would book as well as those they ask about), raised again only when the list changes; one item per meeting
    /// whose booking may have gone; one per meeting no room took. Each raised once (the state file remembers).</summary>
    public static int Raise(Config config, List<DayPlan> plans, State state, DateTime now, Action<string> log)
    {
        if (config.Review is null)
            return 0;
        var queue = config.Review.Open(() => now);
        var source = config.Review.Source;
        var raised = 0;
        foreach (var plan in plans)
        {
            var day = plan.Day.ToString("yyyy-MM-dd", En);
            var asks = plan.Candidates.Where(c => c.Status is Status.Ask or Status.NoRoom || (!config.Act && c.Status == Status.Book))
                .OrderBy(c => c.Meeting.Start).ToList();
            var lines = asks.Select(c => ItemLine(c, config.Act)).ToList();
            var hash = Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", lines))))[..16];
            var before = state.Raised[day] as JsonObject;
            if (before is not null && State.Str(before, "hash") == hash)
                continue;
            if (before is not null)
            {
                queue.Done([State.Str(before, "id")]);
                state.Raised.Remove(day);
            }
            if (lines.Count == 0)
                continue;
            var fields = new JsonObject
            {
                ["title"] = $"{Planner.DayName(plan.Day)}: {lines.Count} meeting{(lines.Count == 1 ? "" : "s")} without a room to decide",
                ["day"] = day,
                ["mode"] = config.Mode,
                ["meetings"] = new JsonArray([.. lines.Select(l => (JsonNode)JsonValue.Create(l))]),
                ["book"] = "calendar-room-booking book \"<day> <HH:MM>\" [--room \"<room>\"] --yes; skip \"<day> <HH:MM>\" for no",
            };
            var id = queue.Add(source, "ask", fields);
            state.Raised[day] = new JsonObject { ["hash"] = hash, ["id"] = id };
            log($"asked: {fields["title"]}");
            raised++;
        }
        foreach (var c in plans.SelectMany(p => p.Candidates).Where(c => c.Status is Status.MayHaveGone or Status.GaveUp))
        {
            var kind = c.Status == Status.MayHaveGone ? "may-have-gone" : "no-room";
            var entry = state.Entry(c.Meeting);
            if (State.Str(entry, "raised") == kind)
                continue;
            var m = c.Meeting;
            queue.Add(source, kind, new JsonObject
            {
                ["title"] = kind == "may-have-gone"
                    ? $"{m.Start:ddd dd MMM HH:mm} {m.Subject}: a room booking may have gone; check the room in Outlook"
                    : $"{m.Start:ddd dd MMM HH:mm} {m.Subject}: no room took it ({c.StatusWhy})",
                ["day"] = m.Start.ToString("yyyy-MM-dd", En),
                ["start"] = m.Start.ToString("HH:mm", En),
                ["subject"] = m.Subject,
                ["why"] = c.StatusWhy,
            });
            entry["raised"] = kind;
            log($"raised {kind}: {m.Start:dd/MM HH:mm} {m.Subject}");
            raised++;
        }
        return raised;
    }
}
