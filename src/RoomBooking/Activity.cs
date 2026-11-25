using System.Globalization;
using ActivityJournal;
using MeetingRooms;

namespace RoomBooking;

/// <summary>
/// The shared activity journal (0.2.0): one line for each room asked for and each meeting skipped, written through
/// background-tools-lib's ActivityJournal to journal_file (by convention %EA_LOCAL%\journal.jsonl), for a "while you
/// were away" view. Only what was sent: never `plan`, a preview, ask mode's review items, or a request Microsoft
/// refused; a request whose answer was lost is marked unconfirmed. Meeting subjects make every line "private". Writing
/// it never fails the run: a problem is logged.
/// <para><c>what</c> is <c>book</c> (by auto for a run in act mode, hand for <c>book --yes</c>) or <c>skip</c> (by hand).
/// A booking has no <c>undo</c>: there's no command that gives a room back (an attendee can't cancel a room's booking,
/// and on your own meeting the room is taken off in Outlook).</para>
/// </summary>
internal static class Activity
{
    public const string Tool = "calendar-room-booking";

    private static readonly CultureInfo En = CultureInfo.InvariantCulture;

    private static string Subject(Meeting m) => $"{m.Subject}, {m.Start.ToString("ddd d MMM HH:mm", En)}";

    /// <summary>A line for a room asked for: how, and the room's answer as far as it's known. Never throws.</summary>
    public static void Booked(Config config, DateTime now, Action<string> log, Meeting m, Room room, string by, string? rule,
        Outcome outcome)
    {
        var how = m.IsOrganiser ? "added to your meeting (everyone on it got an update)" : "the invitation forwarded to the room";
        var answer = outcome switch
        {
            Outcome.Accepted => "it accepted",
            Outcome.Declined => "it declined",
            Outcome.MayHaveGone => "the answer was lost: it may have gone",
            _ => "no answer yet",
        };
        Write(config, now, log, new JournalEntry
        {
            Tool = Tool,
            What = "book",
            By = by,
            Subject = Subject(m),
            Rule = by == By.Hand ? "by hand" : rule,
            Why = $"{room.Name}: {how}; {answer}",
            Item = m.Id,
            Sensitivity = Sensitivity.Private,
            Outcome = outcome == Outcome.MayHaveGone ? JournalEntry.Unconfirmed : null,
        });
    }

    /// <summary>A line for a meeting `skip` marked as wanting no room. Never throws.</summary>
    public static void Skipped(Config config, DateTime now, Action<string> log, Meeting m) =>
        Write(config, now, log, new JournalEntry
        {
            Tool = Tool,
            What = "skip",
            By = By.Hand,
            Subject = Subject(m),
            Rule = "by hand",
            Why = "no room for this meeting at this time",
            Item = m.Id,
            // `book` changes your mind: a preview until --yes, which books it and clears the skip.
            Undo = $"calendar-room-booking book {m.Id}",
            Sensitivity = Sensitivity.Private,
        });

    private static void Write(Config config, DateTime now, Action<string> log, JournalEntry entry)
    {
        if (config.JournalFile is not { } path)
            return;
        try
        {
            new Journal(path, clock: () => new DateTimeOffset(now)).TryWrite(entry, log);
        }
        catch (Exception error) when (error is not OutOfMemoryException)
        {
            log($"journal: couldn't write to {path}: {error.GetType().Name}: {error.Message}");
        }
    }
}
