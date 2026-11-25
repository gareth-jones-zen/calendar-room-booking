using ActivityJournal;
using M365Graph;
using static RoomBooking.Tests.Make;

namespace RoomBooking.Tests;

/// <summary>0.2.0: each room asked for, and each meeting skipped, is a line in the shared activity journal
/// (background-tools-lib's ActivityJournal); never plan, a preview, ask mode or a refused request, and writing it never
/// fails the run. Every person, room and meeting here is made up.</summary>
public sealed class JournalTests
{
    private static string Run(FakeCalendar calendar, Config config, out int code, params string[] args)
    {
        var output = new StringWriter();
        code = Cli.Run(["--config", config.Path, .. args], Host(calendar, output));
        return output.ToString();
    }

    private static IReadOnlyList<JournalEntry> Lines(string folder)
    {
        var path = Path.Combine(folder, "journal.jsonl");
        return File.Exists(path) ? new Journal(path).Read() : [];
    }

    [Fact]
    public void Act_mode_journals_each_booking_with_its_rule_how_and_answer()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"\njournal_file = \"journal.jsonl\"");
        var mine = Meeting("My 1:1", At(Thu, 9), 30, "me", "sam");
        calendar.Items.Add(mine);
        calendar.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar.Accepting.Add("booth3.9@example.com");

        Run(calendar, config, out _, "plan");
        Assert.Empty(Lines(folder));

        Run(calendar, config, out var code, "run");
        Assert.Equal(0, code);
        var lines = Lines(folder);
        Assert.Equal(2, lines.Count);
        Assert.All(lines, l =>
        {
            Assert.Equal("calendar-room-booking", l.Tool);
            Assert.Equal("book", l.What);
            Assert.Equal(By.Auto, l.By);
            Assert.Equal(Sensitivity.Private, l.Sensitivity);
            Assert.Null(l.Undo);
            Assert.Null(l.Outcome);
        });
        Assert.Equal("My 1:1, Thu 26 Nov 09:00", lines[0].Subject);
        Assert.Equal("rule 1 (1:1s)", lines[0].Rule);
        Assert.Equal("Booth 3.9: added to your meeting (everyone on it got an update); it accepted", lines[0].Why);
        Assert.Equal(mine.Id, lines[0].Item);
        Assert.Equal("Booth 3.9: the invitation forwarded to the room; it accepted", lines[1].Why);

        // Nothing more to book: nothing more journalled.
        Run(calendar, config, out _, "run");
        Assert.Equal(2, Lines(folder).Count);
    }

    [Fact]
    public void Ask_mode_and_a_preview_write_nothing_and_book_and_skip_by_hand_do()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"ask\"\njournal_file = \"journal.jsonl\"");
        var peer = Meeting("Peer 1:1", At(Thu, 11), 30, "alex");
        var chat = Meeting("Catch-up", At(Thu, 14), 30, "alex", "sam");
        calendar.Items.Add(peer);
        calendar.Items.Add(chat);
        Run(calendar, config, out _, "run");
        Run(calendar, config, out _, "book", "2026-11-26 11:00");
        Assert.Empty(Lines(folder));

        Run(calendar, config, out var code, "book", "2026-11-26 11:00", "--yes");
        Assert.Equal(0, code);
        var booked = Assert.Single(Lines(folder));
        Assert.Equal(By.Hand, booked.By);
        Assert.Equal("by hand", booked.Rule);
        Assert.Equal("Room 3.1: the invitation forwarded to the room; no answer yet", booked.Why);

        Run(calendar, config, out _, "skip", "2026-11-26 14:00");
        var skipped = Lines(folder)[1];
        Assert.Equal("skip", skipped.What);
        Assert.Equal(By.Hand, skipped.By);
        Assert.Equal("Catch-up, Thu 26 Nov 14:00", skipped.Subject);
        Assert.Equal($"calendar-room-booking book {chat.Id}", skipped.Undo);
        Assert.Equal(Sensitivity.Private, skipped.Sensitivity);
    }

    [Fact]
    public void A_lost_answer_is_unconfirmed_and_a_refused_request_writes_nothing()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"\njournal_file = \"journal.jsonl\"");
        calendar.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar.WriteFails = new GraphDeliveryUnknownException("POST", "/me/events/x/forward", System.Net.HttpStatusCode.ServiceUnavailable, "busy");
        Run(calendar, config, out var code, "run");
        Assert.Equal(1, code);
        var line = Assert.Single(Lines(folder));
        Assert.Equal(JournalEntry.Unconfirmed, line.Outcome);
        Assert.EndsWith("the answer was lost: it may have gone", line.Why);

        var (config2, calendar2, folder2) = Setup(extra: "mode = \"act\"\njournal_file = \"journal.jsonl\"");
        calendar2.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar2.WriteFails = new GraphRequestException(System.Net.HttpStatusCode.Forbidden, "POST", "/me/events/x/forward", "no");
        Run(calendar2, config2, out _, "run");
        Assert.Empty(Lines(folder2));
    }

    [Fact]
    public void Without_journal_file_nothing_is_written_and_a_bad_journal_never_fails_the_run()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"");
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        Run(calendar, config, out var code, "run");
        Assert.Equal(0, code);
        Assert.Empty(Directory.GetFiles(folder, "*.jsonl", SearchOption.AllDirectories));

        var (config2, calendar2, folder2) = Setup(extra: "mode = \"act\"\njournal_file = \"journal.jsonl\"");
        Directory.CreateDirectory(Path.Combine(folder2, "journal.jsonl"));   // a folder where the file should be
        calendar2.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        Run(calendar2, config2, out code, "run");
        Assert.Equal(0, code);
        Assert.Single(calendar2.Writes);
        Assert.Contains("journal: couldn't write", File.ReadAllText(Path.Combine(folder2, "log.txt")));
    }

    [Fact]
    public void Journal_file_comes_from_a_base_with_its_path_variables()
    {
        var folder = Path.Combine(Path.GetTempPath(), "room-booking-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "base-config.toml"), $"journal_file = '%SHARED%/journal.jsonl'\n[paths]\nSHARED = '{folder}\\shared'\n");
        var path = Path.Combine(folder, "settings.toml");
        File.WriteAllText(path, "base = 'base-config.toml'\nstate_file = \"state.json\"\nlog_file = \"log.txt\"\n");
        Assert.Equal(Path.Combine(folder, "shared", "journal.jsonl"), Path.GetFullPath(Config.Load(path, _ => null).JournalFile!));
    }
}
