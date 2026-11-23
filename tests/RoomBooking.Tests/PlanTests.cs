using M365Graph;
using MeetingRooms;
using static RoomBooking.Tests.Make;

namespace RoomBooking.Tests;

public sealed class PlanTests
{
    private static List<DayPlan> Judge(Config config, FakeCalendar calendar, State? state = null, DateTime? now = null)
    {
        var first = DateOnly.FromDateTime(now ?? Now);
        var last = first.AddDays(config.DaysAhead);
        var plans = Planner.Days(config, FakeCalendar.MeAddress, now ?? Now, calendar.Plan, calendar.Items, first, last);
        Planner.Judge(config, plans, calendar.RoomSchedules(Planner.RoomsToRead(config, plans), At(first, 0), At(last.AddDays(1), 0)),
            state ?? new State(), now ?? Now);
        return plans;
    }

    private static Candidate Only(List<DayPlan> plans, string subject) =>
        plans.SelectMany(p => p.Candidates).Single(c => c.Meeting.Subject == subject);

    [Fact]
    public void Only_office_days_have_candidates()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("Remote day 1:1", At(Thu.AddDays(-1), 10), 30, "me", "sam"));
        calendar.Items.Add(Meeting("Office day 1:1", At(Thu, 10), 30, "me", "sam"));
        var plans = Judge(config, calendar);
        Assert.Equal([Thu], plans.Where(p => p.Office).Select(p => p.Day));
        Assert.Equal("work plan: remote", plans.Single(p => p.Day == Thu.AddDays(-1)).Why);
        Assert.Equal(["Office day 1:1"], plans.SelectMany(p => p.Candidates).Select(c => c.Meeting.Subject));
    }

    [Fact]
    public void Leave_means_no_office_day()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(new Meeting("hol", "hol", "Annual leave", At(Thu, 0), At(Thu.AddDays(1), 0)) { AllDay = true, IsOrganiser = true, Response = "organizer", ShowAs = "oof" });
        calendar.Items.Add(Meeting("1:1", At(Thu, 10), 30, "me", "sam"));
        var plans = Judge(config, calendar);
        Assert.False(plans.Single(p => p.Day == Thu).Office);
        Assert.StartsWith("leave", plans.Single(p => p.Day == Thu).Why);
    }

    [Fact]
    public void Meetings_that_arent_candidates_say_why()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("Own block", At(Thu, 9), 30, "me"));
        calendar.Items.Add(Meeting("Lunch with the team", At(Thu, 12), 30, "me", "sam"));
        calendar.Items.Add(Meeting("In person", At(Thu, 13), 30, "me", "sam") with { Online = false });
        calendar.Items.Add(Meeting("Evening", At(Thu, 17), 30, "me", "sam"));
        calendar.Items.Add(Meeting("Unanswered", At(Thu, 14), 30, "alex", "sam") with { Response = "notResponded" });
        calendar.Items.Add(Meeting("Shown free", At(Thu, 15), 30, "alex", "sam") with { ShowAs = "free" });
        calendar.Items.Add(Meeting("Workshop", At(Thu, 8), 300, "alex", "sam"));
        var others = Judge(config, calendar).Single(p => p.Day == Thu).Others.ToDictionary(o => o.Meeting.Subject, o => o.Why);
        Assert.Equal("your own entry (no one else)", others["Own block"]);
        Assert.Equal("skip: own blocks", others["Lunch with the team"]);
        Assert.Equal("not a Teams meeting", others["In person"]);
        Assert.Equal("outside office hours", others["Evening"]);
        Assert.Equal("not accepted (notResponded)", others["Unanswered"]);
        Assert.Equal("shown as free", others["Shown free"]);
        Assert.Equal("longer than 240 minutes", others["Workshop"]);
    }

    [Fact]
    public void Rules_first_match_and_categories_first()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        calendar.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar.Items.Add(Meeting("Peer 1:1", At(Thu, 11), 30, "alex"));
        calendar.Items.Add(Meeting("All hands", At(Thu, 12), 30, "alex", [.. Enumerable.Range(1, 25).Select(n => $"p{n}")]));
        calendar.Items.Add(Meeting("My review", At(Thu, 13), 30, "me", "sam", "jo"));
        calendar.Items.Add(Meeting("Tagged", At(Thu, 14), 30, "alex", "sam") with { Categories = ["Room"] });
        calendar.Items.Add(Meeting("Not this one", At(Thu, 15), 30, "me", "sam") with { Categories = ["No room"] });
        var plans = Judge(config, calendar);
        Assert.Equal(Want.Yes, Only(plans, "My 1:1").Want);
        Assert.Equal("rule 1 (1:1s)", Only(plans, "Boss 1:1").WantWhy);
        Assert.Equal(Want.Ask, Only(plans, "Peer 1:1").Want);
        Assert.Equal(Status.No, Only(plans, "All hands").Status);
        Assert.Equal("rule 3 (calls I chair)", Only(plans, "My review").WantWhy);
        Assert.Equal(Want.Yes, Only(plans, "Tagged").Want);
        Assert.Equal(Status.No, Only(plans, "Not this one").Status);
        Assert.Equal("category \"No room\"", Only(plans, "Not this one").StatusWhy);
    }

    [Fact]
    public void Rooms_follow_the_rules_preferences_then_default_rooms_then_floors()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        calendar.Items.Add(Meeting("My review", At(Thu, 13), 30, "me", "sam", "jo"));
        calendar.Schedules["room3.1@example.com"].Add(new ScheduleItem(At(Thu, 12), At(Thu, 16), "busy", "Someone Else"));
        var plans = Judge(config, calendar);
        Assert.Equal("Booth 3.9", Only(plans, "My 1:1").Choice!.Name);
        Assert.Equal("preference 1 (rule 1's rooms)", Only(plans, "My 1:1").ChoiceWhy);
        // Room 3.1 (default_rooms) is taken: the third floor next; the booth seats 2, too few for 3.
        Assert.Equal("Room 3.2", Only(plans, "My review").Choice!.Name);
        Assert.Equal("floor 3, size not known", Only(plans, "My review").ChoiceWhy);
    }

    [Fact]
    public void Back_to_back_meetings_share_a_room_free_for_all_of_them()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("First", At(Thu, 10), 30, "me", "sam", "jo"));
        calendar.Items.Add(Meeting("Second", At(Thu, 10, 35), 25, "me", "sam", "jo"));
        calendar.Items.Add(Meeting("Third", At(Thu, 11), 60, "me", "sam", "jo"));
        calendar.Schedules["room3.1@example.com"].Add(new ScheduleItem(At(Thu, 11, 30), At(Thu, 12), "busy", "Someone Else"));
        var plans = Judge(config, calendar);
        // Room 3.1 is free for the first two but not the third: one room for all three is Room 3.2.
        Assert.All(new[] { "First", "Second", "Third" }, s => Assert.Equal("Room 3.2", Only(plans, s).Choice!.Name));
        Assert.StartsWith("one room for the back-to-back meetings 10:00-12:00", Only(plans, "First").ChoiceWhy);
    }

    [Fact]
    public void A_gap_longer_than_join_gap_is_not_back_to_back()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("First", At(Thu, 10), 30, "me", "sam", "jo"));
        calendar.Items.Add(Meeting("Later", At(Thu, 11), 60, "me", "sam", "jo"));
        calendar.Schedules["room3.1@example.com"].Add(new ScheduleItem(At(Thu, 11, 30), At(Thu, 12), "busy", "Someone Else"));
        var plans = Judge(config, calendar);
        Assert.Equal("Room 3.1", Only(plans, "First").Choice!.Name);
        Assert.Equal("Room 3.2", Only(plans, "Later").Choice!.Name);
    }

    [Fact]
    public void In_a_run_the_biggest_meetings_preference_leads_and_a_rule_falls_back_to_default_rooms()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 10), 30, "me", "sam"));
        calendar.Items.Add(Meeting("My review", At(Thu, 10, 30), 30, "me", "sam", "jo"));
        calendar.Items.Add(Meeting("Later 1:1", At(Thu, 15), 30, "me", "sam"));
        calendar.Schedules["booth3.9@example.com"].Add(new ScheduleItem(At(Thu, 15), At(Thu, 15, 30), "busy", "Someone Else"));
        var plans = Judge(config, calendar);
        // The booth seats 2, too few for the review: the review's first choice, Room 3.1 (default_rooms), for both.
        Assert.Equal("Room 3.1", Only(plans, "My 1:1").Choice!.Name);
        Assert.Equal("Room 3.1", Only(plans, "My review").Choice!.Name);
        // The booth is taken: the 1:1 rule's rooms, then default_rooms.
        Assert.Equal("Room 3.1", Only(plans, "Later 1:1").Choice!.Name);
        Assert.Equal("preference 2 (default_rooms)", Only(plans, "Later 1:1").ChoiceWhy);
    }

    [Fact]
    public void No_room_free_says_why()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("Peer call", At(Thu, 10), 30, "alex", "sam", "jo"));
        foreach (var (address, items) in calendar.Schedules)
            items.Add(new ScheduleItem(At(Thu, 9), At(Thu, 12), "busy", "Someone Else"));
        Assert.Equal("no room free that seats 4: 3 busy, 1 too small", Only(Judge(config, calendar), "Peer call").StatusWhy);
    }

    [Fact]
    public void Booked_is_judged_from_the_rooms_schedule_not_the_copy()
    {
        var (config, calendar, _) = Setup();
        // Lists Room 2.1 on your copy, but the room declined: its schedule doesn't show it.
        var listed = Meeting("Listed only", At(Thu, 10), 30, "alex", "sam", "jo") with
        {
            Attendees = [new Attendee("room2.1@example.com", "Room 2.1", "resource", "declined"), new Attendee(FakeCalendar.MeAddress, "Pat", "required", "accepted")],
        };
        // Not on your copy, but the room's schedule shows it under the organiser's name (a forward that didn't update it).
        var shown = Meeting("On the schedule", At(Thu, 14), 30, "alex", "sam", "jo");
        calendar.Items.AddRange([listed, shown]);
        calendar.Schedules["room3.2@example.com"].Add(new ScheduleItem(At(Thu, 14), At(Thu, 14, 30), "busy", "Alex Example"));
        var plans = Judge(config, calendar);
        Assert.Equal(Status.Ask, Only(plans, "Listed only").Status);
        Assert.Equal(Status.Booked, Only(plans, "On the schedule").Status);
        Assert.Equal("Room 3.2", Only(plans, "On the schedule").BookedRoom);
    }

    [Fact]
    public void A_room_that_didnt_take_it_is_followed_by_the_next_then_given_up()
    {
        var (config, calendar, _) = Setup();
        var m = Meeting("My 1:1", At(Thu, 10), 30, "me", "sam");
        calendar.Items.Add(m);
        var state = new State();
        var entry = state.Entry(m);
        ((System.Text.Json.Nodes.JsonArray)entry["attempts"]!).Add((System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject
        {
            ["room"] = "booth3.9@example.com", ["name"] = "Booth 3.9", ["how"] = "add", ["state"] = "done", ["when"] = "2026-11-23T08:55:00",
        });
        // Five minutes on, still waiting for its answer.
        Assert.Equal(Status.Waiting, Only(Judge(config, calendar, state), "My 1:1").Status);
        // Twenty minutes on, its silence is a decline: the next room in the order.
        var later = Judge(config, calendar, state, Now.AddMinutes(15));
        Assert.Equal(Status.Book, Only(later, "My 1:1").Status);
        Assert.Equal("Room 3.1", Only(later, "My 1:1").Choice!.Name);
        Assert.Contains("after no answer from Booth 3.9", Only(later, "My 1:1").ChoiceWhy);
        ((System.Text.Json.Nodes.JsonArray)entry["attempts"]!).Add((System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject
        {
            ["room"] = "room3.1@example.com", ["name"] = "Room 3.1", ["how"] = "add", ["state"] = "done", ["when"] = "2026-11-23T09:10:00",
        });
        Assert.Equal(Status.GaveUp, Only(Judge(config, calendar, state, Now.AddMinutes(40)), "My 1:1").Status);
    }

    [Fact]
    public void A_moved_meeting_is_a_new_one()
    {
        var m = Meeting("My 1:1", At(Thu, 10), 30, "me", "sam");
        Assert.NotEqual(m.Key, (m with { Start = At(Thu, 11), End = At(Thu, 11, 30) }).Key);
    }

    [Fact]
    public void Adding_a_room_keeps_everyone_and_the_teams_location()
    {
        var m = Meeting("Review", At(Thu, 10), 30, "me", "sam", "jo") with
        {
            Attendees = [new Attendee("sam@example.com", "Sam", "required", "accepted"), new Attendee("jo@example.com", "Jo", "optional", "none"),
                new Attendee("room2.1@example.com", "Room 2.1", "resource", "declined")],
        };
        var body = Calendar.AddRoomBody(m, new Room("Room 3.1", "room3.1@example.com"), ["room2.1@example.com"]);
        Assert.Equal("""{"attendees":[{"emailAddress":{"address":"sam@example.com","name":"Sam"},"type":"required"},{"emailAddress":{"address":"jo@example.com","name":"Jo"},"type":"optional"},{"emailAddress":{"address":"room3.1@example.com","name":"Room 3.1"},"type":"resource"}],"locations":[{"displayName":"Microsoft Teams Meeting","locationType":"default"},{"displayName":"Room 3.1","locationEmailAddress":"room3.1@example.com","locationType":"conferenceRoom"}]}""",
            body.ToJsonString());
        Assert.Equal("""{"ToRecipients":[{"EmailAddress":{"Address":"room3.1@example.com","Name":"Room 3.1"}}],"Comment":""}""",
            Calendar.ForwardBody(new Room("Room 3.1", "room3.1@example.com"), "").ToJsonString());
    }
}

public sealed class RunTests
{
    private static string Run(FakeCalendar calendar, Config config, out int code, DateTime? now = null, params string[] args)
    {
        var output = new StringWriter();
        code = Cli.Run(["--config", config.Path, .. args], Host(calendar, output, now));
        return output.ToString();
    }

    private static List<System.Text.Json.Nodes.JsonObject> Items(string folder) =>
        new ReviewQueue.QueueFile(Path.Combine(folder, "review.json")).Items("calendar-room-booking").ToList();

    [Fact]
    public void Ask_mode_writes_nothing_and_raises_one_item_per_day_once()
    {
        var (config, calendar, folder) = Setup();
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        calendar.Items.Add(Meeting("Peer 1:1", At(Thu, 11), 30, "alex"));
        var output = Run(calendar, config, out var code, null, "run");
        Assert.Equal(0, code);
        Assert.Empty(calendar.Writes);
        var item = Assert.Single(Items(folder));
        Assert.Equal("ask", (string)item["type"]!);
        Assert.Equal("Thu 26 Nov: 2 meetings without a room to decide", (string)item["title"]!);
        Assert.Equal(["09:00-09:30 My 1:1 (2 people; rule 1 (1:1s)): book Booth 3.9?", "11:00-11:30 Peer 1:1 (2 people; rule 4 (everything else)): book Room 3.1?"],
            item["meetings"]!.AsArray().Select(x => (string)x!));
        Assert.Contains("1 the rules would book", output);

        // The same list again: not raised again, even once it's been cleared.
        new ReviewQueue.QueueFile(Path.Combine(folder, "review.json")).Done(["calendar-room-booking"]);
        Run(calendar, config, out _, null, "run");
        Assert.Empty(Items(folder));

        // The list changes: raised again.
        calendar.Items.Add(Meeting("Another", At(Thu, 14), 30, "alex", "sam"));
        Run(calendar, config, out _, null, "run");
        Assert.Contains("3 meetings", (string)Assert.Single(Items(folder))["title"]!);
    }

    [Fact]
    public void Act_mode_adds_to_your_meeting_and_forwards_the_rest()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"");
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        calendar.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar.Accepting.Add("booth3.9@example.com");
        var output = Run(calendar, config, out var code, null, "run");
        Assert.Equal(0, code);
        Assert.Equal(["add Booth 3.9 to My 1:1", "forward Boss 1:1 to Booth 3.9"], calendar.Writes);
        Assert.Contains("Booth 3.9 accepted: it's booked.", output);
        // The next run sees both on the booth's schedule: nothing more.
        calendar.Writes.Clear();
        Run(calendar, config, out _, null, "run");
        Assert.Empty(calendar.Writes);
        Assert.Contains("Rooms asked for:", Run(calendar, config, out _, null, "list"));
    }

    [Fact]
    public void A_lost_answer_is_never_sent_again_and_is_raised_once()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"");
        calendar.Items.Add(Meeting("Boss 1:1", At(Thu, 10), 30, "boss"));
        calendar.WriteFails = new GraphDeliveryUnknownException("POST", "/me/events/x/forward", System.Net.HttpStatusCode.ServiceUnavailable, "busy");
        var output = Run(calendar, config, out var code, null, "run");
        Assert.Equal(1, code);
        Assert.Contains("it may have gone", output);
        var state = State.Load(Path.Combine(folder, "state.json"));
        Assert.Equal("sending", Assert.Single(state.Attempts(calendar.Items[0].Key)).State);

        calendar.Writes.Clear();
        output = Run(calendar, config, out _, Now.AddHours(1), "run");
        Assert.Empty(calendar.Writes);
        Run(calendar, config, out _, Now.AddHours(2), "run");
        Assert.Single(Items(folder), i => (string)i["type"]! == "may-have-gone");
        // And book refuses it too.
        Assert.Contains("never sent again", Run(calendar, config, out var refused, Now.AddHours(2), "book", "2026-11-26 10:00", "--yes"));
        Assert.Equal(2, refused);
        Assert.Empty(calendar.Writes);
    }

    [Fact]
    public void A_declined_room_on_your_meeting_is_followed_by_the_next_once()
    {
        var (config, calendar, folder) = Setup(extra: "mode = \"act\"");
        var m = Meeting("My review", At(Thu, 13), 30, "me", "sam", "jo");
        calendar.Items.Add(m);
        calendar.Accepting.Add("room3.2@example.com");
        // Room 3.1 says no: the fake writes its decline onto the meeting when it's added.
        var calendarDeclines = new DecliningCalendar(calendar, "room3.1@example.com");
        var output = new StringWriter();
        var code = Cli.Run(["--config", config.Path, "run"], Host(calendar, output) is var h
            ? new CliHost { Out = output, Err = output, Env = _ => null, Now = () => Now, Calendar = _ => calendarDeclines, Sleep = _ => { }, Notify = (_, _) => { }, Lock = false }
            : h);
        Assert.Equal(0, code);
        Assert.Equal(["add Room 3.1 to My review", "add Room 3.2 to My review"], calendar.Writes);
        Assert.Contains("Room 3.2 accepted", output.ToString());
        // The second add takes the declined room off the meeting.
        Assert.DoesNotContain("room3.1", calendar.Bodies[1]["attendees"]!.ToJsonString());
    }

    [Fact]
    public void More_bookings_than_max_bookings_stops_the_run()
    {
        var (config, calendar, _) = Setup(extra: "mode = \"act\"");
        for (var h = 9; h < 16; h++)
            calendar.Items.Add(Meeting($"Chair {h}", At(Thu, h), 30, "me", "sam", "jo"));
        var output = Run(calendar, config, out var code, null, "run");
        Assert.Equal(1, code);
        Assert.Contains("more than max_bookings (5)", output);
        Assert.Empty(calendar.Writes);
    }

    [Fact]
    public void Book_previews_without_yes_and_skip_remembers_no()
    {
        var (config, calendar, folder) = Setup();
        calendar.Items.Add(Meeting("Peer 1:1", At(Thu, 11), 30, "alex"));
        var output = Run(calendar, config, out var code, null, "book", "2026-11-26", "11:00", "--room", "Room 2.1");
        Assert.Equal(0, code);
        Assert.Contains("room   Room 2.1 <room2.1@example.com> (your choice)", output);
        Assert.Contains("forwarded to the room", output);
        Assert.Contains("Nothing sent.", output);
        Assert.Empty(calendar.Writes);

        Run(calendar, config, out _, null, "skip", "2026-11-26 11:00");
        Assert.Contains("you said no room", Run(calendar, config, out _, null, "plan"));
        Run(calendar, config, out _, null, "book", "2026-11-26 11:00", "--yes");
        Assert.Equal(["forward Peer 1:1 to Room 3.1"], calendar.Writes);
    }

    [Fact]
    public void Plan_shows_each_office_day()
    {
        var (config, calendar, _) = Setup();
        calendar.Items.Add(Meeting("My 1:1", At(Thu, 9), 30, "me", "sam"));
        var output = Run(calendar, config, out var code, null, "plan");
        Assert.Equal(0, code);
        Assert.Contains("Thu 26 Nov  office  (work plan: office 08:00-17:00)", output);
        Assert.Contains("09:00-09:30  My 1:1", output);
        Assert.Contains("would book Booth 3.9 by adding it to your meeting (preference 1 (rule 1's rooms); rule 1 (1:1s))", output);
        Assert.Contains("(nothing changed)", output);
    }

    [Fact]
    public void Rooms_lists_the_order_per_rule()
    {
        var (config, calendar, _) = Setup();
        var output = Run(calendar, config, out _, null, "rooms");
        Assert.Contains("rule 1 (1:1s):", output);
        Assert.Matches(@"Default order[^\n]*\n\s+Room 3\.1 ", output);
    }
}

/// <summary>A room that declines: added to a meeting you organise, it shows as declined on your copy.</summary>
internal sealed class DecliningCalendar(FakeCalendar inner, string declining) : ICalendar
{
    public string Me() => inner.Me();
    public IReadOnlyList<Segment> WorkPlan(DateOnly first, DateOnly last) => inner.WorkPlan(first, last);
    public IReadOnlyList<Meeting> Events(DateOnly first, DateOnly last) => inner.Events(first, last);
    public Meeting? Event(string id) => inner.Event(id);
    public IReadOnlyDictionary<string, IReadOnlyList<ScheduleItem>?> RoomSchedules(IReadOnlyList<string> addresses, DateTime from, DateTime to) =>
        inner.RoomSchedules(addresses, from, to);
    public void Forward(Meeting meeting, Room room, string comment) => inner.Forward(meeting, room, comment);

    public void AddRoom(Meeting meeting, Room room, IReadOnlyCollection<string> drop)
    {
        inner.AddRoom(meeting, room, drop);
        var i = inner.Items.FindIndex(e => e.Id == meeting.Id);
        var response = string.Equals(room.Address, declining, StringComparison.OrdinalIgnoreCase) ? "declined" : "none";
        inner.Items[i] = inner.Items[i] with
        {
            Attendees = [.. inner.Items[i].Attendees.Where(a => !drop.Contains(a.Address, StringComparer.OrdinalIgnoreCase)), new Attendee(room.Address, room.Name, "resource", response)],
        };
    }
}

public sealed class SettingsTests
{
    [Fact]
    public void The_example_settings_load()
    {
        var folder = Path.Combine(Path.GetTempPath(), "room-booking-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var example = Path.Combine(AppContext.BaseDirectory, "calendar-room-booking.example.toml");
        var text = File.ReadAllText(example).Replace("base = \"base-config.toml\"", "");
        var path = Path.Combine(folder, "settings.toml");
        File.WriteAllText(path, text);
        var config = Config.Load(path, name => name == "LOCALAPPDATA" ? folder : null);
        Assert.Equal("ask", config.Mode);
        Assert.Equal(4, config.Rooms.Count);
        Assert.Equal(5, config.Rules.Count);
        Assert.Equal(["Booth 4.8"], config.Rules[0].Rooms!);
        Assert.Equal("calendar-room-booking", config.Review!.Source);
    }

    [Theory]
    [InlineData("mode = \"maybe\"", "mode should be")]
    [InlineData("default_rooms = [\"Nowhere\"]", "isn't one of the [[rooms]]")]
    [InlineData("colour = 1", "unknown setting 'colour'")]
    public void Bad_settings_say_what_is_wrong(string line, string message)
    {
        var folder = Path.Combine(Path.GetTempPath(), "room-booking-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "settings.toml");
        File.WriteAllText(path, line + "\n[[rooms]]\nname = \"R\"\naddress = \"r@example.com\"\n");
        Assert.Contains(message, Assert.Throws<UsageException>(() => Config.Load(path, _ => null)).Message);
    }
}
