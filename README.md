# calendar-room-booking

Books meeting rooms for the Teams meetings you're going to on your office days, **chosen by rules, not for every
meeting**. On a day your work plan puts you in the office, a call you'd rather not take at your desk gets a room: your
1:1s in a booth, the calls you chair in your usual room, nothing for a big all-hands you're only listening to, and a
question for the rest.

```
calendar-room-booking plan
```

```
mode = ask: a run books nothing; it asks about these in the review queue.
  Mon 23 Nov  -       (work plan: remote)
  Thu 26 Nov  office  (work plan: office 09:00-17:00)
    10:00-10:30  Weekly 1-2-1 - Jo                           2  would book Booth 4.8 by adding it to your meeting (preference 1 (rule 1's rooms); rule 1 (1:1s))
    10:30-11:00  Team stand-up                               5  would book Room 4.2 by forwarding to the room (one room for the back-to-back meetings 10:30-12:00; rule 3 (my team's meetings))
    11:00-12:00  Product review                              6  ask: Room 4.2 (one room for the back-to-back meetings 10:30-12:00; rule 5 (everything else))
    14:00-14:30  Company all-hands                          85  no room (rule 2 (big calls))
    16:00-16:30  Supplier call                               4  booked: Room 3.1 (on the room's schedule under the organiser)
      - 12:00-13:00  Lunch: skip: your own blocks and reminders
      - 17:00-18:00  Travel: not a Teams meeting
23/11 to 07/12: 1 office days, 5 meetings: 1 with a room, 2 the rules would book, 1 to ask, 1 no (nothing changed)
```

## Install

Windows (x64) only; nothing else is needed.

1. Download `calendar-room-booking-<version>-win-x64.zip` from [Releases](https://github.com/gareth-jones-zen/calendar-room-booking/releases).
2. Unzip it to a folder that will stay put, such as `%LOCALAPPDATA%\Programs\calendar-room-booking`.
3. `calendar-room-booking config` creates the settings file (or start from `calendar-room-booking.example.toml`): list
   your rooms and write your rules. `calendar-room-booking sign-in` signs you in (once).
4. `calendar-room-booking plan` shows what it would do. Leave `mode = "ask"` for a while: `run` books nothing and puts
   what it would book in the review queue. Switch to `mode = "act"` when the plan looks right;
   `calendar-room-booking autostart on` then runs it every 15 minutes in working hours.

**It needs Calendars.ReadWrite** on Microsoft's "Graph Command Line Tools" app; `sign-in` asks for it. Forwarding an
invitation is a calendar action (Graph's `event: forward`), so no mail permission is needed.

## What it does

### Office days

From your work plan (Outlook and Teams' work hours and location): a day with office time is an office day, **for those
hours only**. A day of leave or a bank holiday has none (the shared `[leave]` rules, below); on a half day off, meetings
in the half you're off are left out. Today and the next `days_ahead` days (default 14).

### Candidates

On an office day, a meeting is a candidate when:

- **you're going:** you organise it, or your response is one of `responses` (accepted or tentatively accepted);
- **it's a Teams meeting** (it has a join link);
- **someone else is at it**, it isn't shown as free, and no `[[skip]]` rule matches its subject or category (your own
  blocks and reminders);
- **it falls inside your office hours**, is no longer than `max_minutes` (default 240), and hasn't started.

`plan` lists the meetings that aren't candidates too, with why.

### Which want a room: categories, then `[[rules]]`, first match

An Outlook category on the meeting decides first: `room_category` ("Room") means yes, `no_room_category` ("No room")
means no: your one-click override for a single meeting. Otherwise the first `[[rules]]` entry whose tests all pass
decides, and with no match it's ask. The tests, all optional:

| Key | Passes when |
|---|---|
| `subject` | the subject matches the pattern (any case) |
| `organiser` | the organiser's address or name is one of these; `"me"` is a meeting you organise |
| `one_to_one` | exactly two people are at it (you and one other), or not |
| `min_attendees`, `max_attendees` | how many people are at it (you, the organiser and everyone invited; rooms don't count) |
| `response` | your response: `organizer`, `accepted`, `tentativelyAccepted` |
| `series` | the meeting's series master id |
| `category` | it carries this Outlook category |
| `with` | any of these addresses is at it |

Each rule says `room = "yes"`, `"no"` or `"ask"`, and may give `rooms = [...]`: the rooms to try first for meetings it
matches, in order. `why` names it in `plan`, the log and review items.

### Which room

The rooms come from `[[rooms]]` in the settings (name, address, and optionally floor and capacity). For each meeting
they're tried in this order: **the rule's `rooms`, then `default_rooms`, then rooms on `room_floors` (in that order),
then the smallest with a known capacity, then the rest by name**, leaving out any whose known capacity is less than the
number of people. The first that's **free for the whole meeting** (its free/busy schedule) is chosen. `calendar-room-booking
rooms` shows the order per rule.

**Back-to-back meetings share a room.** Meetings with gaps of `join_gap_minutes` (default 10) or less form a run; where
one room is free for every meeting in the run that needs one, they all get it (each meeting is booked on its own, in the
same room). A room already booked for a meeting in the run comes first, then the biggest meeting's order. If no single
room is free throughout, each meeting gets its own.

### "Has a room" is judged from the room's own schedule

A meeting has a room only when **a room's own schedule shows it**: an item at the meeting's times whose subject names
the organiser (many organisations' room calendars show the organiser's name as the subject) or is the meeting's subject.
Your copy of a meeting can list a room that declined, so a listing alone doesn't count; and a room whose schedule shows
the meeting counts even if your copy doesn't list it. A room that hides subjects is matched on times alone, but only if
it's listed on the meeting or this tool asked it.

So **a meeting moved to a time its room isn't free** (the room declines the update) has no room at the next run, and
gets one.

### How a room is added

- **A meeting you organise:** the room is added to the meeting itself, as a resource attendee and a location (the Teams
  location is kept). **Every attendee gets an update.** A room this tool added earlier that didn't take the booking is
  taken off at the same time.
- **A meeting you attend:** the invitation is **forwarded to the room** (Graph's `event: forward`), as Outlook's "Add
  room" does for an attendee. The room's mailbox books it under the organiser's name; Exchange tells the organiser that
  the invitation was forwarded, and adds the room to the organiser's copy. Nobody else is told.

After asking, it reads the room's schedule (and, for your own meeting, the room's response) up to four times, five
seconds apart. A room that declines your own meeting is followed by the next free room straight away. A room that
hasn't shown the booking within `answer_wait_minutes` (default 10) counts as a decline at the next run, and the next free
room is asked. After `max_attempts` rooms (default 2), the meeting is raised as "no room".

### Never the same meeting twice

Every request is recorded in the state file (`state_file`) **before** it's sent, as "sending", keyed by the meeting's
iCalUId and start time, and marked "done" once Microsoft answers. If the answer is lost (a 5xx, a timeout, a dropped
connection: background-tools-lib's send-once writes), the record stays "sending": **nothing is sent for that meeting at
that time again**, by `run` or by `book`, and a review item says to check the room in Outlook. A moved meeting is a new
key, so it can be booked again at its new time. A state file that can't be read stops every write.

### Ask mode, act mode and the review queue

- **`mode = "ask"` (the default)** books nothing. Each run writes **one review item per office day** listing the
  meetings to decide: those the rules would book (with the room) and those whose rule says ask (with the room it
  suggests). An item is raised again only when the day's list changes; one you've cleared isn't raised again for the
  same list. Answer with `book` or `skip`.
- **`mode = "act"`** books the "yes" meetings itself, and raises only the "ask" ones, and those that want a room with
  none free.
- Other items: **`may-have-gone`** (a request whose answer was lost: check the room) and **`no-room`** (no room took
  it after `max_attempts`), each raised once per meeting.

Items go to the queue in `[review]` (background-tools-lib's ReviewQueue: `file`, `source`); with no `[review]`, none
are raised. Item types: `ask`, `may-have-gone`, `no-room`.

**Safety:** a run that would book more than `max_bookings` (default 10) rooms books none and says so. A run that fails
is logged, and notified on the third failure in a row (or at once for a sign-in or a safety stop), once, until a run
succeeds.

### Known limitations

- **Rooms are never given back.** A room forwarded for a meeting you attend stays booked if the day stops being an
  office day: an attendee can't cancel a room's booking, and a leftover booking is a small price for days that rarely
  change. For **your own meetings** the room also stays when the day or the rule changes: taking it off would send
  everyone another update, so it's left for you to remove in Outlook.
- **A forward relies on the room accepting forwarded invitations.** A room whose booking policy refuses them, or
  refuses the length or how far ahead, declines; the next room is tried.
- **Room sizes** come only from the settings (`capacity`); a room without one counts as fitting.

## Commands

```
calendar-room-booking plan                    each office day: the meetings, the rule, the room, and why
calendar-room-booking run                     ask mode: review items only; act mode: book the "yes" meetings
calendar-room-booking book MEETING [--room R] [--yes]   book one (a preview without --yes)
calendar-room-booking skip MEETING            no room for that one (at that time)
calendar-room-booking list                    the rooms it has asked for, and the meetings skipped
calendar-room-booking rooms                   the rooms in the order they're tried, per rule
calendar-room-booking autostart on            run on the [schedule] (off, restart, status)
calendar-room-booking sign-in                 sign in to Microsoft (once; opens a browser)
calendar-room-booking config                  create the settings file (if missing) and say where it is
```

`MEETING` is an event id or `"YYYY-MM-DD HH:MM"` (its start; the date and time may also be two arguments). `book`
books whatever the rules say (and clears a `skip`), on any day, but never a meeting that already has a room or whose
earlier request may have gone; `--room` takes a name or address from `[[rooms]]` and must be free. `plan`, `list` and
`rooms` only read.

**Exit codes:** 0 done; 1 failed, stopped for safety, or a request that may have gone; 2 settings, arguments or sign-in.

## Settings

`%APPDATA%\calendar-room-booking\config.toml` by default (`--config` or `ROOM_BOOKING_CONFIG` for another). Read
strictly: an unknown setting, or one of the wrong type, is an error that names it. `calendar-room-booking.example.toml`
is a complete example.

```toml
base = "base-config.toml"         # optional: settings shared with other tools (background-tools-lib's BaseSettings)
login_hint = "you@example.com"    # also your address, for counting people
tenant = "organizations"
msal_cache = "..."                # optional: another tool's token cache, to share its sign-in
time_zone = "GMT Standard Time"
log_file = "..."                  # default %LOCALAPPDATA%\calendar-room-booking\calendar-room-booking.log
state_file = "..."                # the send-once record; default %LOCALAPPDATA%\calendar-room-booking\state.json
enabled = true
mode = "ask"                      # or "act"
days_ahead = 14
max_minutes = 240
join_gap_minutes = 10
max_bookings = 10
max_attempts = 2
answer_wait_minutes = 10
notify = true
room_category = "Room"
no_room_category = "No room"
responses = ["accepted", "tentativelyAccepted"]
forward_comment = ""
default_rooms = ["Room 4.2"]
room_floors = ["4", "3"]

[[rooms]]
name = "Room 4.2"                 # as Outlook shows it
address = "room4.2@example.com"
floor = "4"                       # optional
capacity = 8                      # optional

[[skip]]
subject = '^\s*(lunch|travel|focus)\b'   # and/or category = "..."
why = "your own blocks"

[[rules]]
one_to_one = true
room = "yes"
rooms = ["Booth 4.8"]
why = "1:1s"

[leave]                           # background-tools-lib's shared [leave] keys; a base file's [leave] can supply them
[review]                          # the review queue: file, source (a base file's [review] can supply file)
[schedule]                        # background-tools-lib's shared [schedule] keys
```

### [schedule]

When the task runs, in the keys every background-tools-lib tool shares (`every_minutes`, `daily_at`, `hours`, `days`,
`at_logon`, `after_unlock_minutes`, `end_date`): see background-tools-lib's README. This tool's defaults:

```toml
[schedule]
every_minutes = 15
hours = "07:00-19:00"
days = "Mon-Fri"
at_logon = true
```

`autostart on` registers the task from it and records it in the task's description; `autostart status` says when the
settings have changed since (run `autostart on` again). A run outside `hours` or `days` does nothing.

## Building

.NET 11 (see `global.json`), with the shared libraries from
[background-tools-lib](https://github.com/gareth-jones-zen/background-tools-lib) checked out beside this repo
(`MeetingRooms` holds the room list, room order, free/busy reading and the "does the room's schedule show this meeting"
test):

```
dotnet test -c Release
dotnet publish src/RoomBooking.Console -c Release -r win-x64 -o out    # Native AOT
```

A tag `vX.Y.Z` builds a release (`.github/workflows/release.yml`), against the library commit pinned there.

**This is a personal project. It is not an official product of any employer**, and isn't supported by one.
