// calendar-room-booking.exe. Output is Unicode to a console and UTF-8 when redirected (see ConsoleText).
return RoomBooking.Cli.Run(args, new RoomBooking.CliHost { Out = WinBackground.ConsoleText.Out, Err = WinBackground.ConsoleText.Error });
