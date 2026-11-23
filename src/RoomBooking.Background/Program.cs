// calendar-room-booking-background.exe: no window, ever. What the scheduled task runs (`run`); its log is the settings' log_file.
return RoomBooking.Cli.Run(args, new RoomBooking.CliHost { Out = TextWriter.Null, Err = TextWriter.Null });
