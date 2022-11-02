#SingleInstance Force
AHKPath:="C:\Program Files\AutoHotkey\AutoHotkey.exe"

Run %AHKPath% "EmailGreetings.ahk"
Run %AHKPath% "TicketTemplate.ahk"

^!+r::Reload