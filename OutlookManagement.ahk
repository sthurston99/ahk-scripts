#SingleInstance Force
#Include EmailFuncs.ahk

#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    ^r::
        Send, ^+r
        Sleep, 400
        Send, !e2as{Down 3}{Enter}{Up 2}
        Send, % GenerateGreeting()
        Send, {Space}
        Send, % GetFirstName()
        Send, ,{Enter 2}
    return

    ^n::Send, ^n!nas{Down 3}{Enter}
}