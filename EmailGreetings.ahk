#SingleInstance Force
#Include EmailFuncs.ahk

#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    ^r::
        Send, ^r
        Send, % GenerateGreeting()
        Send, {Space}
        Send, % GetFirstName()
        Send, {Enter 2},
    return
}