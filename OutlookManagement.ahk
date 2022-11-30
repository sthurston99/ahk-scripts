#NoEnv
#SingleInstance Force
#Include EmailFuncs.ahk

#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    ^r::
        GetCurrentEmail().replyall().Display()
        Send, !has{Down 3}{Enter}{Up 2}
        Send, % GenerateGreeting() . " " . GetFirstName(GetStandardName(email.SenderName)) . ","
        Send, {Enter 2}
    return

    ^n::Send, ^n!nas{Down 3}{Enter}
}