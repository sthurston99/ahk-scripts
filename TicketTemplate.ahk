#SingleInstance Force
#If (WinActive("ahk_exe RangerMSP.exe"))
{
    ^n::
        Send, ^n
        Sleep, 500
        Send {Tab 3}
        clipboard := "$User contacted us via $ContactType about the following issue:"
        Send, ^v{Enter}
    return
}