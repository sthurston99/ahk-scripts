#SingleInstance Force
#If (WinActive("ahk_exe RangerMSP.exe"))
{
    ^n::
        Send, ^n
        Sleep, 500
        Send, {Tab 3}
        clipboard := "$User contacted us via $ContactType about the following issue:"
        Send, ^v{Enter}
    return
    
    ^+e::
        WinActivate, ahk_exe OUTLOOK.EXE
        Send, {Alt}jdm{Down}{Enter}
        Sleep, 200
        userName := clipboard
        WinActivate, ahk_exe RangerMSP.exe
        Send, ^a^x
        clipboard := StrReplace(clipboard, "$User", userName)
        clipboard := StrReplace(clipboard, "$ContactType", "email")
        Sleep, 200
        Send, ^v
    return
}