#If (WinActive("ahk_exe OUTLOOK.EXE"))
    ^r::
        Send ^r
        Send +{Tab 3}
        Send ^a
        Sleep, 150
        Send ^c
        Sleep, 150
        fullEmail:=clipboard
        RegExMatch(fullEmail,"^\w+", firstName)
        Sleep, 150
        clipboard:=firstName
        Sleep, 150
        Send {Tab 3}Hi{Space}^v,{Enter 2}
    return
#If
