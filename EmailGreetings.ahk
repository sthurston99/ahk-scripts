#SingleInstance Force
#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    responses := ["Hi","Hey"]
    ^r::
        Send, ^r+{Tab 3}^a
        Sleep, 150
        Send, ^c
        Sleep, 150
        fullEmail:=clipboard
        RegExMatch(fullEmail,"^\w+", firstName)
        Sleep, 150
        clipboard:=firstName
        Sleep, 150
        Random, greet, 1, responses.Length()
        response := responses[greet]
        Send, {Tab 3}
        Send, %response%
        Send, {Space}^v,{Enter 2}
    return
}