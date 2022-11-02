#SingleInstance Force
#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    responses := ["Hi","Hey"]
    ^r::
        Send, ^r+{Tab 3}^a
        Sleep, 200
        Send, ^c
        Sleep, 200
        fullEmail:=clipboard
        RegExMatch(fullEmail,"^\w+", firstName)
        Sleep, 200
        clipboard:=firstName
        Sleep, 200
        Random, greet, 1, responses.Length()
        response := responses[greet]
        Send, {Tab 3}
        Send, %response%
        Send, {Space}^v,{Enter 2}
    return
}