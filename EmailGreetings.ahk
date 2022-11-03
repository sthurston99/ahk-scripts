#SingleInstance Force

#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    responses := ["Hi","Hey"]
    ^r::
        Send, ^r
        olItem := ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
        RegExMatch(olItem.SenderName,"\w{2,}(?: )", match)
        clipboard := match
        Random, greet, 1, responses.Length()
        response := responses[greet]
        Send, %response%
        Send, {Space}^v,{Enter 2}
    return
}