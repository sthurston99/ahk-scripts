#SingleInstance Force
SetKeyDelay, 30
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
        olItem := ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
        Send, ^a^x
        clipboard := StrReplace(clipboard, "$User", RegExReplace(olItem.SenderName, "\w\. "))
        clipboard := StrReplace(clipboard, "$ContactType", "email")
        Send, ^v
    return
}