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
        olItem := ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
        Send, ^a^x
        clipboard := StrReplace(clipboard, "$User", RegExReplace(olItem.SenderName, "\w\. "))
        clipboard := StrReplace(clipboard, "$ContactType", "email")
        Send, ^v
    return

    ^+p::
        InputBox, userName, Name:,,,150,100
        InputBox, userAccount, Account:,,,150,100
        Send, ^a^x
        clipboard := StrReplace(clipboard, "$User", userName)
        clipboard := StrReplace(clipboard, "$ContactType", "phone")
        Send, ^v
        Send, +{Tab 3}
        Send, %userAccount%
        KeyWait, Enter, D
        Send, {Tab 3}
    return
}