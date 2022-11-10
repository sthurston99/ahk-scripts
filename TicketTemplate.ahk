#SingleInstance Force

GetEmailBody()
{
    olItem := ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
    sender := RegExReplace(olItem.SenderName, "\w\. ")
    RegExMatch(olItem.SenderName,"\w{2,}(?: )", senderFN)
    body := Trim(RegExReplace(olItem.Body, "s)(?=" . senderFN . "|" . sender . "|" . olItem.SenderName . ").*"), " `t`n`r")
    return body
}

SetMyLabel()
{
    Sleep, 50
    Click, 300 100
        Click, 50 250 0 Rel
        Loop, 50
        {
            Click, WD 1
            Sleep, 5
        }
    Click
}

#If (WinActive("RangerMSP 28 SQL"))
{
    ^n::
        Send, ^n
        Sleep, 500
        Send, {Tab 3}
        clipboard := "$User contacted us via $ContactType about the following issue:"
        Send, ^v{Enter}
    return

    ^+l::SetMyLabel()
}

#If (WinActive("New Ticket"))
{
    ^+e::
        olItem := ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
        sender := RegExReplace(olItem.SenderName, "\w\. ")
        Send, ^a^x
        clipboard := StrReplace(clipboard, "$User", sender)
        clipboard := StrReplace(clipboard, "$ContactType", "email")
        RegExMatch(olItem.SenderEmailAddress, "(?<=@).*(?=\.)", userAccount)
        userAccount := SubStr(userAccount, 1, 4)
        Send, ^v
        Send, +{Tab 3}
        Send, %userAccount%
        KeyWait, Enter, D
        Send, {Enter}{Tab 3}
        body := GetEmailBody()
        Send, % body
        olItem.Categories := "ST"
        olItem.UnRead := False
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

    ^g::SetMyLabel()
}

#If (WinActive("New Charge - (Labor)"))
{
    ^e::
        Send, +{Tab 4}Labor{Space}+9R{Enter}{Tab 3}{Enter}
    return
}