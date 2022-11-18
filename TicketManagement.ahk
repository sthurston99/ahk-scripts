#SingleInstance Force
#Include EmailFuncs.ahk

SetLabel() {
    Sleep, 50
    Click, 300 100
    Click, 50 250 0 Rel
    Loop, 50 {
        Click, WD 1
        Sleep, 5
    }
    Click
}

SetRemoteLabor() {
    Send, {Click 215 185}Labor{Space}+9R{Enter}{Tab 3}{Enter}
}

#If (WinActive("AHK_exe RangerMSP.exe"))
{
    #If (WinActive("RangerMSP 28 SQL"))
    {
        ^n::
            Send, ^n
            Sleep, 500
            Send, {Tab 3}
            clipboard := "$User contacted us via $ContactType about the following issue:"
            Send, ^v{Enter}
        return

        ^+l::SetLabel()

        !+c::
            Send, !+c
            Sleep, 500
            If(WinActive("Confirm")) {
                Send, y
            }
            Click, 30 395
            SetRemoteLabor()
        return
    }

    #If (WinActive("New Ticket"))
    {
        ^+e::
            Send, ^a^x
            clipboard := StrReplace(clipboard, "$User", GetSender())
            clipboard := StrReplace(clipboard, "$ContactType", "email")
            userAccount := SubStr(GetEmailDomain(), 1, 4)
            Send, ^v
            Send, % GetEmailBody()
            Send, +{Tab 3}
            Send, %userAccount%
            KeyWait, Enter, D
            Send, {Enter}{Tab 3}^g
            SetAsHandled()
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

    #If (WinActive("New Charge - (Labor)"))
    {
        ^e::
            Send, % GetFirstName()
            Send, {Space}Emailed in:{Enter}
            Send, % GetEmailBody()
            SetAsHandled()
        return

        ^r::SetRemoteLabor()

        ^g::
            Click, 355, 130
            Sleep, 250
            Send, ^g
        return

        ^t::
            InputBox, mins, Minutes:,,,150,100
            KeyWait, Enter, D
            Click, 320, 130
            Send, % mins
            Click, 30 395
        return
    }

    #If (WinActive("Timer"))
    {
        ^Enter::
            Click, 230, 50
            WinWaitActive, New Charge - (Labor),,500
            Click, 30 395
            Sleep, 100
            SetRemoteLabor()
        return
    }
}