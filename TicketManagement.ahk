#NoEnv
#SingleInstance Force
#Include EmailFuncs.ahk

SetLabel() {
    Click, 300 100
    WinWait, DatLabelSelectChkListFrm
    ControlClick, Simon
}

SetRemoteLabor() {
    Send, {Click 215 185}Labor{Space}+9R{Enter}{Tab 3}{Enter}
}

#If (WinActive("AHK_exe RangerMSP.exe"))
{
    #If (WinActive("RangerMSP"))
    {
        ^n::
            Send, ^n
            WinWaitActive, ahk_class TDatNewSupportTicketsFrm
            Send, {Tab 3}{F8}{Enter}
        return

        ^+l::SetLabel()

        !+c::
            Send, !+c
            Sleep, 500
            WinWaitActive, New Charge
            SetRemoteLabor()
        return
    }

    #If (WinActive("New Ticket"))
    {
        ^e::
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

        ^p::
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
            ControlClick, TBitBtn1,,,,,NA
            WinWaitActive, New Charge
            While WinActive("ahk_class TDatSlipsDtlFrm") {
                ControlClick, TButton1,,,,,NA
            }
        return

        ^t::
            InputBox, mins, Minutes:,,,150,100
            KeyWait, Enter, D
            ControlClick, TAdrockDateTimeEdit1
            Send, % mins
            ControlClick, TCmtDBMemoValueSelect1
            Send, {Down 10}
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