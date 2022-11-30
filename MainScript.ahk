#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%.
AHKPath:="C:\Program Files\AutoHotkey\AutoHotkey.exe"

; Function Definitions
; Previously EmailFuncs.ahk and parts of TicketManagement.ahk

; Fetches current selected email within Outlook
GetCurrentEmail() {
    return ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
}

; Fetches the Full Name of the sender of the COM MailItem Object, or the currently selected email
GetSender(email:="") {
    if(email = "") {
        email := GetCurrentEmail()
    }
    return email.SenderName
}

; Converts a Full Name into a Standard Name by removing the Middle Initial
; If none passed in, defaults to sender of Current Email
GetStandardName(name:="") {
    if(name = "") {
        name := GetSender()
    }
    return RegExReplace(name,"\w\. ")
}

; Extracts the first name from a Standard Name
; If none, defaults to sender of Current email
GetFirstName(name:="") {
    if(name = "") {
        name := GetStandardName()
    }
    return RegExReplace(name,"\s.*")
}

; Trims out the contents of an email after the name provided
; If no sender provided, passes empty string internally
; If no email provided, uses current selected email in outlook
GetEmailBody(email:="",name:="") {
    whitespace := " `t`n`r"
    regexstr := "s)(?=From:|" . GetFirstName(name) . "|" . GetStandardName(name) . "|"
    linecleaner := "\s{2,}"
    
    if(email = "") {
        email := GetCurrentEmail().Body
    }

    if(name = "") {
        regexstr := regexstr . GetSender() . ").*"
    } else {
        regexstr := regexstr . name . ").*"
    }
    return Trim(RegExReplace(RegExReplace(email, regexstr), linecleaner, "`n`n"), whitespace)
}

; Sets the Category of the email to mark that it was handled by me, and marks as read
SetAsHandled() {
    email := GetCurrentEmail()
    email.Categories := "ST"
    email.UnRead := False
    return
}

; Returns the domain of an email address
GetEmailDomain(address:="") {
    if(address = "") {
        address := GetCurrentEmail().SenderEmailAddress
    }
    RegExMatch(address, "(?<=@).*(?=\.)", account)
    return account
}

; Returns a random boilerplate email greeting
GenerateGreeting() {
    greetings := ["Hi","Hey", "Good $time"]
    FormatTime, hour,, H
    if (hour > 7) and (hour < 12)
        timestr := "morning"
    else if (hour > 12) and (hour < 18)
        timestr := "afternoon"
    else
        timestr := "day"
    Random, idx, 1, greetings.Length()
    greet := greetings[idx]
    return StrReplace(greet, "$time", timestr)
}

; Sets my personal label onto a ticket
;; Originally made to make the task repeatable, however is currently only used for one hotkey
;; Will likely need to refactor or move actual code to hotkey
SetLabel() {
    Click, 300 100
    WinWait, DatLabelSelectChkListFrm
    ControlClick, Simon
}

; Sets the labor type of a charge to be remote
SetRemoteLabor() {
    Send, {Click 215 185}Labor{Space}+9R{Enter}{Tab 3}{Enter}
}

; General Hotkeys

; Reloads Script
^!+r::Reload

; Pauses execution
^!+p::Pause

; VSCodium Hotkeys

#If (WinActive("ahk_exe VSCodium.exe"))
{
    ; Generates a Table of Contents using Headers of a Markdown File
    ;; Should work fine for general applications but is untested outside of current usecase context
    ^!m::
        Send, ^a^x
        instr := RegExReplace(clipboard, "m)^(?!#+.*).*\R")
        instr := RegExReplace(instr, "^.*[^\s]")
        outstr := ""
        Loop, Parse, instr, `n
        {
            RegExReplace(A_LoopField, "#", "#", rCount)
            rCount -= 1
            tspace :=
            Loop, %rCount%
                tspace := tspace . A_Tab
            str := % A_LoopField
            ; Replace headings with bullets
            str := % RegExReplace(str, "#+", tspace . "*")
            ; Place headings into Markdown Link Format
            str := % RegExReplace(str, "(?<=\*\s)(.*)", "[$1](#$1)")
            ; Trim Newlines
            str := % RegExReplace(str, "m)\R")
            ; Condense filenames and hotkeys to links
            str := % RegExReplace(str, "[\.+](?!\S+\])")
            ; Replace spaces with plusses
            str := % RegExReplace(str, "#(\w+)\s(\w+)", "#$1+$2")
            outstr := outstr . str . "`n"
        }
        RegExMatch(clipboard, "m)^# (?!Table of Contents)(.*\R)*.*", out)
        clipboard := "# Table of Contents`r`n" . outstr . out
        Send, ^v
    return
}

; Outlook Hotkeys

#If (WinActive("ahk_exe OUTLOOK.EXE"))
{
    ; Generates a reply email and autofills with standard greeting and email signature
    ^r::
        GetCurrentEmail().replyall().Display()
        Send, !has{Down 3}{Enter}{Up 2}
        Send, % GenerateGreeting() . " " . GetFirstName(GetStandardName(email.SenderName)) . ","
        Send, {Enter 2}
    return

    ; Autoapplies email signature on new emails
    ^n::Send, ^n!nas{Down 3}{Enter}
}

; RangerMSP Hotkeys

#If (WinActive("AHK_exe RangerMSP.exe"))
{
    #If (WinActive("RangerMSP"))
    {
        ; Creates a new ticket prefilling text
        ^n::
            Send, ^n
            WinWaitActive, ahk_class TDatNewSupportTicketsFrm
            Send, {Tab 3}{F8}{Enter}
        return

        ; Calls the SetLabel Function
        ^+l::SetLabel()

        ; Creates a new charge and 
        !+c::
            Send, !+c
            Sleep, 500
            WinWaitActive, New Charge
            SetRemoteLabor()
        return
    }

    #If (WinActive("New Ticket"))
    {
        ; Autofills a new ticket with the contents of an email
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

        ; Prompts for input for quick phone triage
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
        ; Pastes email content into body of charge
        ^e::
            Send, % GetFirstName()
            Send, {Space}Emailed in:{Enter}
            Send, % GetEmailBody()
            SetAsHandled()
        return

        ; Manually calls SetRemoteLabor function in case of error
        ^r::SetRemoteLabor()

        ; Overwrites save charge button to automatically Round Down time
        ^g::
            ControlClick, TBitBtn1,,,,,NA
            WinWaitActive, New Charge
            While WinActive("ahk_class TDatSlipsDtlFrm") {
                ControlClick, TButton1,,,,,NA
            }
        return

        ; Prompts for minutes spent on charge for quick input
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
        ; Quick close of timer and creation of remote charge
        ^Enter::
            Click, 230, 50
            WinWaitActive, New Charge - (Labor),,500
            Click, 30 395
            Sleep, 100
            SetRemoteLabor()
        return
    }
}