SendMode "Input"
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Window"
AHKPath:="C:\Program Files\AutoHotkey\AutoHotkey.exe"
SetTitleMatchMode "RegEx"

; ahk-scripts by Simon Thurston
; 1.0.0 (December 27, 2022)
; Made with love in New Jersey

; Function Definitions
; Previously EmailFuncs.ahk and parts of TicketManagement.ahk

; Fetches current selected email within Outlook
GetCurrentEmail()
{
	Return ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
}

; Fetches the Full Name of the sender of the COM MailItem Object, or the currently selected email
GetSender(email)
{
	Return email.SenderName
}

; Wrapper for GetSender to get Sender from Current Email
GetCurrentSender()
{
	Return GetSender(GetCurrentEmail())
}

; Converts a Full Name into a Standard Name by removing the Middle Initial
; If none passed in, defaults to sender of Current Email
GetStandardName(name:="")
{
	If(name = "")
	{
		name := GetCurrentSender()
	}
	Return RegExReplace(name,"\w\. ")
}

; Extracts the first name from a Standard Name
; If none, defaults to sender of Current email
GetFirstName(name:="")
{
	If(name = "")
	{
		name := GetStandardName()
	}
	Return RegExReplace(name,"\s.*")
}

; Extracts First and Last initial from a passed in name
GetInitials(name)
{
	Return RegExReplace(GetStandardName(name),"(?!\b\w).")
}

; Trims out the contents of an email after the name provided
; If no sender provided, passes empty string internally
; If no email provided, uses current selected email in outlook
GetEmailBody(email:="",name:="")
{
	whitespace := " `t`n`r"
	regexstr := "s)(?=From:|" . GetFirstName(name) . "|" . GetStandardName(name) . "|"
	
	If(email = "")
	{
		email := GetCurrentEmail().Body
	}

	If(name = "")
	{
		regexstr := regexstr . GetCurrentSender() . ").*"
	}
	Else
	{
		regexstr := regexstr . name . ").*"
	}
	cleanstr := RegExReplace(email, regexstr) ; Initial clearance of anything past first detection of full name for basic sig removal
	cleanstr := RegExReplace(cleanstr, "\h+", " ") ; Cleaning of extraneus horizontal whitespace
	cleanstr := RegExReplace(cleanstr, "\v+", "`n") ; Cleaning of extraneous vertical whitespace
	cleanstr := Trim(cleanstr, whitespace) ; Removes trailing/preceding tabs, spaces, and returns.
	cleanstr := RegExReplace(cleanstr, "\s{2,}", "`n`n") ; Handling bulk mixed whitespace.
	cleanstr := RegExReplace(cleanstr, "([!+#^{}])","`{${1}`}") ; Wrap AHK symbols for safe pasting
	Return cleanstr
}

; Sets the Category of the email to mark that it was handled by me, and marks as read
SetAsHandled()
{
	email := GetCurrentEmail()
	If(email.Categories = "")
	{
		email.Categories := SubStr(A_UserName, 1, 2)
	}
	Else If(!InStr(email.Categories, SubStr(A_UserName, 1, 2)))
	{
		email.Categories := email.Categories . ", " . SubStr(A_UserName, 1, 2)
	}
	email.UnRead := False
	email.Save
	Return
}

; Returns the domain of an email address
GetEmailDomain(address:="")
{
	If(address = "")
	{
		address := GetCurrentEmail().SenderEmailAddress
	}
	RegExMatch(address, "(?<=@).*(?=\.)", &account)
	Return account[0]
}

; Returns a random boilerplate email greeting
GenerateGreeting()
{
	greetings := ["","Hi","Hey","Good $time"]
	hour := FormatTime(,"H")
	If (hour >= 7) and (hour < 12)
	{
		timestr := "morning"
	}
	Else If (hour >= 12) and (hour < 18)
	{
		timestr := "afternoon"
	}
	Else
	{
		timestr := "day" ; In case hour is outside of normal operating hours
	}
	idx := Random(1, greetings.Length)
	greet := greetings[idx]
	If (greet != "")
	{
		greet := greet . " "
	}
	Return StrReplace(greet, "$time", timestr)
}

; Sets my personal label onto a ticket
;; Originally made to make the task repeatable, however is currently only used for one hotkey
;; Will likely need to refactor or move actual code to hotkey
SetLabel(lbl:="")
{
	If(lbl = "")
	{
		lbl := A_UserName
	}
	WinActivate("RangerMSP")
	Click "300 100"
	WinWait "DatLabelSelectChkListFrm"
	try {
		ControlClick SubStr(lbl, 2)
	} catch {
		ControlClick "Simon Thurston"
	}
}

; Sets the labor type of a charge to be remote
SetRemoteLabor()
{
	ControlClick("TSOEdit1")
	Send "Labor{Space}+9R{Enter}{Tab 3}{Enter}"
}

; General Hotkeys

; Reloads Script
^!+r::Reload

; Pauses execution
^!+p::Pause

^!+d::Send GetEmailBody()

; VSCodium Hotkeys

#HotIf (WinActive("ahk_exe VSCodium.exe"))
{
	; Generates a Table of Contents using Headers of a Markdown File
	;; Should work fine for general applications but is untested outside of current usecase context
	^!m::
	{
		A_Clipboard := ""
		Send "^a^x"
		ClipWait
		instr := RegExReplace(A_Clipboard, "m)^(?!#+.*).*\R")
		instr := RegExReplace(instr, "^.*[^\s]")
		outstr := ""
		Loop Parse instr, "`n"
		{
			RegExReplace(A_LoopField, "#", "#", rCount)
			rCount -= 1
			tspace := ""
			Loop(%rCount%)
				tspace := tspace . A_Tab
			str := %A_LoopField%
			; Replace headings with bullets
			str := %RegExReplace(str, "#+", tspace . "*")%
			; Place headings into Markdown Link Format
			str := %RegExReplace(str, "(?<=\*\s)(.*)", "[$1](#$1)")%
			; Trim Newlines
			str := %RegExReplace(str, "m)\R")%
			; Condense filenames and hotkeys to links
			str := %RegExReplace(str, "[\.+](?!\S+\])")%
			; Replace spaces with plusses
			str := %RegExReplace(str, "#(\w+)\s(\w+)", "#$1+$2")%
			outstr := outstr . str . "`n"
		}
		out := ""
		RegExMatch(A_Clipboard, "m)^# (?!Table of Contents)(.*\R)*.*", out)
		A_Clipboard := "# Table of Contents`r`n" . outstr . out
		Send "^v"
		Return
	}
}

; Outlook Hotkeys

#HotIf (WinActive("ahk_exe OUTLOOK.EXE"))
{
	; Generates a reply email and autofills with standard greeting and email signature
	^r::
	{
		GetCurrentEmail().ReplyAll.Display
		WinWaitActive("RE: ")
		Send "!has2{Down 4}{Enter}{Up 2}"
		Send(GenerateGreeting() . GetFirstName() . ",")
		Send "{Enter 2}"
		Return
	}

	; Autoapplies email signature on new emails
	^n::Send "^n!nas2{Down 4}{Enter}"

	#HotIf (WinActive("Message \(.*\)"))
	{
		:*:spammail::I've got this blocked. You can delete the original message if you haven't already. Let us know if you receive any more or if you need any further assistance.
	}
}

; RangerMSP Hotkeys

#HotIf (WinActive("AHK_exe RangerMSP.exe"))
{
	#HotIf (WinActive("RangerMSP"))
	{
		; Creates a new ticket prefilling text
		^n::
		{
			Send "^n"
			WinWaitActive("ahk_class TDatNewSupportTicketsFrm")
			Send "{Tab 3}{F8}{Enter}"
			Return
		}
		; Calls the SetLabel Function
		^+l::SetLabel(InputBox("Label:",,"W150 H100").Value)

		; Creates a new charge and automatically calls SetRemoteLabor
		!+c::
		{
			Send "!+c"
			Sleep(500)
			WinWaitActive("New Charge")
			SetRemoteLabor()
			Return
		}
	}

	#HotIf (WinActive("New Ticket"))
	{
		; Autofills a new ticket with the contents of an email
		^e::
		{
			A_Clipboard := ""
			Send "^a^x"
			ClipWait
			A_Clipboard := StrReplace(A_Clipboard, "$User", GetCurrentSender())
			A_Clipboard := StrReplace(A_Clipboard, "$ContactType", "email")
			userAccount := SubStr(GetEmailDomain(), 1, 4)
			Send "^v"
			Send GetEmailBody()
			Send "+{Tab 3}"
			Send userAccount
			KeyWait "Enter", "D" 
			Send "{Enter}{Tab 3}"
			SetAsHandled()
			Return
		}

		; Prompts for input for quick phone triage
		^p::
		{
			A_Clipboard := ""
			userName := InputBox("Name:",,"W150 H100")
			userAccount := InputBox("Account:",,"W150 H100")
			WinActivate("New Ticket")
			Send "^a^x"
			ClipWait
			A_Clipboard := StrReplace(A_Clipboard, "$User", userName.Value)
			A_Clipboard := StrReplace(A_Clipboard, "$ContactType", "phone")
			Send "^v"
			Send "+{Tab 3}"
			Send userAccount.Value
			KeyWait "Enter", "D"
			Send "{Tab 3}"
			Return
		}
	}

	#HotIf (WinActive("New Charge - \(Labor\)"))
	{
		; Pastes email content into body of charge
		^e::
		{
			Send GetFirstName()
			Send "{Space}Emailed in:{Enter}"
			Send GetEmailBody()
			SetAsHandled()
			Return
		}

		; Manually calls SetRemoteLabor function in case of error
		^r::SetRemoteLabor()

		; Overwrites save charge button to automatically Round Down time
		^g::
		{
			ControlClick("TBitBtn1",,,,,"NA")
			WinWaitActive("New Charge")
			If(WinActive("ahk_class TMessageForm"))
			{
				ControlClick("TButton1",,,,,"NA")
			}
			While WinActive("ahk_class TDatSlipsDtlFrm")
			{
				ControlClick("TButton1",,,,,"NA")
			}
			Return
		}

		; Prompts for minutes spent on charge for quick input
		^t::
		{
			mins := InputBox("Minutes:",,150 100)
			KeyWait "Enter", "D"
			If(mins < 10)
			{
				mins := "0" . mins
			}
			mins := "00" . mins
			ControlSend "TAdrockDateTimeEdit1", %mins%
			ControlFocus "TCmtDBMemoValueSelect1" 
			Send "{Down 10}"
			Return
		}
	}

	#HotIf (WinActive("Timer"))
	{
		; Quick close of timer and creation of remote charge
		^Enter::
		{
			Click "230 50"
			WinWaitActive("New Charge - \(Labor\)",,500)
			Click "30 395"
			Sleep(100)
			SetRemoteLabor()
			Return
		}
	}
}