# Table of Contents

* [AutoHotkey Scripts](#AutoHotkey+Scripts)
	* [MasterScript.ahk](#MasterScriptahk)
		* [Ctrl+Alt+Shift+r](#CtrlAltShiftr)
		* [Ctrl+Alt+Shift+q](#CtrlAltShiftq)
	* [EmailGreetings.ahk](#EmailGreetingsahk)
		* [Ctrl+r](#Ctrlr)
	* [TicketTemplate.ahk](#TicketTemplateahk)
		* [SetLabel()](#SetLabel())
		* [Ctrl+n](#Ctrln)
		* [Ctrl+l](#Ctrll)
		* [Ctrl+Shift+e](#CtrlShifte)
		* [Ctrl+Shift+p](#CtrlShiftp)
		* [Ctrl+g](#Ctrlg)
		* [Ctrl+e](#Ctrle)
	* [MDTableOfContents.ahk](#MDTableOfContentsahk)
		* [Ctrl+Alt+m](#CtrlAltm)
	* [EmailFuncs](#EmailFuncs)
		* [GetCurrentEmail()](#GetCurrentEmail())
		* [GetSender()](#GetSender())
		* [GetStandardName(name)](#GetStandardName(name))
		* [GetFirstName(name)](#GetFirstName(name))
		* [GetEmailBody(email,name)](#GetEmailBody(email,name))
		* [GenerateGreeting()](#GenerateGreeting())

# AutoHotkey Scripts

I feel like just about everyone gets sick and tired of doing things by hand, so what better way to stick it to my poor, poor, RSI riddled hands than to automate everything and put them out of a job! Take THAT, bones!

Anyways I'll keep a general description of what each script is and what hotkeys it has here. In case you want to steal them for yourself,

## MasterScript.ahk

This is simply a launcher file I use to have all of my scripts run at startup. There's nothing to this one.

### Ctrl+Alt+Shift+r

This triggers a force reload of all scripts, to update them with edits on the fly.

### Ctrl+Alt+Shift+q

Closes all scripts except for the Master Script, in case of anything catastrophic.

## EmailGreetings.ahk

This script is one I use to automate boring email stuff. Because Outlook doesn't let you have any fun.

### Ctrl+r

This is meant to go on top of the reply shortcut. On press, it will pass through the shortcut to Outlook, opening the editor pane to reply to a message. It will then automatically copy the info from the "to" line, parse a regex on it extracting the first name, then paste it back into the body of the editor with a boilerplate greeting attached. This, combined with a signature that already includes a boilerplate sign-off, makes it so that I only have to focus on writing boilerplate body messages.

## TicketTemplate.ahk

I use Ranger MSP (Previously Commit) at work as the primary ticketing software. Having to set up tickets manually all the time is a pain, and when someone calls in, you have to be quick with getting info in. These are a bunch of quick scripts to help speed up data entry.

### SetLabel()

Automatically sets my label on a ticket. Turned into function due to frequency in which this is called.

### Ctrl+n

Creates a new ticket in the system, and automatically fills out a quick boilerplate header. To be used in combination with other inputs to generate full tickets.

### Ctrl+l

Calls the SetLabel() function from the main Ranger tickets screen.

### Ctrl+Shift+e

Pulls all the relevant info from a client who emails in, scraping it directly from Outlook. This is used in combination with an Outlook Macro to copy the sender name to the clipboard to be inserted back into Ranger.

### Ctrl+Shift+p

Gives prompts for Triaging info from phone calls and filling out the ticket with them.

### Ctrl+g

Overwrites the regular save action on the new ticket screen to automatically call the SetLabel() function on save.

### Ctrl+e

Automatically pulls in info from emails to add as charges quickly.

## MDTableOfContents.ahk

What I used to make the Table of Contents for this file! If you want a nice Table of Contents for your readme, go ahead and use this!

### Ctrl+Alt+m

Will Select All and Cut the current Window Text, grab all markdown headers from it, then build a Table of Contents from said headers, pasting it back in at the beginning of the document.

## EmailFuncs

### GetCurrentEmail()

Grabs the currently selected email within Outlook and returns it as a Mailitem VBA object via OLE.

### GetSender()

Returns the Full Name of the sender of the current selected email.

### GetStandardName(name)

When given a full name as an input (i.e. including Office formatted middle initials), will strip it down to FirstName[[:space:]]LastName for the purposes of data sanitization.

### GetFirstName(name)

Cuts out the Middle Initial and/or the Last Name of the given name to return on the the First Name of the sender.

### GetEmailBody(email,name)

Uses the name of the sender of an email to crop out the email signature at the earliest point possible.

### GenerateGreeting()

Returns a random greeting to use when creating new emails.
