# Table of Contents

* [AutoHotkey Scripts](#AutoHotkey+Scripts)
	* [Functions](#Functions)
		* [GetCurrentEmail](#GetCurrentEmail)
		* [GetSender](#GetSender)
		* [GetCurrentSender](#GetCurrentSender)
		* [GetStandardName](#GetStandardName)
		* [GetFirstName](#GetFirstName)
		* [GetEmailBody](#GetEmailBody)
		* [SetAsHandled](#SetAsHandled)
		* [GetEmailDomain](#GetEmailDomain)
		* [GenerateGreeting](#GenerateGreeting)
		* [SetLabel](#SetLabel)
		* [SetRemoteLabor](#SetRemoteLabor)
	* [General Hotkeys](#General+Hotkeys)
		* [Ctrl+Alt+Shift+r](#CtrlAltShiftr)
		* [Ctrl+Alt+Shift+p](#CtrlAltShiftp)
	* [VSCodium](#VSCodium)
		* [Alt+Shift+m](#AltShiftm)
	* [Outlook](#Outlook)
		* [Ctrl+r](#Ctrlr)
		* [Ctrl+n](#Ctrln)
	* [RangerMSP](#RangerMSP)
		* [Main Window](#Main+Window)
			* [Ctrl+n](#Ctrln)
			* [Ctrl+Shift+l](#CtrlShiftl)
			* [Alt+Shift+c](#AltShiftc)
		* [New Ticket](#New+Ticket)
			* [Ctrl+e](#Ctrle)
			* [Ctrl+p](#Ctrlp)
		* [New Charge](#New+Charge)
			* [Ctrl+e](#Ctrle)
			* [Ctrl+r](#Ctrlr)
			* [Ctrl+g](#Ctrlg)
			* [Ctrl+t](#Ctrlt)
		* [Timer](#Timer)
			* [Ctrl+Enter](#CtrlEnter)

# AutoHotkey Scripts

I do a lot of my work with a keyboard. And because of that, I do a lot of repetetive things with a keyboard. To reduce the amount of keystrokes I have to send, I made this script (which was originally several scripts but that was getting annoying) to simplify a lot of the actions I needed to perform. And now this thing automates a good amount of the documentation I have to do while working, which, is pretty sweet. Here's a brief summary of all the cool stuff that this script does, and the reasoning behind it.

Also, a lot of this script is heavily specific to me, my workflow, and the details within my environment. It should only take a little bit of tooling to make it fit what you're doing, but if you're modifying this, go over it with a fine tooth comb and test everything before actually doing live-fire testing on it. I've almost caused so many problems by not running this in a test environment during the initial stages of development, and you never want to mess with how bad AHK can mess up your situation.

**The main three pieces to modify to fit this to your own workflow are: the [SetLabel](#SetLabel) and [SetAsHandled](#SetAsHandled) functions, and the hotkeys that exist for [Outlook](#Outlook).**

These are the 3 places that I know for sure have components that are specific to setting things *as me*. As in, setting my email signature, setting my label, and setting my email category.

The rest of it is up to you to tool into whatever uses you need. Have fun.

## Functions

As with any software project, you need to have a lot of base functions as the building blocks for automation. Luckily, with hotkeys, you don't actually need as many functions, since many of the actions you need to perform per-hotkey are unique. However, that does not completely let you off when a lot of what you're doing is similar, so I've created a small set of functions that perform some of the more repetetive tasks, and built them in a way that they are hopefully extensible to future applications.

### GetCurrentEmail

tl;dr Fetches current selected email within Outlook

This uses the windows COM interface to fetch the currently selected email in Outlook. It returns it as the standard MailItem object so that further manipulation can be done to it.

### GetSender

tl;dr Fetches the Full Name of the sender of the COM MailItem Object

Simply takes a MailItem object as input and returns its SenderName property.

### GetCurrentSender

tl;dr Wrapper for GetSender to get Sender from Current Email

Calls GetSender with the parameter of GetCurrentEmail to return the current email's sender.

### GetStandardName

tl;dr Converts a Full Name into a Standard Name by removing the Middle Initial

Takes a full name (First Name, Middle Initial/Name, Last Name) and removes the middle initial from the string.

### GetFirstName

tl;dr Extracts the first name from a Standard Name

Takes a standardized name (Firstname Lastname) and crops it to just the first name.

### GetEmailBody

tl;dr Trims out the contents of an email after the name provided

Takes an email body and a senders name as input, and constructs a regex string to remove all contents of an email after the senders name is detected as a hacky way to remove email signatures.

### SetAsHandled

tl;dr Sets the Category of the email to mark that it was handled by me, and marks as read

Sends a few commands to the current email in Outlook to set as read with my initials as the category within our shared email.

### GetEmailDomain

tl;dr Returns the domain of an email address

Takes an email string as input, returns the top level subdomain as the output.

### GenerateGreeting

tl;dr Returns a random boilerplate email greeting

Pulls a random greeting from a preset list to be used at the opening of an email. Includes a generic timeframe determination script to determine whether it's morning/afternoon for the sake of filling out time-based greetings.

### SetLabel

tl;dr Sets my personal label onto a ticket

Uses Control Hooks within Ranger to automatically label a ticket as mine

### SetRemoteLabor

tl;dr Sets the labor type of a charge to be remote

Quickly adjusts the labor type of a ticket to Remote when focused on the charge description window.

## General Hotkeys

These Hotkeys are mostly about managing the script and the execution thereof itself. These aren't super special, so here they are in brief:

### Ctrl+Alt+Shift+r

Reloads the script to refresh changes quickly

### Ctrl+Alt+Shift+p

Pauses the current script's execution for quick error management

## VSCodium

My primary text editor while at work is VSCodium. It's pretty handy but does not actually do everything magically. These scripts are mostly for automating some text processing that I do a lot within it.

### Alt+Shift+m

Given a Markdown file with Headers, generates a Table of Context from said headers, then places it at the opening of the file.

## Outlook

Most businesses use Microsoft Outlook as their email program. We also have to deal with it here. There's a lot of repetetive nonsense that Outlook just doesn't handle by default, so I've added a few things to automate them.

### Ctrl+r

Rebinds Reply to Reply All, and automatically prefills a standard greeting and applies my signature.

### Ctrl+n

Automatically applies my email signature to any new email.

## RangerMSP

Our ticketing software at my work is RangerMSP, formerly Commit. It's got plenty of quirks, but it does have a pretty good structure for hooking into with automations. So, I've done that, a lot, to make my life just a smidge easier when having to document all of my stuff.

### Main Window

My primary focused Window during the day is the main Ticket List, so many of the shortcuts here have to do with creating or modifying tickets in some way.

#### Ctrl+n

Creates a new ticket and automatically prefills the body text from the F8 menu, using the boilerplate ticket text I have stored there.

#### Ctrl+Shift+l

Applies my label to the currently selected ticket from the Label dropdown.

#### Alt+Shift+c

When making a new charge, Automatically applies the Remote Labor labor type, as that is the primary type used in my day-to-day work.

### New Ticket

These hotkeys all run when the New Ticket window is open and focused, and they allow you to put in ticket triage details quickly to move onto actual troubleshooting.

#### Ctrl+e

Calls GetEmailBody to scrape info from an email into a ticket to prefill a ticket.

#### Ctrl+p

Gives 2 input prompts for Name and Account to prefill basic info before triaging the issue.

### New Charge

By far the most time I spend in Ranger is inputting and editing charges. There's plenty of tiny little annoying things that you have to do, and my goal is to get rid of most if not all of them.

#### Ctrl+e

Pulls email info via GetEmailBody into a charge.

#### Ctrl+r

Quickly sets labor type as remote in case most of the other charge creation methods that call it fail.

#### Ctrl+g

Rounds down time when saving a charge. Does not handle all error prompts automatically, but will usually handle at least one of them.

#### Ctrl+t

Adds a quick text prompt that allows you to fill in the amount of minutes you spent on a charge.

### Timer

The Timer function in Ranger is usually a pretty good tool However, the window it opens for it is almost always terrible to use with keyboard-focused inputs. So I made these scripts to make it better.

#### Ctrl+Enter

Shortcut to automatically create charge from timer.
