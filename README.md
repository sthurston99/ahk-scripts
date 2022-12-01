# AutoHotkey Scripts

I do a lot of my work with a keyboard. And because of that, I do a lot of repetetive things with a keyboard. To reduce the amount of keystrokes I have to send, I made this script (which was originally several scripts but that was getting annoying) to simplify a lot of the actions I needed to perform. And now this thing automates a good amount of the documentation I have to do while working, which, is pretty sweet. Here's a brief summary of all the cool stuff that this script does, and the reasoning behind it.

Also, a lot of this script is heavily specific to me, my workflow, and the details within my environment. It should only take a little bit of tooling to make it fit what you're doing, but if you're modifying this, go over it with a fine tooth comb and test everything before actually doing live-fire testing on it. I've almost caused so many problems by not running this in a test environment during the initial stages of development, and you never want to mess with how bad AHK can mess up your situation.

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

### SetLabel()

tl;dr Sets my personal label onto a ticket

Uses Control Hooks within Ranger to automatically label a ticket as mine

### SetRemoteLabor()

tl;dr Sets the labor type of a charge to be remote

Quickly adjusts the labor type of a ticket to Remote when focused on the charge description window.

## General Hotkeys

## VSCodium

## Outlook

## RangerMSP