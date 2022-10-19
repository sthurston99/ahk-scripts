# AutoHotkey Scripts

I feel like just about everyone gets sick and tired of doing things by hand, so what better way to stick it to my poor, poor, RSI riddled hands than to automate everything and put them out of a job! Take THAT, bones!

Anyways I'll keep a general description of what each script is and what hotkeys it has here. In case you want to steal them for yourself,

## MasterScript.ahk

This is simply a launcher file I use to have all of my scripts run at startup. There's nothing to this one.

## EmailGreetings.ahk

This script is one I use to automate boring email stuff. Because Outlook doesn't let you have any fun.

### Ctrl+r

This is meant to go on top of the reply shortcut. On press, it will pass through the shortcut to Outlook, opening the editor pane to reply to a message. It will then automatically copy the info from the "to" line, parse a regex on it extracting the first name, then paste it back into the body of the editor with a boilerplate greeting attached. This, combined with a signature that already includes a boilerplate sign-off, makes it so that I only have to focus on writing boilerplate body messages.