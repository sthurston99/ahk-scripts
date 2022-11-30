#NoEnv
#SingleInstance Force
#If (WinActive("ahk_exe VSCodium.exe"))
{
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