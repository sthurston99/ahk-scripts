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
            str := % RegExReplace(str, "#+", tspace . "*")
            str := % RegExReplace(str, "(?<=\*\s)(.*)", "[$1](#$1)")
            str := % RegExReplace(str, "m)\R")
            str := % RegExReplace(str, "[\.+](?!\S+\])")
            str := % RegExReplace(str, "#(\w+)\s(\w+)#", "#$1+$2")
            outstr := outstr . str . "`n"
        }
        RegExMatch(clipboard, "m)^# (?!Table of Contents)(.*\R)*.*", out)
        clipboard := "# Table of Contents`r`n" . outstr . out
        Send, ^v
    return
}