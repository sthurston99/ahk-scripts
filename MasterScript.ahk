#SingleInstance Force

AHKPath:="C:\Program Files\AutoHotkey\AutoHotkey.exe"

Loop %A_WorkingDir%\*
{
    If (InStr(A_LoopFileName, "ahk") and not InStr(A_LoopFileName, A_ScriptName))
    {
        Run %AHKPath% %A_LoopFileName%
        ; FAILSAFE IF LOOP FAILS
        If A_Index > 30
            return
    }
}



^!+r::Reload
^!+q::
    WinGet, pList, List, ahk_class AutoHotkey
    Loop %pList%
        MsgBox, 'pList%A_Index%'

    Loop %A_WorkingDir%\*
    {
        If (InStr(A_LoopFileName, "ahk") and not InStr(A_LoopFileName, A_ScriptName))
        {
            WinKill, %A_LoopFileName%
            ; FAILSAFE IF LOOP FAILS
            If A_Index > 30
                return
        }
    }