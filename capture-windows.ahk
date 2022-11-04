#NoEnv
#SingleInstance force

AutoExec() {
    SetTitleMatchMode, RegEx
    SetBatchLines, -1
}

GetWindowsCached() {
    global windows
    SetTimer, GetWindows, -0
    return windows
}

GetWindows() {
    global wintitles, windows
    result := ""
    Loop, Parse, wintitles, `n
    {
        patternIdx := A_Index - 1
        WinGet, idList, List, % A_LoopField
        Loop, % idList
        {
            wintitle := "ahk_id " idList%A_Index%
            result .= patternIdx "`r"          ; pattern_idx
            result .= WinExist(wintitle) "`r"  ; id
            result .= WinActive(wintitle) "`r" ; focused
            WinGetTitle, title, % wintitle
            result .= title "`r"               ; title
            WinGetClass, class, % wintitle
            result .= class "`r"               ; class
            WinGet, exe, ProcessName, % wintitle
            result .= exe "`r"                 ; exe
            result .= "`n"
        }
    }
    windows := result  ; for GetWindowsCached() next time
    return result  ; in case we call GetWindows() directly
}
