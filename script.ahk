#NoEnv
#SingleInstance force

AutoExec() {
    SetTitleMatchMode, RegEx
    SetBatchLines, -1
}

WinGetWH(wintitle) {
    global
    WinGetPos, _, _, w, h, % wintitle
    return w != ""
}

GetMatchWinTitle(wintitle) {
    global
    WinGetTitle, title, % wintitle
    WinGetClass, class, % wintitle
    WinGet, exe, ProcessName, % wintitle
}

GetWindowsCached(wintitles) {
    global _wintitles, windows
    _wintitles := wintitles
    SetTimer, GetWindows, -0
    return windows  ; return version immediately
}

GetWindows() {
    global _wintitles, windows
    result := ""
    Loop, Parse, _wintitles, `n
    {
        patternIdx := A_Index - 1
        WinGet, idList, List, % A_LoopField
        Loop, % idList
        {
            wintitle := "ahk_id " idList%A_Index%
            result .= patternIdx "`r"          ; pattern_idx
            result .= WinExist(wintitle) "`r"  ; exists
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
