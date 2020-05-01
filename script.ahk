#NoEnv
#SingleInstance force


GetWH(win_title, is_regex := False) {
    global
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    WinGetPos, _, _, w, h, %win_title%
    return % w != ""
}

ActiveWin() {
    global
    WinGetTitle, title, A
    WinGetClass, class, A
    WinGet, state, MinMax, A
}

WinActiveRegEx(win_title, is_regex := False) {
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    return WinActive(win_title)
}
