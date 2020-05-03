#NoEnv
#SingleInstance force


WinGetWH(win_title, is_regex := False) {
    global
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    WinGetPos, _, _, w, h, %win_title%
    return % w != ""
}

ActiveWinGet() {
    global
    WinGetTitle, title, A
    WinGetClass, class, A
    WinGet, state, MinMax, A
}

WinActiveRegEx(win_title, is_regex := False) {
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    return WinActive(win_title)
}
