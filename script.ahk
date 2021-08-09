#NoEnv
#SingleInstance force


WinGetWH(win_title, is_regex) {
    global
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    WinGetPos, _, _, w, h, % win_title
    return w != ""
}

ActiveWinGet() {
    global
    WinGetTitle, title, A
    WinGetClass, class, A
    WinGet, exe, ProcessName, A
    WinGet, state, MinMax, A
}

WinActiveRegEx(win_title, is_regex) {
    SetTitleMatchMode, % is_regex ? "RegEx" : "3"
    return WinActive(win_title)
}
