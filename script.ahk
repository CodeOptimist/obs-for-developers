#NoEnv
#SingleInstance force

AutoExec() {
    SetTitleMatchMode, RegEx
}

WinGetWH(win_title) {
    global
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

WinActiveRegEx(win_title) {
    return WinActive(win_title)
}
