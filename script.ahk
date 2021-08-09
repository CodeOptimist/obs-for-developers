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

WinGet(win_title) {
    global
    WinGetTitle, title, % win_title
    WinGetClass, class, % win_title
    WinGet, exe, ProcessName, % win_title
    WinGet, state, MinMax, % win_title
}
