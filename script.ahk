#NoEnv
#SingleInstance force


GetWH(win_title, is_regex := False) {
    global
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    WinGetPos, _, _, w, h, %win_title%
    return % w != ""
}

ActiveTitle() {
    WinGetTitle, title, A
    return % title
}

WinActiveRegEx(win_title, is_regex := False) {
    SetTitleMatchMode, % is_regex ? "RegEx" : "1"
    return WinActive(win_title)
}
