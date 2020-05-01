#NoEnv
#SingleInstance force
Init()

Init() {
}

GetWH(win_title) {
    global
    WinGetPos, _, _, w, h, %win_title%
    return % w != ""
}

ActiveTitle() {
    WinGetTitle, title, A
    return % title
}