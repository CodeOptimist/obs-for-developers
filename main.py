import json, re
if __name__ != '__main__':
    import obspython as obs
from autohotkey import Script, AhkExitException
from datetime import datetime
from collections import namedtuple
from win32api import OutputDebugString


ahk = Script()
AhkWindow = namedtuple('AhkWindow', ['ahk_title', 'title', 'class_', 'exe', 'is_regex'])
video_info = None
scene = None

windows = {
    'Visual Studio': {'spec': "RimMods - Microsoft Visual Studio:HwndWrapper[DefaultDomain;;f1776b62-97a2-4920-9344-4a8e003b5404]:devenv.exe"},
    'dnSpy': {'spec': "dnSpy v6.0.5 (64-bit):HwndWrapper[dnSpy.exe;;dcb937d0-a05d-4507-8a73-c965d495a0ce]:dnSpy.exe"},
    'TortoiseHg Workbench': {'spec': "JobsOfOpportunity - TortoiseHg Workbench:Qt5QWindowIcon:thgw.exe"},
    'TortoiseHg Commit': {'spec': "JobsOfOpportunity - commit:Qt5QWindowIcon:thgw.exe"},
    'RimWorld': {'spec': "RimWorld by Ludeon Studios:UnityWndClass:RimWorldWin64.exe"},
    'BC Text Compare': {'spec': "/.*@.* <--> .*@.* - Text Compare - Beyond Compare/:TViewForm:BCompare.exe"},
}


def log(func):
    def wrapper(*args, **kwargs):
        before = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after = obs.bnum_allocs()
        print("{} {} {}() {}".format(datetime.now().strftime("%I:%M:%S:%f"), before, func.__name__, after))
        return result
    return wrapper


def update_source(source, spec, cond=None):
    data = obs.obs_save_source(source)
    source_info = json.loads(obs.obs_data_get_json(data))
    obs.obs_data_release(data)
    if cond is None or cond(source_info):
        source_info['settings']['window'] = spec
        new_data = obs.obs_data_create_from_json(json.dumps(source_info['settings']))
        obs.obs_source_update(source, new_data)
        obs.obs_data_release(new_data)


def timer():
    try:
        for name, window in windows.items():
            ahk_window = get_ahk_window(window['spec'])

            if ahk.f('WinActiveRegEx', ahk_window.ahk_title, ahk_window.is_regex):
                scene_item = obs.obs_scene_find_sceneitem_by_id(scene, window['sceneitem_id'])
                source = obs.obs_sceneitem_get_source(scene_item)

                ahk.call('ActiveWin')
                spec = ":".join([(ahk.get('title').replace(':', '#3A')), ahk.get('class'), ahk_window.exe])
                update_source(source, cond=lambda source_info: source_info['settings']['window'] != spec, spec=spec)

                # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
                state = ahk.get('state')
                if state != window.get('state', state):
                    obs.timer_add(lambda: update_source(source, spec=spec) or obs.remove_current_callback(), 2500)
                window['state'] = state

                # obs.obs_source_release(source)
                center_item(scene_item, ahk_window)
                obs.obs_sceneitem_set_order_position(scene_item, len(windows) - 1)
                # obs.obs_sceneitem_release(scene_item)
    except AhkExitException as ex:
        obs.timer_remove(timer)
        print(ex)
        return


def get_ahk_window(window_spec):
    title, class_, exe = window_spec.split(':')
    regex = re.compile(r'#[\dA-F]{2}')
    # replace things like #3A with :
    new_title = regex.sub(lambda m: m.group().replace(m.group(), bytes.fromhex(m.group()[1:]).decode('utf-8')), title)
    is_regex = title.startswith('/') and title.endswith('/')
    if is_regex:
        new_title = new_title[1:-1]
    ahk_title = new_title + " ahk_exe " + exe
    return AhkWindow(ahk_title, title, class_, exe, is_regex)


def center_item(item, ahk_window):
    if not ahk.f('GetWH', ahk_window.ahk_title, ahk_window.is_regex):
        print("Couldn't find {ahk_window}".format(**locals()))
        return

    vec2 = obs.vec2()
    vec2.x = video_info.base_width / 2 - ahk.get('w') / 2
    vec2.y = video_info.base_height / 2 - ahk.get('h') / 2
    obs.obs_sceneitem_set_pos(item, vec2)


@log
def script_load(settings):
    global ahk, scene, video_info
    if ahk is not None:
        ahk.exit()

    try:
        ahk = Script.from_file(r'script.ahk')
    except Exception as ex:
        print(ex)
        return

    video_info = obs.obs_video_info()
    obs.obs_get_video_info(video_info)

    @log
    def wipe_scene():
        items = obs.obs_scene_enum_items(scene)
        for item in items:
            obs.obs_sceneitem_remove(item)
            source = obs.obs_sceneitem_get_source(item)
            obs.obs_source_remove(source)
            obs.obs_source_release(source)
            obs.obs_sceneitem_release(item)
        obs.sceneitem_list_release(items)

    scene = obs.obs_scene_from_source(obs.obs_get_source_by_name("Scene 1"))
    wipe_scene()

    for name, window in windows.items():
        source_info = {'id': 'window_capture', 'name': name, 'settings': {'method': 2, 'priority': 1, 'window': window['spec']}}
        data = obs.obs_data_create_from_json(json.dumps(source_info))
        source = obs.obs_load_source(data)
        obs.obs_data_release(data)

        scene_item = obs.obs_scene_add(scene, source)
        obs.obs_source_release(source)
        window['sceneitem_id'] = obs.obs_sceneitem_get_id(scene_item)
        obs.obs_sceneitem_set_locked(scene_item, True)

        ahk_window = get_ahk_window(window['spec'])
        center_item(scene_item, ahk_window)
        obs.obs_sceneitem_release(scene_item)

    obs.timer_add(timer, 50)


@log
def script_unload():
    obs.obs_scene_release(scene)
    ahk.exit()


def script_description():
    return "Python & AutoHotkey powered OBS."

