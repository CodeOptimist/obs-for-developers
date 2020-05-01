import json, re
if __name__ != '__main__':
    import obspython as obs
from autohotkey import Script
from datetime import datetime
from collections import namedtuple


ahk = Script()
AhkWindow = namedtuple('Window', ['ahk_title', 'title', 'class_', 'exe'])
video_info = None
scene = None
source_jsons = dict()
window_sceneitem_ids = dict()


# scoped to avoid unwanted global variables
def init():
    obs_specs = {
        'Visual Studio': "RimMods - Microsoft Visual Studio:HwndWrapper[DefaultDomain;;f1776b62-97a2-4920-9344-4a8e003b5404]:devenv.exe",
        'dnSpy': "dnSpy v6.0.5 (64-bit):HwndWrapper[dnSpy.exe;;dcb937d0-a05d-4507-8a73-c965d495a0ce]:dnSpy.exe",
        'TortoiseHg Workbench': "JobsOfOpportunity - TortoiseHg Workbench:Qt5QWindowIcon:thgw.exe",
        'TortoiseHg Commit': "JobsOfOpportunity - commit:Qt5QWindowIcon:thgw.exe",
        'RimWorld': "RimWorld by Ludeon Studios:UnityWndClass:RimWorldWin64.exe",
        'BC Text Compare': ":TViewForm:BCompare.exe",
    }

    for name, obs_spec in obs_specs.items():
        source_jsons[name] = {'id': 'window_capture', 'name': name, 'settings': {'method': 2, 'priority': 1, 'window': obs_specs[name]}}


def log(func):
    def wrapper(*args, **kwargs):
        before = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after = obs.bnum_allocs()
        print("{} {} {}() {}".format(datetime.now().strftime("%I:%M:%S:%f"), before, func.__name__, after))
        return result
    return wrapper


def timer():
    for name, source_json in source_jsons.items():
        ahk_window = get_ahk_title(source_json)
        if ahk.f('WinActive', ahk_window.ahk_title):
            scene_item = obs.obs_scene_find_sceneitem_by_id(scene, window_sceneitem_ids[name])
            source = obs.obs_sceneitem_get_source(scene_item)
            title = ahk.f('ActiveTitle')
            obs_title = title.replace(':', '#3A')
            new_spec = ":".join([obs_title, ahk_window.class_, ahk_window.exe])

            data = obs.obs_save_source(source)
            source_json = json.loads(obs.obs_data_get_json(data))
            obs.obs_data_release(data)
            if source_json['settings']['window'] != new_spec:
                source_json['settings']['window'] = new_spec
                new_data = obs.obs_data_create_from_json(json.dumps(source_json['settings']))
                obs.obs_source_update(source, new_data)
                obs.obs_data_release(new_data)

            #obs.obs_source_release(source)
            center_item(scene_item, ahk_window)
            obs.obs_sceneitem_set_order_position(scene_item, len(source_jsons) - 1)
            #obs.obs_sceneitem_release(scene_item)


def get_ahk_title(window_source):
    settings_ = window_source['settings']
    title, class_, exe = settings_['window'].split(':')
    regex = re.compile(r'#[\dA-F]{2}')
    # replace things like #3A with :
    new_title = regex.sub(lambda m: m.group().replace(m.group(), bytes.fromhex(m.group()[1:]).decode('utf-8')), title)
    ahk_title = new_title if title != "" else "ahk_class " + class_ if class_ != "" else "ahk_exe " + exe
    return AhkWindow(ahk_title, title, class_, exe)


def center_item(item, ahk_window):
    if not ahk.f('GetWH', ahk_window.ahk_title):
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
    ahk = Script.from_file(r'script.ahk')

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

    for name, source_json in source_jsons.items():
        data = obs.obs_data_create_from_json(json.dumps(source_json))
        source = obs.obs_load_source(data)
        obs.obs_data_release(data)
        scene_item = obs.obs_scene_add(scene, source)
        obs.obs_source_release(source)
        window_sceneitem_ids[name] = obs.obs_sceneitem_get_id(scene_item)
        obs.obs_sceneitem_set_locked(scene_item, True)

        ahk_window = get_ahk_title(source_json)
        center_item(scene_item, ahk_window)
        obs.obs_sceneitem_release(scene_item)

    obs.timer_add(timer, 1000)


@log
def script_unload():
    obs.obs_scene_release(scene)
    ahk.exit()


def script_description():
    return "Python & AutoHotkey powered OBS."


init()
