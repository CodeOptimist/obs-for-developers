import json, re, yaml, os
if __name__ != '__main__':
    import obspython as obs
from autohotkey import Script, AhkExitException
from datetime import datetime
from collections import namedtuple
from win32api import OutputDebugString


os.chdir(r'C:\Dropbox\Python\obs')
ahk = Script()
AhkWindow = namedtuple('AhkWindow', ['ahk_title', 'title', 'class_', 'exe', 'is_regex'])
video_info = None
loaded = None
windows = None


def init():
    global loaded, windows
    with open(r'captures.yaml', encoding='utf-8') as f:
        loaded = yaml.safe_load(f)
    #windows = loaded['scenes']['Jobs of Opportunity']


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
                sceneitem = obs.obs_scene_find_sceneitem_by_id(scene, window['sceneitem_id'])
                source = obs.obs_sceneitem_get_source(sceneitem)

                ahk.call('ActiveWin')
                spec = ":".join([(ahk.get('title').replace(':', '#3A')), ahk.get('class'), ahk_window.exe])
                update_source(source, cond=lambda source_info: source_info['settings']['window'] != spec, spec=spec)

                # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
                state = ahk.get('state')
                if state != window.get('state', state):
                    obs.timer_add(lambda: update_source(source, spec=spec) or obs.remove_current_callback(), 2500)
                window['state'] = state

                center_item(sceneitem, ahk_window)
                obs.obs_sceneitem_set_order_position(sceneitem, len(windows) - 1)
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
def wipe_scenes():
    sources = obs.obs_frontend_get_scenes()
    for idx, source in enumerate(sources):
        if idx == 0:
            continue
        obs.obs_source_remove(source)
    obs.source_list_release(sources)


@log
def get_scene_by_name(scene_name):
    try:
        sources = obs.obs_frontend_get_scenes()
        for source in sources:
            name = obs.obs_source_get_name(source)
            if name == scene_name:
                return obs.obs_scene_from_source(source)
        return None
    finally:
        obs.source_list_release(sources)


@log
def wipe_scene(scene):
    items = obs.obs_scene_enum_items(scene)
    for item in items:
        obs.obs_sceneitem_remove(item)
        source = obs.obs_sceneitem_get_source(item)
        obs.obs_source_remove(source)
    obs.sceneitem_list_release(items)


@log
def scenes_loaded():
    scene_names = obs.obs_frontend_get_scene_names()
    if not scene_names:
        return
    obs.remove_current_callback()

    for scene_name, scene_windows in loaded['scenes'].items():
        scene = get_scene_by_name(scene_name)

        create_scene = scene is None
        if create_scene:
            scene = obs.obs_scene_create(scene_name)
        else:
            wipe_scene(scene)

        for window_name, window in scene_windows.items():
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {'method': 2, 'priority': 1, 'window': window['spec']}}
            data = obs.obs_data_create_from_json(json.dumps(source_info))
            source = obs.obs_load_source(data)
            obs.obs_data_release(data)

            sceneitem = obs.obs_scene_add(scene, source)
            obs.obs_source_release(source)
            window['sceneitem_id'] = obs.obs_sceneitem_get_id(sceneitem)
            obs.obs_sceneitem_set_locked(sceneitem, True)

            ahk_window = get_ahk_window(window['spec'])
            center_item(sceneitem, ahk_window)

        if create_scene:
            obs.obs_scene_release(scene)

    #obs.timer_add(timer, 50)


@log
def script_load(settings):
    global ahk, video_info
    init()

    if ahk is not None:
        ahk.exit()

    try:
        ahk = Script.from_file(r'script.ahk')
    except Exception as ex:
        print(ex)
        return

    video_info = obs.obs_video_info()
    obs.obs_get_video_info(video_info)

    obs.timer_add(scenes_loaded, 50)




@log
def script_unload():
    ahk.exit()


def script_description():
    return "Python & AutoHotkey powered OBS."


if __name__ == '__main__':
    init()
