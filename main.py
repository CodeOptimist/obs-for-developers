import json, re, yaml, os
if __name__ != '__main__':
    import obspython as obs
from autohotkey import Script, AhkExitException
from datetime import datetime
from collections import namedtuple
from win32api import OutputDebugString

os.chdir(r'C:\Dropbox\Python\obs')
# globals are reset on reload
AhkWindowSpec = namedtuple('AhkWindowSpec', ['win_title', 'is_re', 'title', 'class_', 'exe'])
ahk = None
video_info = None
loaded = None


def init():
    global loaded
    with open(r'captures.yaml', encoding='utf-8') as f:
        loaded = yaml.safe_load(f)


def log(func):
    def wrapper(*args, **kwargs):
        before = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after = obs.bnum_allocs()
        print("{} {} {}() {}".format(datetime.now().strftime("%I:%M:%S:%f"), before, func.__name__, after))
        return result

    return wrapper


def update_source(source, obs_spec, cond=None):
    data = obs.obs_save_source(source)
    source_info = json.loads(obs.obs_data_get_json(data))
    obs.obs_data_release(data)
    if cond is None or cond(source_info):
        source_info['settings']['window'] = obs_spec
        new_data = obs.obs_data_create_from_json(json.dumps(source_info['settings']))
        obs.obs_source_update(source, new_data)
        obs.obs_data_release(new_data)


def timer():
    try:
        update_active_win_sources()
    except AhkExitException:
        obs.remove_current_callback()
        raise


def update_active_win_sources():
    cur_scene_source = obs.obs_frontend_get_current_scene()
    cur_scene = obs.obs_scene_from_source(cur_scene_source)
    cur_scene_name = obs.obs_source_get_name(cur_scene_source)
    obs.obs_source_release(cur_scene_source)
    # print("Current scene: {cur_scene_name}".format(**locals()))

    windows = loaded['scenes'].get(cur_scene_name, {})
    for scene_window_name, scene_window in windows.items():
        ahk_window = loaded['ahk_windows'][scene_window_name]
        ahk_spec = ahk_window['spec']

        if ahk.f('WinActiveRegEx', ahk_spec.win_title, ahk_spec.is_re):
            sceneitem = obs.obs_scene_find_sceneitem_by_id(cur_scene, scene_window['sceneitem_id'])
            source = obs.obs_sceneitem_get_source(sceneitem)

            ahk.call('ActiveWinGet')
            obs_spec = ":".join([(ahk.get('title').replace(':', '#3A')), ahk.get('class'), ahk_spec.exe])
            update_source(source, obs_spec, cond=lambda source_info: source_info['settings']['window'] != obs_spec)

            # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
            state = ahk.get('state')
            if state != ahk_window.get('state', state):
                obs.timer_add(lambda: update_source(source, obs_spec) or obs.remove_current_callback(), 2500)
            ahk_window['state'] = state

            center_item(sceneitem, ahk_spec)
            obs.obs_sceneitem_set_order_position(sceneitem, len(windows) - 1)


def get_ahk_spec(window_spec):
    title, class_, exe = window_spec.split(':')
    regex = re.compile(r'#[\dA-F]{2}')
    # replace things like #3A with :
    # new_title = regex.sub(lambda m: m.group().replace(m.group(), bytes.fromhex(m.group()[1:]).decode('utf-8')), title)
    new_title = title.replace('#3A', ':')
    is_re = title.startswith('/') and title.endswith('/')
    if is_re:
        new_title = new_title[1:-1]
    win_title = new_title + " ahk_exe " + exe
    return AhkWindowSpec(win_title, is_re, title, class_, exe)


def center_item(sceneitem, ahk_spec):
    if not ahk.f('WinGetWH', ahk_spec.win_title, ahk_spec.is_re):
        print("Matchless {ahk_spec}".format(**locals()))
        return

    vec2 = obs.vec2()
    vec2.x = video_info.base_width / 2 - ahk.get('w') / 2
    vec2.y = video_info.base_height / 2 - ahk.get('h') / 2
    obs.obs_sceneitem_set_pos(sceneitem, vec2)


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


def wipe_scene(scene):
    sceneitems = obs.obs_scene_enum_items(scene)
    for sceneitem in sceneitems:
        obs.obs_sceneitem_remove(sceneitem)
        source = obs.obs_sceneitem_get_source(sceneitem)
        obs.obs_source_remove(source)
    obs.sceneitem_list_release(sceneitems)


@log
def scenes_loaded():
    for scene_name, scene_windows in loaded['scenes'].items():
        scene = get_scene_by_name(scene_name)

        creating_scene = scene is None
        if creating_scene:
            scene = obs.obs_scene_create(scene_name)
        else:
            wipe_scene(scene)

        for window_name, window_spec in scene_windows.items():
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {'method': 2, 'priority': 1, 'window': window_spec}}
            data = obs.obs_data_create_from_json(json.dumps(source_info))
            source = obs.obs_load_source(data)
            obs.obs_data_release(data)

            sceneitem = obs.obs_scene_add(scene, source)
            obs.obs_source_release(source)
            obs.obs_sceneitem_set_locked(sceneitem, True)

            ahk_spec = get_ahk_spec(window_spec)
            center_item(sceneitem, ahk_spec)

            # change our loaded data from just the spec to a dict
            scene_windows[window_name] = {'spec': window_spec, 'sceneitem_id': obs.obs_sceneitem_get_id(sceneitem)}
            # global (not scene-specific) ahk window data
            loaded.setdefault('ahk_windows', {})[window_name] = {'spec': ahk_spec}

        if creating_scene:
            obs.obs_scene_release(scene)

    obs.timer_add(timer, 50)


@log
def script_load(settings):
    global ahk, video_info
    init()
    ahk = Script.from_file(r'script.ahk')

    video_info = obs.obs_video_info()
    obs.obs_get_video_info(video_info)
    obs.timer_add(wait_for_load, 1000)


# checking scenes with obs_frontend_get_scene_names() still wouldn't tell us if all the scene items were loaded
def wait_for_load():
    obs.remove_current_callback()
    scenes_loaded()


@log
def script_unload():
    obs.timer_remove(wait_for_load)
    if ahk is not None: # None if failed to load
        ahk.exit()


def script_description():
    return "Python & AutoHotkey powered OBS."


if __name__ == '__main__':
    init()
