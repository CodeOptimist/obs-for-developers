# https://obsproject.com/docs/scripting.html
# https://obsproject.com/docs/reference-core.html
# https://obsproject.com/docs/reference-scenes.html
# https://obsproject.com/docs/reference-sources.html
import json
import re

import yaml

if __name__ != '__main__':
    # noinspection PyUnresolvedReferences
    import obspython as obs
# noinspection PyUnresolvedReferences
from ahkunwrapped import Script, AhkException
from datetime import datetime
from pathlib import Path
from typing import NamedTuple
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString
from dataclasses import dataclass
from typing import Dict, NewType, Optional
from collections import defaultdict

class VideoInfo(NamedTuple):
    base_width: int
    base_height: int


SceneItem = NewType('SceneItem', object)


@dataclass
class SceneItemWindow:
    # determines display order
    state: Optional[str]
    is_re: bool
    title: str
    class_: str
    exe: str
    win_title: str

    def __init__(self, sceneitem: SceneItem, window_spec: str) -> None:
        self.state = None
        self.title, self.class_, self.exe = window_spec.split(':')
        self.sceneitem: SceneItem = sceneitem
        new_title = self.title.replace('#3A', ':')

        self.is_re = self.title.startswith('/') and self.title.endswith('/')
        if self.is_re:
            self.win_title: str = f'{new_title[1:-1]} ahk_exe {re.escape(self.exe)}'
        else:
            self.win_title: str = f'{new_title} ahk_exe {self.exe}'

    def center(self):
        if not ahk.f('WinGetWH', self.win_title, self.is_re):
            print(f"Matchless {self}")
            return

        vec2 = obs.vec2()
        vec2.x = video_info.base_width / 2 - ahk.get('w') / 2
        vec2.y = video_info.base_height / 2 - ahk.get('h') / 2
        obs.obs_sceneitem_set_pos(self.sceneitem, vec2)


# globals are reset on reload
ahk: Script
loaded: dict
window_sceneitems: Dict[str, Dict[str, SceneItemWindow]] = defaultdict(defaultdict)
video_info: VideoInfo


def init():
    global loaded
    # don't use os.chdir() or it will break OBS
    data_path = Path(r'C:\Dropbox\Python\obs\captures.yaml')
    with data_path.open(encoding='utf-8') as f:
        loaded = yaml.safe_load(f)


def log(func):
    def wrapper(*args, **kwargs):
        before = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after = obs.bnum_allocs()
        print(f"{datetime.now().strftime('%I:%M:%S:%f')} {before} {func.__name__}() {after}")
        return result

    return wrapper


def update_source(source, obs_spec, cond=None):
    data = obs.obs_save_source(source)
    source_info = json.loads(obs.obs_data_get_json(data))
    obs.obs_data_release(data)
    if cond is None or cond(source_info):
        print(f"Updating source to {obs_spec}")
        source_info['settings']['window'] = obs_spec
        new_data = obs.obs_data_create_from_json(json.dumps(source_info['settings']))
        obs.obs_source_update(source, new_data)
        obs.obs_data_release(new_data)


def timer():
    try:
        update_active_win_sources()
    except Exception:
        obs.remove_current_callback()
        raise


def update_active_win_sources():
    cur_scene_source = obs.obs_frontend_get_current_scene()
    # cur_scene = obs.obs_scene_from_source(cur_scene_source)
    cur_scene_name = obs.obs_source_get_name(cur_scene_source)
    obs.obs_source_release(cur_scene_source)
    # print(f"Current scene: {cur_scene_name}")

    windows = window_sceneitems.get(cur_scene_name, {})
    for window_name, window_sceneitem in windows.items():
        if ahk.f('WinActiveRegEx', window_sceneitem.win_title, window_sceneitem.is_re):
            source = obs.obs_sceneitem_get_source(window_sceneitem.sceneitem)

            ahk.call('ActiveWinGet')
            obs_spec = ":".join([(ahk.get('title').replace(':', '#3A')), ahk.get('class'), window_sceneitem.exe])
            update_source(source, obs_spec, cond=lambda source_info: source_info['settings']['window'] != obs_spec)

            # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
            state = ahk.get('state')
            if state != (window_sceneitem.state or state):
                obs.timer_add(lambda: update_source(source, obs_spec) or obs.remove_current_callback(), 2500)
            window_sceneitem.state = state

            window_sceneitem.center()
            obs.obs_sceneitem_set_order_position(window_sceneitem.sceneitem, len(windows) - 1)


def get_scene_by_name(scene_name):
    sources = None
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
def scenes_loaded():
    global_scene = get_scene_by_name("Global")
    global_sceneitems = obs.obs_scene_enum_items(global_scene)
    global_count = sum(1 for _ in global_sceneitems)
    obs.sceneitem_list_release(global_sceneitems)
    print(f"Global count is: {global_count}")

    for scene_name, scene_windows in loaded['scenes'].items():
        scene = get_scene_by_name(scene_name)

        creating_scene = scene is None
        if creating_scene:
            # scene = obs.obs_scene_create(scene_name)
            OBS_SCENE_DUP_REFS = 0
            scene = obs.obs_scene_duplicate(global_scene, scene_name, OBS_SCENE_DUP_REFS)
        else:
            def wipe_scene(scene):
                sceneitems = obs.obs_scene_enum_items(scene)
                for sceneitem in sceneitems:
                    obs.obs_sceneitem_remove(sceneitem)
                    source = obs.obs_sceneitem_get_source(sceneitem)
                    source_name = obs.obs_source_get_name(source)
                    if obs.obs_scene_find_source(global_scene, source_name) is None:
                        obs.obs_source_remove(source)
                obs.sceneitem_list_release(sceneitems)

            wipe_scene(scene)

        for idx, (window_name, window_spec) in enumerate(scene_windows.items()):
            # we set 'window' here for the cosmetic name within OBS; OBS doesn't actually support regex, that's our own addition
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {'method': 2, 'priority': 1, 'window': window_spec}}
            data = obs.obs_data_create_from_json(json.dumps(source_info))
            source = obs.obs_load_source(data)
            obs.obs_data_release(data)

            # obs_scene_create() creates a scene, but obs_scene_add() adds and returns a sceneitem
            sceneitem = obs.obs_scene_add(scene, source)
            obs.obs_source_release(source)
            obs.obs_sceneitem_set_locked(sceneitem, True)

            window_sceneitem = SceneItemWindow(sceneitem, window_spec)
            window_sceneitem.center()
            window_sceneitems[scene_name][window_name] = window_sceneitem

        if creating_scene:
            obs.obs_scene_release(scene)

    obs.timer_add(timer, 50)


@log
def script_load(settings):
    global ahk, video_info
    init()
    ahk = Script.from_file(Path(r'C:\Dropbox\Python\obs\script.ahk'))

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
    if ahk is not None:  # None if failed to load
        ahk.exit()


def script_description():
    return "Python & AutoHotkey powered OBS."


if __name__ == '__main__':
    init()
