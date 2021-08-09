# https://obsproject.com/docs/scripting.html
# https://obsproject.com/docs/reference-frontend-api.html
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
from dataclasses import dataclass, field
from typing import Dict, NewType, Optional, Iterable, Callable
from collections import defaultdict
from functools import partial

class VideoInfo(NamedTuple):
    # and many more
    base_width: int
    base_height: int


class Vec2(NamedTuple):
    x: float
    y: float


Source = NewType('Source', object)
Scene = NewType('Scene', object)
SceneItem = NewType('SceneItem', object)
Data = NewType('Data', object)


@dataclass
class Window:
    # determines print() display order
    exists: bool
    min_max: Optional[str]
    priority: int
    title: str
    class_: str
    exe: str
    re_win_title: str
    source: Source = field(repr=False)
    sceneitem: SceneItem = field(repr=False)

    def __init__(self, window_spec: dict) -> None:
        self.exists = False
        self.min_max = None
        self.priority = ({'title': 1, 'type': 0, 'exe': 2}[window_spec.get('fallback', 'title')])

        # just use for win_title
        def r(text: str) -> str:
            result = text.replace('#3A', ':')
            is_re = result.startswith('/') and result.endswith('/')
            result = result[1:-1] if is_re else re.escape(result)
            return result

        self.title, self.class_, self.exe = window_spec['window'].split(':')
        self.re_win_title: str = r(self.title)
        if self.class_:
            self.re_win_title += f" ahk_class {r(self.class_)}"
        if self.exe:
            self.re_win_title += f" ahk_exe {r(self.exe)}"

    def center(self) -> None:
        if not ahk.f('WinGetWH', self.re_win_title):
            print(f"Matchless {self}")
            return

        vec2: Vec2 = obs.vec2()
        vec2.x = video_info.base_width / 2 - ahk.get('w') / 2
        vec2.y = video_info.base_height / 2 - ahk.get('h') / 2
        obs.obs_sceneitem_set_pos(self.sceneitem, vec2)

    def obs_spec(self) -> str:
        ahk.call('WinGet', self.re_win_title)

        def s(name: str) -> str:
            return ahk.get(name).replace(':', '#3A')

        result = ":".join([s('title'), s('class'), s('exe')])
        return result


# globals are reset on reload
ahk: Script
loaded: dict
windows: Dict[str, Dict[str, Window]] = defaultdict(defaultdict)
video_info: VideoInfo


def update_active_win_sources() -> None:
    cur_scene_source: Source = obs.obs_frontend_get_current_scene()
    cur_scene_name: str = obs.obs_source_get_name(cur_scene_source)
    obs.obs_source_release(cur_scene_source)

    scene_windows = windows[cur_scene_name]
    for window_name, window in windows[cur_scene_name].items():
        was_closed = window.exists and not ahk.f('WinExist', window.re_win_title)
        if was_closed:
            # hide so OBS doesn't fallback to something undesirable (folder with same name as program, etc.)
            obs.obs_sceneitem_set_visible(window.sceneitem, False)
            window.exists = False

        def update_source(window, force=False) -> None:
            data: Data = obs.obs_save_source(window.source)
            source_info = json.loads(obs.obs_data_get_json(data))
            obs.obs_data_release(data)

            obs_spec = window.obs_spec()
            if force or source_info['settings']['window'] != obs_spec:
                print(f"Updating source to {obs_spec}")
                source_info['settings']['window'] = obs_spec
                source_info['settings']['priority'] = window.priority
                new_data: Data = obs.obs_data_create_from_json(json.dumps(source_info['settings']))
                obs.obs_source_update(window.source, new_data)
                obs.obs_data_release(new_data)

        if ahk.f('WinActive', window.re_win_title):
            update_source(window)

            # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
            min_max = ahk.get('min_max')
            if min_max != window.min_max:
                def do_update_source(window):
                    update_source(window, force=True)
                    obs.remove_current_callback()
                # obs.timer_add(partial(do_update_source, window), 2500)
            window.min_max = min_max

            window.center()
            obs.obs_sceneitem_set_order_position(window.sceneitem, len(scene_windows) - 1)
            obs.obs_sceneitem_set_visible(window.sceneitem, True)
            window.exists = True


def log(func: Callable) -> Callable:
    def wrapper(*args, **kwargs):
        before: int = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after: int = obs.bnum_allocs()
        print(f"{datetime.now().strftime('%I:%M:%S:%f')} {before} {func.__name__}() {after}")
        return result

    return wrapper


@log
def scenes_loaded() -> None:
    def get_scene_by_name(scene_name: str) -> Optional[Scene]:
        sources = ()
        try:
            sources: Iterable[Source] = obs.obs_frontend_get_scenes()
            for source in sources:
                name: str = obs.obs_source_get_name(source)
                if name == scene_name:
                    return obs.obs_scene_from_source(source)
            return None
        finally:
            obs.source_list_release(sources)

    global_scene = get_scene_by_name("Global")
    global_sceneitems: Iterable[SceneItem] = obs.obs_scene_enum_items(global_scene)
    global_count = sum(1 for _ in global_sceneitems)
    obs.sceneitem_list_release(global_sceneitems)
    print(f"Global count is: {global_count}")

    for scene_name, scene_windows in loaded['scenes'].items():
        scene = get_scene_by_name(scene_name)

        creating_scene = scene is None
        if creating_scene:
            # scene: Scene = obs.obs_scene_create(scene_name)
            OBS_SCENE_DUP_REFS = 0
            scene: Scene = obs.obs_scene_duplicate(global_scene, scene_name, OBS_SCENE_DUP_REFS)
        else:
            def wipe_scene(scene: Scene) -> None:
                sceneitems: Iterable[SceneItem] = obs.obs_scene_enum_items(scene)
                for sceneitem in sceneitems:
                    obs.obs_sceneitem_remove(sceneitem)
                    source: Source = obs.obs_sceneitem_get_source(sceneitem)
                    source_name: str = obs.obs_source_get_name(source)
                    if obs.obs_scene_find_source(global_scene, source_name) is None:
                        obs.obs_source_remove(source)
                obs.sceneitem_list_release(sceneitems)

            wipe_scene(scene)

        for idx, (window_name, window_spec) in enumerate(scene_windows.items()):
            if isinstance(window_spec, str):
                window_spec = {'window': window_spec}

            window = Window(window_spec)
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {
                # initialize window now for cosmetic text in OBS
                # our /some_regex/ syntax is just plaintext to OBS but could still capture something
                'window': f'{window.title}::{window.exe}',  # blank type to avoid fallback
                'method': 2,
                'priority': 0,  # type capture fallback
                'cursor': window_spec.get('cursor', True),
                'client_area': window_spec.get('client_area', False)
            }}
            data: Data = obs.obs_data_create_from_json(json.dumps(source_info))
            source: Source = obs.obs_load_source(data)
            obs.obs_data_release(data)
            window.source = source

            # obs_scene_create() creates a scene, but obs_scene_add() adds and returns a sceneitem
            sceneitem: SceneItem = obs.obs_scene_add(scene, source)
            obs.obs_source_release(source)
            window.sceneitem = sceneitem

            # hide by default as could be an unintentional capture
            obs.obs_sceneitem_set_visible(sceneitem, False)
            obs.obs_sceneitem_set_locked(sceneitem, True)
            window.center()
            windows[scene_name][window_name] = window

        if creating_scene:
            obs.obs_scene_release(scene)

    def timer() -> None:
        try:
            update_active_win_sources()
        except Exception:
            obs.remove_current_callback()
            raise

    obs.timer_add(timer, 50)


def init() -> None:
    global loaded
    # don't use os.chdir() or it will break OBS
    data_path = Path(r'C:\Dropbox\Python\obs\captures.yaml')
    with data_path.open(encoding='utf-8') as f:
        loaded = yaml.safe_load(f)


@log
def script_load(settings) -> None:
    global ahk, video_info
    init()
    ahk = Script.from_file(Path(r'C:\Dropbox\Python\obs\script.ahk'))

    video_info = obs.obs_video_info()
    obs.obs_get_video_info(video_info)
    obs.timer_add(wait_for_load, 1000)


# checking scenes with obs_frontend_get_scene_names() still wouldn't tell us if all the scene items were loaded
def wait_for_load() -> None:
    obs.remove_current_callback()
    scenes_loaded()


@log
def script_unload() -> None:
    obs.timer_remove(wait_for_load)
    if ahk is not None:  # None if failed to load
        ahk.exit()


def script_description() -> str:
    return "ahkUnwrapped powered OBS."


if __name__ == '__main__':
    init()
