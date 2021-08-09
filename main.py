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
from dataclasses import dataclass
from typing import Dict, NewType, Optional, Iterable, Callable
from collections import defaultdict

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
class SceneItemWindow:
    # determines print() display order
    state: Optional[str]
    title: str
    class_: str
    exe: str
    win_title: str

    def __init__(self, sceneitem: SceneItem, window_spec: dict) -> None:
        self.state = None
        self.sceneitem: SceneItem = sceneitem

        def r(text: str) -> str:
            result = text.replace('#3A', ':')
            is_re = result.startswith('/') and result.endswith('/')
            result = result[1:-1] if is_re else re.escape(result)
            return result

        self.title, self.class_, self.exe = window_spec['window'].split(':')
        self.win_title: str = r(self.title)
        if self.class_:
            self.win_title += f" ahk_class {r(self.class_)}"
        if self.exe:
            self.win_title += f" ahk_exe {r(self.exe)}"

    def center(self) -> None:
        if not ahk.f('WinGetWH', self.win_title):
            print(f"Matchless {self}")
            return

        vec2: Vec2 = obs.vec2()
        vec2.x = video_info.base_width / 2 - ahk.get('w') / 2
        vec2.y = video_info.base_height / 2 - ahk.get('h') / 2
        obs.obs_sceneitem_set_pos(self.sceneitem, vec2)


# globals are reset on reload
ahk: Script
loaded: dict
window_sceneitems: Dict[str, Dict[str, SceneItemWindow]] = defaultdict(defaultdict)
video_info: VideoInfo


def update_active_win_sources() -> None:
    cur_scene_source: Source = obs.obs_frontend_get_current_scene()
    cur_scene_name: str = obs.obs_source_get_name(cur_scene_source)
    obs.obs_source_release(cur_scene_source)

    windows = window_sceneitems[cur_scene_name]
    for window_name, window_sceneitem in window_sceneitems[cur_scene_name].items():
        if ahk.f('WinActiveRegEx', window_sceneitem.win_title):
            def update_source(source: Source, obs_spec: str, cond: Callable = None) -> None:
                data: Data = obs.obs_save_source(source)
                source_info = json.loads(obs.obs_data_get_json(data))
                obs.obs_data_release(data)
                if cond is None or cond(source_info):
                    print(f"Updating source to {obs_spec}")
                    source_info['settings']['window'] = obs_spec
                    new_data: Data = obs.obs_data_create_from_json(json.dumps(source_info['settings']))
                    obs.obs_source_update(source, new_data)
                    obs.obs_data_release(new_data)

            def e(text: str) -> str:
                return text.replace(':', '#3A')

            ahk.call('ActiveWinGet')
            obs_spec = ":".join([(e(ahk.get('title'))), e(ahk.get('class')), e(ahk.get('exe'))])
            source: Source = obs.obs_sceneitem_get_source(window_sceneitem.sceneitem)
            update_source(source, obs_spec, cond=lambda source_info: source_info['settings']['window'] != obs_spec)

            # todo AHK can't detect minimized & restore, crap https://autohotkey.com/board/topic/94409-detect-minimized-windows/
            state = ahk.get('state')
            if window_sceneitem.state is None or state != window_sceneitem.state:
                obs.timer_add(lambda: update_source(source, obs_spec) or obs.remove_current_callback(), 2500)
            window_sceneitem.state = state

            window_sceneitem.center()
            obs.obs_sceneitem_set_order_position(window_sceneitem.sceneitem, len(windows) - 1)


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

            # we set 'window' here for the cosmetic name within OBS; OBS doesn't actually support regex, that's our own addition
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {
                'window': window_spec['window'],
                'method': 2,
                'priority': ({'title': 1, 'type': 2, 'exe': 3}[window_spec.get('fallback', 'title')]),
                'cursor': window_spec.get('cursor', True),
                'client_area': window_spec.get('client_area', False)
            }}
            data: Data = obs.obs_data_create_from_json(json.dumps(source_info))
            source: Source = obs.obs_load_source(data)
            obs.obs_data_release(data)

            # obs_scene_create() creates a scene, but obs_scene_add() adds and returns a sceneitem
            sceneitem: SceneItem = obs.obs_scene_add(scene, source)
            obs.obs_source_release(source)
            obs.obs_sceneitem_set_locked(sceneitem, True)

            window_sceneitem = SceneItemWindow(sceneitem, window_spec)
            window_sceneitem.center()
            window_sceneitems[scene_name][window_name] = window_sceneitem

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
