# todo re-use existing Source for Window sceneitems in completely different scenes?
# https://obsproject.com/docs/scripting.html
# https://obsproject.com/docs/reference-frontend-api.html
# https://obsproject.com/docs/reference-core.html
# https://obsproject.com/docs/reference-scenes.html
# https://obsproject.com/docs/reference-sources.html
# To retain sanity, EVERY "= obs." should be preceded by a type annotation.

import json
import re
# noinspection PyUnresolvedReferences
import threading
import time
from contextlib import suppress

import yaml

if __name__ != '__main__':
    # noinspection PyUnresolvedReferences
    import obspython as obs
from ahkunwrapped import Script, AhkExitException
from datetime import datetime
from pathlib import Path
from typing import NamedTuple, Iterator, ContextManager, ClassVar, List, Tuple, Set
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString
from dataclasses import dataclass, field
from typing import Dict, NewType, Optional, Callable
from collections import defaultdict
from contextlib import contextmanager

# todo make this an OBS script property
# https://obsproject.com/wiki/Getting-Started-With-OBS-Scripting#global-script-functions-for-editable-properties
# https://obsproject.com/docs/reference-properties.html#c.obs_properties_create
POLL_INTERVAL_MS = 250  # Script execution is "blocking", so polling too quickly can overload the video encoder.

class VideoInfo(NamedTuple):
    # and many more
    base_width: int
    base_height: int


class Vec2(NamedTuple):
    x: float
    y: float


Data = NewType('Data', object)
Source = NewType('Source', object)
Scene = NewType('Scene', object)
SceneItem = NewType('SceneItem', object)


# https://stackoverflow.com/questions/49733699/python-type-hints-and-context-managers
@contextmanager
def get_data(data) -> ContextManager[Data]:
    try: yield data
    finally: obs.obs_data_release(data)

@contextmanager
def get_source(source) -> ContextManager[Source]:
    try: yield source
    finally: obs.obs_source_release(source)

@contextmanager
def get_source_list(source_list) -> ContextManager[Iterator[Source]]:
    try: yield source_list
    finally: obs.source_list_release(source_list)

@contextmanager
def get_sceneitem_list(sceneitem_list) -> ContextManager[Iterator[SceneItem]]:
    try: yield sceneitem_list
    finally: obs.sceneitem_list_release(sceneitem_list)


@dataclass
class LoadedCaptureInfo:
    OBS_CAPTURE_FALLBACK: ClassVar = {'type': 0, 'title': 1, 'exe': 2}
    OBS_CAPTURE_METHOD: ClassVar = {'auto': 0, 'bitblt': 1, 'wgc': 2}

    name: str
    # temporary used to set `ahk_wintitle`
    window: str  # actual pattern, like obs_spec (`title:class:exe`) but each is optional and `/regex/` support

    method: str = field(default='wgc')  # source_info[settings]['method']
    fallback: str = field(default='title')  # source_info[settings]['priority']
    cursor: bool = field(default=True)  # source_info[settings]['cursor']
    client_area: bool = field(default=False)  # source_info[settings]['client_area']

    ahk_wintitle: str = field(init=False)

    def __post_init__(self) -> None:
        title, class_, exe = self.window.split(':')
        assert title or class_ or exe, "Empty capture pattern in file"

        if self.fallback == 'title' and not title:
            self.fallback = 'exe'

        def to_ahk(text: str) -> str:
            result = text.replace('#3A', ':')
            is_re = result.startswith('/') and result.endswith('/')
            result = result[1:-1] if is_re else re.escape(result)
            return result

        self.ahk_wintitle = "^"
        if title:
            self.ahk_wintitle += to_ahk(title)
        if class_:
            self.ahk_wintitle += f" ahk_class {to_ahk(class_)}"
        if exe:
            self.ahk_wintitle += f" ahk_exe {to_ahk(exe)}"
        self.ahk_wintitle += "$"


@dataclass
class OsWindow:
    # we disable comparisons for use with set() to find closed & opened windows
    scene_name: str = field(compare=False)

    # from ahk
    scene_pattern_idx: None = field(compare=False, repr=False)  # temporary to set `pattern`
    id: str
    focused: bool = field(compare=False)
    title: str = field(compare=False)
    class_: str = field(compare=False)
    exe: str = field(compare=False)

    pattern: LoadedCaptureInfo = field(init=False, compare=False)
    obs_spec: str = field(init=False, compare=False)

    def __hash__(self) -> int:
        return hash(self.id)

    def __str__(self) -> str:
        return f"{self.pattern.name} {self.id}"

    def __post_init__(self) -> None:
        self.focused = self.focused != '0x0'
        self.pattern = scene_patterns[self.scene_name][int(self.scene_pattern_idx)]
        self.scene_pattern_idx = None  # just to avoid confusion

        def to_obs(text: str) -> str:
            return text.replace(':', '#3A')

        # class is required here, title and exe are not (no matter OBS capture 'priority' or 'method')
        # (of course we want to be as specific as possible to get the right window)
        self.obs_spec = f"{to_obs(self.title)}:{to_obs(self.class_)}:{to_obs(self.exe)}"

    @staticmethod
    def get_windows(scene_name: str, from_cache: bool):
        ahk.set('wintitles', '\n'.join(pattern.ahk_wintitle for pattern in scene_patterns[scene_name]))
        result_str = ahk.f('GetWindowsCached' if from_cache else 'GetWindows')
        result = set(OsWindow(scene_name, *window_str.split('\r')[:-1]) for window_str in result_str.split('\n')[:-1])
        return result


# globals are reset on reload
ahk: Script
loaded: dict
video_info: VideoInfo
center: Vec2
scene_patterns: Dict[str, List[LoadedCaptureInfo]] = defaultdict(list)  # user data
scene_windows: Dict[str, Set[OsWindow]] = defaultdict(set)  # os data

# obs data
name_scenes: Dict[str, Scene] = defaultdict()
scene_window_sceneitems: Dict[str, Dict[str, SceneItem]] = defaultdict(dict)


def update_window_sceneitems() -> None:
    with get_source(obs.obs_frontend_get_current_scene()) as cur_scene_source:
        cur_scene_name: str = obs.obs_source_get_name(cur_scene_source)  # None if given None

    group_sceneitem: SceneItem = obs.obs_scene_get_group(name_scenes.get(cur_scene_name), "Windows")  # None if scene is None
    if group_sceneitem is None:
        return
    group_scene: Optional[Scene] = None

    last_windows = scene_windows[cur_scene_name]
    # ahkUnwrapped itself is <1ms but `ahk.f('WinActive', f"ahk_id {match.id}")` takes 30ms; much too slow (blocks encoding)
    cur_windows = OsWindow.get_windows(cur_scene_name, from_cache=True)
    scene_windows[cur_scene_name] = cur_windows

    for closed_window in last_windows - cur_windows:
        sceneitem = scene_window_sceneitems[cur_scene_name].pop(closed_window.id)
        print_debug(f"Removing '{closed_window}' ('{closed_window.obs_spec}')")
        obs.obs_sceneitem_group_remove_item(group_sceneitem, sceneitem)
        obs.obs_sceneitem_remove(sceneitem)
        source: Source = obs.obs_sceneitem_get_source(sceneitem)
        obs.obs_source_remove(source)  # notifies reference holders to release

    for opened_window in cur_windows - last_windows:
        sceneitem = scene_window_sceneitems[cur_scene_name].get(opened_window.id)
        if sceneitem is None:
            print_debug(f"Creating '{opened_window}'")
            if group_scene is None:
                group_scene: Scene = obs.obs_sceneitem_group_get_scene(group_sceneitem)
            create_in_obs(cur_scene_name, group_sceneitem, group_scene, opened_window)

    for window in cur_windows:  # includes opened
        if window.focused:
            sceneitem = scene_window_sceneitems[cur_scene_name][window.id]
            source: Source = obs.obs_sceneitem_get_source(sceneitem)

            with get_data(obs.obs_save_source(source)) as data:
                source_info = json.loads(obs.obs_data_get_json(data))
                # OutputDebugString(f"LOADED: {source_info['settings']['window']}")

            # always true on first window activation since we left 'type' empty
            # title change won't affect the capture; flickers without this
            if source_info['settings']['window'].split(':')[1:] != window.obs_spec.split(':')[1:]:
                # No documentation regarding 'settings' of obs_source_info; gleaned from '%AppData%\obs-studio\basic\scenes\Untitled.json'
                # https://obsproject.com/docs/reference-sources.html#source-definition-structure-obs-source-info
                print_debug(f"Updating '{source_info['name']}' to '{window.obs_spec}'")
                source_info['settings']['window'] = window.obs_spec
                source_info['settings']['priority'] = window.pattern.fallback  # won't ever change, just different than initial

                with get_data(obs.obs_data_create_from_json(json.dumps(source_info['settings']))) as new_data:
                    # OutputDebugString(f"SAVING: {source_info['settings']['window']}")
                    obs.obs_source_update(source, new_data)

            # layer the window captures
            obs.obs_sceneitem_set_order_position(sceneitem, len(scene_window_sceneitems[cur_scene_name]) - 1)
            obs.obs_sceneitem_set_visible(sceneitem, True)


def print_debug(str_: str) -> None:
    print(str_)
    OutputDebugString(str_)


def log(func: Callable) -> Callable:
    # noinspection PyMissingTypeHints
    def wrapper(*args, **kwargs):
        # https://obsproject.com/docs/reference-libobs-util-bmem.html#c.bnum_allocs
        before: int = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after: int = obs.bnum_allocs()
        print_debug(f"{datetime.now().strftime('%I:%M:%S:%f')} {before} {func.__name__}() {after}")
        return result

    return wrapper


# this *CAN* still execute a finite amount after script_unload()
def timer() -> None:
    # OutputDebugString(f"Script tick. Thread: {threading.get_ident()}")
    try:
        update_window_sceneitems()
    except AhkExitException:
        obs.remove_current_callback()
    except Exception:
        obs.remove_current_callback()
        raise


@log
def scenes_loaded() -> None:
    global name_scenes

    # noinspection PyShadowingNames
    def get_scene_by_name(scene_name: str) -> Optional[Scene]:
        with get_source_list(obs.obs_frontend_get_scenes()) as sources:
            for source in sources:
                name: str = obs.obs_source_get_name(source)
                if name == scene_name:
                    return obs.obs_scene_from_source(source)
            return None

    for scene_name, windows in scene_windows.items():
        scene = get_scene_by_name(scene_name)
        group_sceneitem: SceneItem = obs.obs_scene_get_group(scene, "Windows")  # None if scene is None
        if group_sceneitem is None:
            continue
        name_scenes[scene_name] = scene
        group_scene: Scene = obs.obs_sceneitem_group_get_scene(group_sceneitem)

        # noinspection PyShadowingNames
        @log
        def wipe_group(group_sceneitem: SceneItem, group_scene: Scene) -> None:
            with get_sceneitem_list(obs.obs_scene_enum_items(group_scene)) as sceneitems:
                for sceneitem in sceneitems:
                    obs.obs_sceneitem_group_remove_item(group_sceneitem, sceneitem)
                    obs.obs_sceneitem_remove(sceneitem)
                    source: Source = obs.obs_sceneitem_get_source(sceneitem)
                    obs.obs_source_remove(source)  # notifies reference holders to release

        wipe_group(group_sceneitem, group_scene)

        obs.obs_sceneitem_set_visible(group_sceneitem, False)
        obs.obs_sceneitem_set_locked(group_sceneitem, True)

        # don't mess with bounding box! that can be customized through GUI to scale/move entire group
        obs.obs_sceneitem_set_pos(group_sceneitem, obs.vec2())  # x = 0, y = 0
        obs.obs_sceneitem_set_alignment(group_sceneitem, 0x1 | 0x4)  # OBS_ALIGN_LEFT | OBS_ALIGN_TOP
        # you can't set the size, only scale
        scale: Vec2 = obs.vec2()
        obs.vec2_set(scale, 1, 1)
        obs.obs_sceneitem_set_scale(group_sceneitem, scale)  # sets 'size' which is critical

        for window in windows:
            create_in_obs(scene_name, group_sceneitem, group_scene, window)

        obs.obs_sceneitem_set_visible(group_sceneitem, True)

    obs.timer_add(timer, POLL_INTERVAL_MS)


def create_in_obs(scene_name: str, group_sceneitem: SceneItem, group_scene: Scene, window: OsWindow):
    source_info = {'id': 'window_capture', 'name': str(window), 'settings': {
        # title and exe for initial cosmetic name in OBS
        'window': f'{window.title}::{window.exe}',  # blank type so as not to capture yet :AvoidBadCapture
        # auto seems to prefer wgc client-area, then full window?, then falls back on bitblt
        #  https://github.com/obsproject/obs-studio/blob/a45cb71f6e5c6410a0d7f950a0a6511b2b930817/plugins/win-capture/window-capture.c#L131
        'method': LoadedCaptureInfo.OBS_CAPTURE_METHOD[window.pattern.method],
        # poor name, more like fallback method when specific window disappears
        #  'title' will actually fallback to another window with identical title
        'priority': LoadedCaptureInfo.OBS_CAPTURE_FALLBACK['type'],  # :AvoidBadCapture
        'cursor': window.pattern.cursor,
        'client_area': window.pattern.client_area,
    }}

    with get_data(obs.obs_data_create_from_json(json.dumps(source_info))) as data:
        with get_source(obs.obs_load_source(data)) as source:
            sceneitem: SceneItem = obs.obs_scene_add(group_scene, source)
            scene_window_sceneitems[scene_name][window.id] = sceneitem

    obs.obs_sceneitem_set_visible(sceneitem, False)
    obs.obs_sceneitem_set_locked(sceneitem, True)

    group_pos: Vec2 = obs.vec2()
    obs.obs_sceneitem_get_pos(group_sceneitem, group_pos)

    new_pos: Vec2 = obs.vec2()
    # if you look at a small window's sceneitem position (Ctrl+E, or Transform -> Edit Transform...)
    # you'll see the position is not absolute coordinates: it will give different values if you hide larger sceneitems
    # seems to depend on the group's size, and you can't set size, only scale
    # (group's size will be the same as the largest visible sceneitem)
    # thus we have to correct by these shifting group position coordinates to get back to true center
    # problem shows up when a opening a non-full-size window with only other non-full-size sceneitems visible
    obs.vec2_set(new_pos, center.x - group_pos.x, center.y - group_pos.y)

    obs.obs_sceneitem_set_alignment(sceneitem, 0)  # OBS_ALIGN_CENTER
    obs.obs_sceneitem_set_pos(sceneitem, new_pos)

def init() -> None:
    global loaded, ahk, scene_patterns, scene_windows
    ahk = Script.from_file(Path(r'C:\Dropbox\Python\obs\script.ahk'))

    # don't use os.chdir() or it will break OBS
    data_path = Path(r'C:\Dropbox\Python\obs\captures.yaml')
    with data_path.open(encoding='utf-8') as f:
        loaded = yaml.safe_load(f)

    # it's nice to do this here since we can debug it without OBS
    for scene_name, pattern_specs in loaded['scenes'].items():
        for pattern_name, pattern_spec in pattern_specs.items():
            if isinstance(pattern_spec, str):
                pattern_spec = {'window': pattern_spec}
            pattern = LoadedCaptureInfo(pattern_name, **pattern_spec)
            scene_patterns[scene_name].append(pattern)

        scene_windows[scene_name] = OsWindow.get_windows(scene_name, from_cache=False)


# noinspection PyUnusedLocal
@log
def script_load(settings) -> None:
    global video_info, center
    init()

    video_info = obs.obs_video_info()  # annotated at global scope https://stackoverflow.com/questions/67527942/
    obs.obs_get_video_info(video_info)
    center = obs.vec2()
    center.x = video_info.base_width / 2
    center.y = video_info.base_height / 2

    obs.timer_add(wait_for_load, 1000)
    # OutputDebugString(f"Script load. Thread: {threading.get_ident()}")


# I'm not aware of a better method than just waiting.
# Checking scenes with obs_frontend_get_scene_names() isn't enough to know if all the sceneitems are loaded.
def wait_for_load() -> None:
    obs.remove_current_callback()
    scenes_loaded()


@log
def script_unload() -> None:
    # OutputDebugString(f"Script unload. Thread: {threading.get_ident()}")
    if ahk is not None:  # None if failed to load
        with suppress(AhkExitException):
            ahk.exit()
        # avoids crash with 'Reload Scripts': 0xc0000409 (EXCEPTION_STACK_BUFFER_OVERRUN)
        #  possibly an OBS bug and not ahkUnwrapped... knock on wood
        #  https://devblogs.microsoft.com/oldnewthing/20190108-00/?p=100655
        #  > "nowadays [...] doesn't actually mean that there is a stack buffer overrun.
        #  > [...] just means that the application decided to terminate itself with great haste."
        #  > Raymond Chen 2019-01-08
        time.sleep(0.01)  # 0.001 is too small, 0.005 seemed large enough


def script_description() -> str:
    return "Powered by ahkUnwrapped."


if __name__ == '__main__':
    init()
