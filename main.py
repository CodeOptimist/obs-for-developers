# todo deleting 'Windows' group = crash
# todo re-use existing Source for new sceneitems
# todo delete/re-use existing sceneitems for certain things e.g. RimWorld.exe (new window each restart)
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
from typing import NamedTuple, Iterator, ContextManager, ClassVar, List
# noinspection PyUnresolvedReferences
from win32api import OutputDebugString
from dataclasses import dataclass, field
from typing import Dict, NewType, Optional, Callable
from collections import defaultdict
from contextlib import contextmanager


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

    # determines print() display order
    name: str
    window: str
    method: str = field(default='wgc')
    fallback: str = field(default='title')
    cursor: bool = field(default=True)
    client_area: bool = field(default=False)

    ahk_wintitle: str = field(init=False)

    def __post_init__(self):
        title, class_, exe = self.window.split(':')

        def to_ahk(text: str) -> str:
            result = text.replace('#3A', ':')
            is_re = result.startswith('/') and result.endswith('/')
            result = result[1:-1] if is_re else re.escape(result)
            return result

        self.ahk_wintitle: str = to_ahk(title)
        if class_:
            self.ahk_wintitle += f" ahk_class {to_ahk(class_)}"
        if exe:
            self.ahk_wintitle += f" ahk_exe {to_ahk(exe)}"


@dataclass
class OsWindow:
    scene_pattern_idx: int = field(compare=False)
    exists: bool
    focused: bool
    title: str = field(compare=False)  # some titles can change often (e.g. tabs), doesn't make OBS un-capture
    class_: str
    exe: str

    pattern: LoadedCaptureInfo = field(init=False, compare=False)
    id: int = field(init=False, compare=False)
    obs_spec: str = field(init=False, compare=False)

    obs_source: Source = field(init=False, compare=False, repr=False)  # used to get/set current capture settings
    obs_sceneitem: SceneItem = field(init=False, compare=False, repr=False)  # used to position and hide/show sceneitem

    def __post_init__(self):
        self.scene_pattern_idx = int(self.scene_pattern_idx)
        self.id = self.exists
        self.exists = self.exists != '0x0'
        self.focused = self.focused != '0x0'

        def to_obs(text: str) -> str:
            return text.replace(':', '#3A')

        self.obs_spec = f"{to_obs(self.title)}:{to_obs(self.class_)}:{to_obs(self.exe)}"


# globals are reset on reload
ahk: Script
loaded: dict
video_info: VideoInfo
center: Vec2
scene_patterns: Dict[str, List[LoadedCaptureInfo]] = defaultdict(list)
scene_windows: Dict[str, Dict[int, OsWindow]] = defaultdict(dict)
scene_group: Dict[str, Scene] = {}


def update_window_sceneitems() -> None:
    with get_source(obs.obs_frontend_get_current_scene()) as cur_scene_source:
        cur_scene_name: str = obs.obs_source_get_name(cur_scene_source)

    group_scene = scene_group.get(cur_scene_name)
    if group_scene is None:
        return

    # even something like `ahk.f('WinActive', f"ahk_id {match.id}")` takes 0.03 seconds, much too long to block waiting for live result
    patterns = scene_patterns[cur_scene_name]
    windows_str = ahk.f('GetWindowsCached', '\n'.join(pattern.ahk_wintitle for pattern in patterns))
    windows = {}
    for window_str in windows_str.split('\n')[:-1]:
        window = OsWindow(*window_str.split('\r')[:-1])
        window.pattern = patterns[window.scene_pattern_idx]
        windows[window.id] = window

    for window in windows.values():
        try:
            existing_window = scene_windows[cur_scene_name][window.id]
            if window == existing_window:
                continue

            # will be the same since window ids match, we only fetched window status above
            window.obs_source = existing_window.obs_source
            window.obs_sceneitem = existing_window.obs_sceneitem

            # compares fields that aren't compare=False
        except KeyError:
            create_in_obs(group_scene, window)

        if not window.exists:
            # hide so OBS doesn't fallback to something undesirable (folder with same name as program, etc.)
            obs.obs_sceneitem_set_visible(window.obs_sceneitem, False)  # :HideBadCapture

        if window.focused:
            with get_data(obs.obs_save_source(window.obs_source)) as data:
                source_info = json.loads(obs.obs_data_get_json(data))
                # OutputDebugString(f"LOADED: {source_info['settings']['window']}")

            # always true on first window activation since we left 'type' empty
            # title doesn't change between e.g. tabs in PyCharm; doesn't break capture; flickering without when updating source
            if source_info['settings']['window'].split(':')[1:] != window.obs_spec.split(':')[1:]:
                # No documentation regarding 'settings' of obs_source_info; gleaned from '%AppData%\obs-studio\basic\scenes\Untitled.json'
                # https://obsproject.com/docs/reference-sources.html#source-definition-structure-obs-source-info
                print(f"Updating source to {window.obs_spec}")
                source_info['settings']['window'] = window.obs_spec
                source_info['settings']['priority'] = window.pattern.fallback  # won't ever change, just different than initial
                with get_data(obs.obs_data_create_from_json(json.dumps(source_info['settings']))) as new_data:
                    # OutputDebugString(f"SAVING: {source_info['settings']['window']}")
                    obs.obs_source_update(window.obs_source, new_data)

            # layer the window captures
            obs.obs_sceneitem_set_order_position(window.obs_sceneitem, len(scene_windows[cur_scene_name]) - 1)
            obs.obs_sceneitem_set_visible(window.obs_sceneitem, True)
        scene_windows[cur_scene_name][window.id] = window


def log(func: Callable) -> Callable:
    # noinspection PyMissingTypeHints
    def wrapper(*args, **kwargs):
        # https://obsproject.com/docs/reference-libobs-util-bmem.html#c.bnum_allocs
        before: int = obs.bnum_allocs()
        result = func(*args, **kwargs)
        after: int = obs.bnum_allocs()
        print(f"{datetime.now().strftime('%I:%M:%S:%f')} {before} {func.__name__}() {after}")
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
    global scene_group

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

        # noinspection PyShadowingNames
        @log
        def wipe_group(group_sceneitem: SceneItem, group_scene: Scene) -> None:
            with get_sceneitem_list(obs.obs_scene_enum_items(group_scene)) as sceneitems:
                for sceneitem in sceneitems:
                    obs.obs_sceneitem_group_remove_item(group_sceneitem, sceneitem)
                    obs.obs_sceneitem_remove(sceneitem)
                    source: Source = obs.obs_sceneitem_get_source(sceneitem)
                    obs.obs_source_remove(source)  # notifies reference holders to release

        group_scene: Scene = obs.obs_sceneitem_group_get_scene(group_sceneitem)
        wipe_group(group_sceneitem, group_scene)
        scene_group[scene_name] = group_scene

        obs.obs_sceneitem_set_locked(group_sceneitem, True)
        obs.obs_sceneitem_set_visible(group_sceneitem, False)

        # don't mess with bounding box! that can be customized through GUI to scale/move entire group
        # make sure position and size are correct
        obs.obs_sceneitem_set_pos(group_sceneitem, obs.vec2())  # x = 0, y = 0
        obs.obs_sceneitem_set_alignment(group_sceneitem, 0x1 | 0x4)  # OBS_ALIGN_LEFT | OBS_ALIGN_TOP
        scale: Vec2 = obs.vec2()
        obs.vec2_set(scale, 1, 1)
        obs.obs_sceneitem_set_scale(group_sceneitem, scale)  # sets 'size' which is critical

        for window in windows.values():
            create_in_obs(group_scene, window)

        obs.obs_sceneitem_set_visible(group_sceneitem, True)

    obs.timer_add(timer, 500)


def create_in_obs(group_scene: Scene, window: OsWindow):
    source_info = {'id': 'window_capture', 'name': f"{window.pattern.name} {window.id}", 'settings': {
        # initialize window now for cosmetic text in OBS
        # our /some_regex/ syntax is just plaintext to OBS so could accidentally capture :HideBadCapture
        'window': f'{window.title}::{window.exe}',  # blank type to avoid initial capture :WaitOnCapture
        # auto seems to prefer wgc client-area, then full window?, then falls back on bitblt
        #  https://github.com/obsproject/obs-studio/blob/a45cb71f6e5c6410a0d7f950a0a6511b2b930817/plugins/win-capture/window-capture.c#L131
        'method': LoadedCaptureInfo.OBS_CAPTURE_METHOD[window.pattern.method],
        # poor name, more like fallback method when specific window disappears
        #  'title' will actually fallback to another window with identical title
        'priority': LoadedCaptureInfo.OBS_CAPTURE_FALLBACK['type'],  # :WaitOnCapture
        'cursor': window.pattern.cursor,
        'client_area': window.pattern.client_area,
    }}

    with get_data(obs.obs_data_create_from_json(json.dumps(source_info))) as data:
        with get_source(obs.obs_load_source(data)) as source:
            window.obs_source = source
            sceneitem: SceneItem = obs.obs_scene_add(group_scene, source)

    obs.obs_sceneitem_set_visible(sceneitem, False)  # :HideBadCapture
    obs.obs_sceneitem_set_pos(sceneitem, center)
    obs.obs_sceneitem_set_alignment(sceneitem, 0)  # OBS_ALIGN_CENTER
    obs.obs_sceneitem_set_locked(sceneitem, True)

    window.obs_sceneitem = sceneitem

def init() -> None:
    global loaded, ahk, scene_patterns, scene_windows
    ahk = Script.from_file(Path(r'C:\Dropbox\Python\obs\script.ahk'))

    # don't use os.chdir() or it will break OBS
    data_path = Path(r'C:\Dropbox\Python\obs\captures.yaml')
    with data_path.open(encoding='utf-8') as f:
        loaded = yaml.safe_load(f)

    for scene_name, pattern_specs in loaded['scenes'].items():
        for idx, (pattern_name, pattern_spec) in enumerate(pattern_specs.items()):
            if isinstance(pattern_spec, str):
                pattern_spec = {'window': pattern_spec}
            pattern = LoadedCaptureInfo(pattern_name, **pattern_spec)
            scene_patterns[scene_name].append(pattern)

            ahk.set('_wintitles', f"{pattern.ahk_wintitle}")
            windows_str = ahk.f('GetWindows')
            for window_str in windows_str.split('\n')[:-1]:
                window = OsWindow(*window_str.split('\r')[:-1])
                window.pattern = pattern
                scene_windows[scene_name][window.id] = window
        pass


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
        #  > "nowadays [...] doesn’t actually mean that there is a stack buffer overrun.
        #  > [...] just means that the application decided to terminate itself with great haste."
        #  > Raymond Chen 2019-01-08
        time.sleep(0.01)  # 0.001 is too small, 0.005 seemed large enough


def script_description() -> str:
    return "Powered by ahkUnwrapped."


if __name__ == '__main__':
    init()
