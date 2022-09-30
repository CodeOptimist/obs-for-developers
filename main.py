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
from typing import NamedTuple, Iterator, ContextManager, ClassVar
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
class Window:
    MATCH_PRIORITY: ClassVar = {'type': 0, 'title': 1, 'exe': 2}

    # determines print() display order
    exists: bool
    priority: int
    title: str
    class_: str
    exe: str
    re_win_title: str
    source: Source = field(repr=False)
    sceneitem: SceneItem = field(repr=False)

    def __init__(self, window_spec: dict) -> None:
        self.exists = False
        self.priority = Window.MATCH_PRIORITY[window_spec.get('fallback', 'title')]

        # just used for win_title
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

    def obs_spec(self) -> str:
        ahk.call('WinGet', self.re_win_title)

        def s(name: str) -> str:
            return ahk.get(name).replace(':', '#3A')

        result = ":".join([s('title'), s('class'), s('exe')])
        return result


# globals are reset on reload
ahk: Script
loaded: dict
video_info: VideoInfo
windows: Dict[str, Dict[str, Window]] = defaultdict(defaultdict)
OBS_ALIGN_CENTER = 0


def update_active_win_sources() -> None:
    with get_source(obs.obs_frontend_get_current_scene()) as cur_scene_source:
        cur_scene_name: str = obs.obs_source_get_name(cur_scene_source)

    scene_windows = windows[cur_scene_name]
    for window_name, window in windows[cur_scene_name].items():
        was_closed = window.exists and not ahk.f('WinExist', window.re_win_title)
        if was_closed:
            # hide so OBS doesn't fallback to something undesirable (folder with same name as program, etc.)
            obs.obs_sceneitem_set_visible(window.sceneitem, False)  # :HideBadCapture
            window.exists = False

        if ahk.f('WinActive', window.re_win_title):
            with get_data(obs.obs_save_source(window.source)) as data:
                source_info = json.loads(obs.obs_data_get_json(data))

            obs_spec = window.obs_spec()
            # always true on first window activation since we left 'type' empty
            if source_info['settings']['window'] != obs_spec:
                # No documentation regarding 'settings' of obs_source_info; gleaned from '%AppData%\obs-studio\basic\scenes\Untitled.json'
                # https://obsproject.com/docs/reference-sources.html#source-definition-structure-obs-source-info
                print(f"Updating source to {obs_spec}")
                source_info['settings']['window'] = obs_spec
                source_info['settings']['priority'] = window.priority
                with get_data(obs.obs_data_create_from_json(json.dumps(source_info['settings']))) as new_data:
                    obs.obs_source_update(window.source, new_data)

            # layer the window captures
            obs.obs_sceneitem_set_order_position(window.sceneitem, len(scene_windows) - 1)
            obs.obs_sceneitem_set_visible(window.sceneitem, True)
            window.exists = True


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
        update_active_win_sources()
    except AhkExitException:
        obs.remove_current_callback()
    except Exception:
        obs.remove_current_callback()
        raise


@log
def scenes_loaded() -> None:
    # noinspection PyShadowingNames
    def get_scene_by_name(scene_name: str) -> Optional[Scene]:
        with get_source_list(obs.obs_frontend_get_scenes()) as sources:
            for source in sources:
                name: str = obs.obs_source_get_name(source)
                if name == scene_name:
                    return obs.obs_scene_from_source(source)
            return None

    for scene_name, scene_windows in loaded['scenes'].items():
        scene = get_scene_by_name(scene_name)
        group: SceneItem = obs.obs_scene_get_group(scene, "Windows")  # None if scene is None
        if group is None:
            continue

        # noinspection PyShadowingNames
        @log
        def wipe_group(group: SceneItem, group_scene: Scene) -> None:
            with get_sceneitem_list(obs.obs_scene_enum_items(group_scene)) as sceneitems:
                for sceneitem in sceneitems:
                    obs.obs_sceneitem_group_remove_item(group, sceneitem)
                    obs.obs_sceneitem_remove(sceneitem)
                    source: Source = obs.obs_sceneitem_get_source(sceneitem)
                    obs.obs_source_remove(source)  # notifies reference holders to release

        group_scene: Scene = obs.obs_sceneitem_group_get_scene(group)
        wipe_group(group, group_scene)

        obs.obs_sceneitem_set_visible(group, False)
        obs.obs_sceneitem_set_locked(group, True)

        center: Vec2 = obs.vec2()
        center.x = video_info.base_width / 2
        center.y = video_info.base_height / 2
        # no effect on a group: obs.obs_sceneitem_set_alignment(group, OBS_ALIGN_CENTER)
        obs.obs_sceneitem_set_pos(group, center)

        for idx, (window_name, window_spec) in enumerate(scene_windows.items()):
            if isinstance(window_spec, str):
                window_spec = {'window': window_spec}

            window = Window(window_spec)
            source_info = {'id': 'window_capture', 'name': window_name, 'settings': {
                # initialize window now for cosmetic text in OBS
                # our /some_regex/ syntax is just plaintext to OBS so could accidentally capture :HideBadCapture
                'window': f'{window.title}::{window.exe}',  # 'type' (between ':') is blank initially
                'method': 2,  # 'Windows 10 (1903 and up)'
                # poor name, more like fallback method when specific window disappears
                #  'title' will actually fallback to another window with identical title
                #  'type' with a blank value in 'window' ^ is the only way to really disable it
                'priority': Window.MATCH_PRIORITY['type'],
                'cursor': window_spec.get('cursor', True),
                'client_area': window_spec.get('client_area', False)
            }}
            with get_data(obs.obs_data_create_from_json(json.dumps(source_info))) as data:
                with get_source(obs.obs_load_source(data)) as source:
                    window.source = source
                    sceneitem: SceneItem = obs.obs_scene_add(group_scene, source)
                    window.sceneitem = sceneitem

            obs.obs_sceneitem_set_visible(sceneitem, False)  # :HideBadCapture
            obs.obs_sceneitem_set_locked(sceneitem, True)
            obs.obs_sceneitem_set_alignment(sceneitem, OBS_ALIGN_CENTER)
            windows[scene_name][window_name] = window

        obs.obs_sceneitem_set_visible(group, True)

    obs.timer_add(timer, 50)


def init() -> None:
    global loaded
    # don't use os.chdir() or it will break OBS
    data_path = Path(r'C:\Dropbox\Python\obs\captures.yaml')
    with data_path.open(encoding='utf-8') as f:
        loaded = yaml.safe_load(f)


# noinspection PyUnusedLocal
@log
def script_load(settings) -> None:
    global ahk, video_info
    init()
    ahk = Script.from_file(Path(r'C:\Dropbox\Python\obs\script.ahk'))

    video_info = obs.obs_video_info()  # annotated at global scope https://stackoverflow.com/questions/67527942/
    obs.obs_get_video_info(video_info)

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
    return "ahkUnwrapped powered OBS."


if __name__ == '__main__':
    init()
