# OBS ‘Window Capture’ for Developers
## Challenge
Unlike streaming video games, showcasing software development entails interaction with many windows, often on multiple monitors.

This inconvenient proliferation of application windows makes anything but 'Display Capture' impractical. This is an unfortunate burden to those who would otherwise be willing to stream from their (often highly-configured) personal or work machines.

To further detriment, arranging windows to fit a solitary display counteracts the productivity boon of extra screen real-estate.

## Solution
Leverage the [ahkUnwrapped](https://pypi.org/project/ahkunwrapped/) Python package and OBS [Scripting API](https://obsproject.com/docs/scripting.html) to asynchronously monitor application windows from an allowlist, populating and layering _Window Captures_ within OBS as necessary to reflect focus, child windows, etc.

## Demo
See recent _Software and Game Development_ broadcasts at https://www.twitch.tv/CodeOptimist.

## Features
* Explicit allowlist for intentional window captures.
* Regular expression matching of window titles.
* On-demand capture of child windows.
* Capture windows from any monitor.
* Automated layering by recently focused.

## Get started
### Automatic configuration
I value minimal setup, but it's not here yet.
### Manual configuration
It's recommended that you aren't broadcasting, or if so, edit the included  _captures.yaml_ before loading the script.
1. Install [Python 3.6](https://www.python.org/downloads/release/python-360/) to a unique path, e.g. `%AppData%\..\Local\Programs\Python\Python36-obs`  
(_Python 3.6_ is the latest OBS supports, and it won't use _virtualenv_.)  
You probably **don't** want to add this to your system _PATH_.
2. Clone (or [download](/../../archive/refs/heads/master.zip)) this repository to a preferred location, e.g. `C:\path\to\repo`
3. Install dependencies via commandline:
    ```
    %AppData%\..\Local\Programs\Python\Python36-obs\python -m pip install -r C:\path\to\repo\requirements.txt
    ```
4. Within OBS, create an empty group named `Windows` on the desired scene(s). The script will manage its contents.
6. Within OBS, select `Tools->Scripts, Python Settings` and browse to your Python install path.
7. Within OBS, select `Tools->Scripts, ➕ (Add Scripts)` and browse to `capture-windows.py`
8. Done. Matching application windows from `captures.yaml` should automatically be captured within OBS.
9. Make changes to _captures.yaml_ and `🔃 (Reload)` _capture-windows.py_ (make sure it's selected) anytime.

## Troubleshooting
`POLL_INTERVAL_MS` at the top of `capture-windows.py` will be adjustable from _Loaded Scripts_ in the future.
If you encounter:
> Encoding overloaded! Consider turning down video settings or using a faster encoding preset.
 
try increasing the value to `500` (half a second). Default is `250` (and can probably be set to `100` on many machines).