# EnforceAudioDevice
A python based application that can be used to enforce the audio device on a per application basis as setting it in the windows settings will often not work as expected.

The app monitors the configured apps added in the `Config.json` starting and then forces the audio device.
The tool doesn't come with a full gui and only shows up in the system tray. 

## Requirements
The actual setting of audio devices is done via nirsoft SoundVolumeView utility. You can get it here [SoundVolumeView](https://www.nirsoft.net/utils/sound_volume_view.html).

This application is windows only!

The tool is mainly based on PyQt5 and WMI. All dependencies are packed into the executable, so you don't have to have python or any of the packages used installed.

## Usage

Right click the tray icon to reload the config, quit the app or to set the autostart option.

- Download the latest version from the [Releases](https://github.com/TheCoCe/EnforceAudioDevice/releases/latest)
- Move the `EnforceAudioDevice.exe` to your preferred location (this will generate files, so you might want to place it into a folder)
- Start the application
- Open the newly created config file (click the tray icon, choose Config → Open Config) and set the `SoundVolumeViewPath":` to your exe path, e.g. `C:\Programs\SoundVolumeView\SoundVolumeView.exe`
- Add applications to the config file and set your preferred audio device per app. The name should be the name of the exe file e.g.:
```json
"Apps": {
    "firefox": { "Device": "System" },
    "spotify": { "Device": "Music" },
    "overwatch": { "Device": "Game", "Delay": "5.0" }, // Some apps need delay as they don't always init audio right away
    "ffxiv": { "Device": "Game"}
  }
```

> :info: **How do I know the exe name?**</br> Open the task manager, find your applicationd and right click and choose `Properties` (You might need to click a subprocess). Go to the `General`. The exact name of the exe will be show at the top.

- Reload the application by right clicking the tray icon and choosing `Config` → `Reload Config`
- Enjoy the correct audio devices

You could also run the tool right from the `EnforceAudioDevice.py` if you have the required packages installed.

## Build
If you want to build the exe yourself you can use `pyinstaller` with the provided `spec` file or run the following command:
```bash
pyinstaller EnforceAudioDevice.py -F --noconsole -i EnforceAudioDevice.ico --add-data "EnforceAudioDevice.ico;." --add-data "EnforceAudioDeviceAlert.ico;." --hidden-import plyer.platforms.win.notification
```