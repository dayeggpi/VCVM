# VCVM
VolumeControl for Voicemeeter

This software is not affiliated and is an independent software using VoicemeeterRemoteAPI.
Support Voicemeeter's amazing tool at https://vb-audio.com/index.htm

## Usage
Launch with `python VCVM.py`

You can also compile it to an exe by having the *.ico and *.spec file in same folder as VCVM.py and launching the file "build-exe.bat". The output exe will be in the "dist" folder.
Then you simply execute the exe file to launch it.

Once launched, the app will run in systray. Then you can change Windows volume via keyboard, laptop buttons, mousewheel or going to Windows Sound Volume in systray, it will impact Voicemeeter volume as well.

![Image](https://i.imgur.com/xjDvio1.gif)


## Systray
Right click on the systray icon to chose if you want to app to start with Windows.
Control from there if you want logging, and if you also want verbose logging (a bit more details).
You can also reload the app (if you adjusted the config.ini).

## Config
A config.ini file will be generated to adjust some settings.

```
[Voicemeeter]
dll_path = c:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll

[Logging]
enabled = false
verbose = false
log_file = VCVM.log

[Settings]
curve_power = 0.55
sync_interval = 0.3
change_timeout = 4
gain_threshold = 3.0
volume_threshold = 1
bus = 0,1,2,3,4

[Startup]
delay_seconds = 5
max_retry_attempts = 5
retry_interval = 2
```
adjust "dll_path" as per your install.

adjust "bus" depending on if you want to control only A1 or A1 to A5 (bus 0 is A1).

adjust "log_file" to edit the name of the log file.

adjust delay_seconds to give few seconds delay once the app will start (note that Windows Tasks will start after 5seconds, can be adjusted in code).

adjust the rest of the settings to play with the curve of volume control.
