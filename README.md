# Office-Auto-Save
This is a simple backgroud app for Windows that checks for edits made to Office documents in realtime and saves the document for you. 
I made this because I found it very annoying that Microsoft only allowed AutoSave on files stored on OneDrive. 
The saving doesn't interfere with the mouse or keyboard so it leaves you free to continue editing.

## Installation
Simply download the installer from the releases page here on github.

## Usage
After installation the app will start automatically when the computer boots up. 
However, the application can be started manually if desired.

### System Tray Icon
An icon will show in the hidden icons tray. This can be clicked or right clicked to close the process.

## Building the Source
To build simply run the .bat file from a terminal to get the .exe output.
For example: `C:\repos\Office-Auto-Save\> ./build.bat`
<br>
You can also, of course, just run it since it's python.
<br>
## Dependencies
Before you run/build it make sure you have installed these dependencies from pip:
<br>

```
pip install pywin32
pip install pynput
pip install pillow
pip install pystray
pip install win10toast
```

