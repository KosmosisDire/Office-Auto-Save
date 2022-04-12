import os
import sys
from win32gui import GetWindowText, GetForegroundWindow, PostMessage, GetWindowRect, EnumChildWindows, GetClassName, GetParent
from win32api import MAKELONG
from win32con import WM_LBUTTONDOWN, WM_LBUTTONUP, MK_LBUTTON, WM_RBUTTONDOWN, WM_RBUTTONUP
from pynput.keyboard import Key, Controller
from pynput import keyboard, mouse
import threading
import atexit
import time
from PIL import ImageFile, Image
from pystray import Icon as trayicon, Menu as menu, MenuItem as item
from win10toast import ToastNotifier

#global side
checking:bool = True
saving:bool = False

def exit_handler():
    global checking, saving, word_checker_thread, saving_thread
    checking = False
    saving = False

    if(word_checker_thread.is_alive()): 
        word_checker_thread.join()
    if(saving_thread.is_alive()):
        saving_thread.join()
    print("Exited");
    
atexit.register(exit_handler)

window_handle = GetForegroundWindow()
def office_open():
    global window_handle
    fore = GetForegroundWindow()
    window_text = GetWindowText(fore)
    is_open = ".docx - Word" in window_text or ".xlsx - Excel" in window_text or ".pptx - PowerPoint" in window_text
    if is_open:
        window_handle = fore
    return is_open

def current_office_window_still_open():
    global window_handle
    return window_handle == GetForegroundWindow()

# wait side
def office_checker():
    global saving, checking
    while(checking):
        if(office_open()):
            checking = False
            saving = True
            print("Office Opened")
            saving_thread.run()
        time.sleep(5)

word_checker_thread = threading.Thread(target=office_checker)

#save side
delta_time = 0;
save_timer = 0;
reset_timer = 0;
has_edited = False;


def get_child_window(parent, class_name):
    if not parent:
        return

    global window_handle
    hwndChildList = []
    EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return [hwnd for hwnd in hwndChildList if class_name == GetClassName(hwnd) and abs(GetWindowRect(hwnd)[1] - GetWindowRect(window_handle)[1]) < 10]


def control_click(x, y, handle, button='left'):

    l_param = MAKELONG(x, y)

    if button == 'left':
        PostMessage(handle, WM_LBUTTONDOWN, MK_LBUTTON, l_param)
        time.sleep(0.1)
        PostMessage(handle, WM_LBUTTONUP, MK_LBUTTON, l_param)

    elif button == 'right':
        PostMessage(handle, WM_RBUTTONDOWN, 0, l_param)
        time.sleep(0.1)
        PostMessage(handle, WM_RBUTTONUP, 0, l_param)

def save_button_click():
    global save_timer, window_handle
    for handle in get_child_window(window_handle, "NetUIHWND"):
        control_click(136, 18, handle, 'left')

    print ("Clicked Save Button")

timer_pause = False
def save_loop():
    global save_timer, reset_timer, delta_time, saving, checking, final, has_edited
    while(saving):
        start_time = time.time();
        if (not timer_pause):
            save_timer += delta_time;

        if(current_office_window_still_open()):
            if(save_timer > 0.2 and has_edited):
                save_button_click()
                has_edited = False;
                save_timer = 0
        else:
            reset_timer += delta_time;
            if(reset_timer > 5):
                reset_timer = 0
                save_timer = 0
                saving = False
                checking = True
                print("Office Closed")
                word_checker_thread.run()
        delta_time = time.time() - start_time;

saving_thread = threading.Thread(target=save_loop)

#main thread side
def on_press(key):
    global save_timer, has_edited, timer_pause
    save_timer = 0;
    timer_pause = True
    has_edited = True;

def on_click(x, y, button, pressed):
    global timer_pause, has_edited
    if(not pressed):
        has_edited = True;

def on_move(x, y):
    global timer_pause, save_timer
    save_timer = 0;
    timer_pause = False

def quit_app():
    global tray_icon
    tray_icon.stop()
    os._exit(0)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


im = Image.open(resource_path("icon.png"))
tray_icon = trayicon('Office Auto Save', im,  menu=menu(item('Quit', quit_app, default=True)))

if __name__ == '__main__':

    #show hidden tray icon
    tray_icon.run_detached()
    toast = ToastNotifier()

    toast.show_toast(
    "Auto Save",
    "Office Auto Save is running in the background. Right click the icon in the system tray to quit.",
    duration = 6,
    icon_path = resource_path("icon.ico"),
    threaded = True,
    )

    keyboard.Listener(on_press=on_press).start()
    mouse.Listener(on_click=on_click).start()
    mouse.Listener(on_move=on_move).start()
    word_checker_thread.run()

    