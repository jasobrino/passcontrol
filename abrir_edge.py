import win32gui
import win32con
import psutil
import win32process
import subprocess
from comun import main_url

winlist, toplist = [], []

def get_edge_pid():
    psutil.process_iter.cache_clear()
    for p in psutil.process_iter(['pid','name','username']):
        if p.info['name'] == 'msedge.exe':
            pid = p.info['pid']
            #parent = psutil.Process(pid).parent()
            #print(f"Parent PID: {parent.pid}, Name: {parent.name()} ({parent.username()}) {parent.cmdline()} ")
            if main_url in  p.cmdline():
                print(f"localizada main_url: {p.cmdline()}")
                return pid
    return None

def abrir_edge(bin_location):

    def add_window(hwnd, results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(add_window, toplist)
    pid_edge = get_edge_pid()

    windows = [(hwnd, title) for hwnd, title in winlist if pid_edge == win32process.GetWindowThreadProcessId(hwnd)[1] and title.lower().startswith('pass gestión')]
    if windows:
        hwnd = windows[0][0]
        print("ventana localizada:", hwnd)     
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    else:
        print("Se procede a abrir nueva ventana 'Pass Gestión'")
        subprocess.Popen([bin_location, main_url])
