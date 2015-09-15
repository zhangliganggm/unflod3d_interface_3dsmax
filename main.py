# -*- coding: utf8 -*-

import sys

import win32clipboard as w
import win32api, win32gui,win32con
import time

#obj_path = u"C:\\Users\\Administrator\\Desktop\\test.obj"

def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()


def get_unfold_hwnd():
    unfold_title = 'Unfold3D Demo'

    class_name = 'wxWindowClassNR'
    hwnd = win32gui.FindWindow(class_name, None)

    title = win32gui.GetWindowText(hwnd)
    print title
    if unfold_title in title:
        return hwnd


def callback(hwnd, hwnds):
    print win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd)
    #print
    if win32gui.GetWindowText(hwnd) == 'Export Mesh (.obj)':
        hwnds.append(hwnd)

    if win32gui.GetWindowText(hwnd) == 'Export Files':
        hwnds.append(hwnd)

    if win32gui.GetWindowText(hwnd) == 'Close':
        hwnds.append(hwnd)


def get_unfold_export_hwnds():

    def callback(hwnd, hwnds):
        if win32gui.GetWindowText(hwnd) == 'Stamper':
            hwnds.append(hwnd)

    hwnds = []

    win32gui.EnumWindows(callback, hwnds)
    return hwnds


def import_to_unfold(obj_path):

    u_hwnd = get_unfold_hwnd()
    win32gui.SetForegroundWindow(u_hwnd)

    time.sleep(0.1)

    win32api.keybd_event(17,0,0,0)
    win32api.keybd_event(79 ,0,0,0)
    win32api.keybd_event(79 ,0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)

    time.sleep(0.4)

    setText(obj_path.decode('gbk'))
    time.sleep(0.4)

    win32api.keybd_event(17, 0, 0,0)
    win32api.keybd_event(86, 0, 0,0)
    win32api.keybd_event(86, 0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(17, 0,win32con.KEYEVENTF_KEYUP,0)

    time.sleep(0.2)

    win32api.keybd_event(13, 0, 0,0)
    win32api.keybd_event(13, 0,win32con.KEYEVENTF_KEYUP,0)


def export_to_3dsmax():
    u_hwnd = get_unfold_hwnd()
    win32gui.SetForegroundWindow(u_hwnd)

    time.sleep(0.5)
    # ctrl+n

    win32api.keybd_event(17, 0, 0,0)
    win32api.keybd_event(78, 0, 0,0)
    win32api.keybd_event(78, 0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(17, 0,win32con.KEYEVENTF_KEYUP,0)

    time.sleep(0.5)
    #find export window

    export_window_hwnd = get_unfold_export_hwnds()
    print export_window_hwnd
    win32gui.SetForegroundWindow(export_window_hwnd[0])

    #find checkbox

    hh = []
    win32gui.EnumChildWindows(export_window_hwnd[0], callback, hh)

    print hh
    time.sleep(0.5)
    win32api.PostMessage(hh[0], win32con.WM_LBUTTONDOWN,0,0)
    win32api.PostMessage(hh[0], win32con.WM_LBUTTONUP,0,0)

    time.sleep(0.5)
    win32api.PostMessage(hh[1], win32con.WM_LBUTTONDOWN,0,0)
    win32api.PostMessage(hh[1], win32con.WM_LBUTTONUP,0,0)

    time.sleep(2)
    win32api.PostMessage(hh[2], win32con.WM_LBUTTONDOWN,0,0)
    win32api.PostMessage(hh[2], win32con.WM_LBUTTONUP,0,0)


if __name__ == '__main__':

    if len(sys.argv) > 2:
        self_path, mode, obj_path = sys.argv

        if mode == 'i':
            import_to_unfold(obj_path)

    elif len(sys.argv) == 2:
        self_path, mode = sys.argv
        if mode == 'o':
            export_to_3dsmax()
