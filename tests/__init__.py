# -*- coding: utf-8 -*-

from __future__ import print_function
import sys
import os
import unittest
import threading
import ctypes
from ctypes.wintypes import BOOL, LPARAM, WPARAM, INT, HWND, DWORD, HANDLE, UINT


BASE_PATH = os.path.dirname(__file__)

if not BASE_PATH:
    BASE_PATH = os.path.dirname(sys.argv[0])

IMPORT_PATH = os.path.abspath(os.path.join(BASE_PATH, '..'))

MB_OK = 0x00000000
PROCESS_QUERY_INFORMATION = 0x0400
MAX_PATH = 260
NULL = None
WM_CLOSE = 0x0010


user32 = ctypes.windll.User32
kernel32 = ctypes.windll.Kernel32

# BOOL PostMessageW(
#   HWND   hWnd,
#   UINT   Msg,
#   WPARAM wParam,
#   LPARAM lParam
# );
_PostMessage = user32.PostMessageW
_PostMessage.restype = BOOL

# int MessageBox(
#   HWND    hWnd,
#   LPCTSTR lpText,
#   LPCTSTR lpCaption,
#   UINT    uType
# );
_MessageBoxW = user32.MessageBoxW
_MessageBoxW.restype = INT


# BOOL CALLBACK EnumWindowsProc(
#   _In_ HWND   hwnd,
#   _In_ LPARAM lParam
# );
_EnumWindows = user32.EnumWindows
_EnumWindows.restype = BOOL

# BOOL IsWindow(
#   HWND hWnd
# );
IsWindow = user32.IsWindow
IsWindow.restype = BOOL

# int GetWindowTextLengthW(
#   HWND hWnd
# );
_GetWindowTextLength = user32.GetWindowTextLengthW
_GetWindowTextLength.restype = INT

# int GetWindowTextW(
#   HWND   hWnd,
#   LPWSTR lpString,
#   int    nMaxCount
# );
_GetWindowText = user32.GetWindowTextW
_GetWindowText.restype = INT

# DWORD GetWindowThreadProcessId(
#   HWND    hWnd,
#   LPDWORD lpdwProcessId
# );
_GetWindowThreadProcessId = user32.GetWindowThreadProcessId
_GetWindowThreadProcessId.restype = DWORD

# HANDLE OpenProcess(
#   DWORD dwDesiredAccess,
#   BOOL  bInheritHandle,
#   DWORD dwProcessId
# );
_OpenProcess = kernel32.OpenProcess
_OpenProcess.restype = HANDLE

# BOOL QueryFullProcessImageNameW(
#   HANDLE hProcess,
#   DWORD  dwFlags,
#   LPWSTR lpExeName,
#   PDWORD lpdwSize
# );
_QueryFullProcessImageName = kernel32.QueryFullProcessImageNameW
_QueryFullProcessImageName.restype = BOOL


WNDENUMPROC = ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)


def EnumWindows():
    windows = []

    def callback(hwnd, lparam):
        windows.append(hwnd)
        return 1

    lpEnumFunc = WNDENUMPROC(callback)
    lParam = NULL

    _EnumWindows(lpEnumFunc, lParam)

    return windows


def GetWindowText(hwnd):
    if not isinstance(hwnd, HWND):
        hwnd = HWND(hwnd)

    nMaxCount = _GetWindowTextLength(hwnd) + 1
    lpString = ctypes.create_unicode_buffer(nMaxCount)

    _GetWindowText(hwnd, lpString, nMaxCount)

    return lpString.value


def GetProcessName(hwnd):
    if not isinstance(hwnd, HWND):
        hwnd = HWND(hwnd)

    lpdwProcessId = DWORD()
    _GetWindowThreadProcessId(hwnd, ctypes.byref(lpdwProcessId))

    hProcess = _OpenProcess(
        DWORD(PROCESS_QUERY_INFORMATION),
        BOOL(False),
        lpdwProcessId
    )

    lpdwSize = DWORD(MAX_PATH)
    lpExeName = ctypes.create_string_buffer(MAX_PATH)

    _QueryFullProcessImageName(
        hProcess,
        DWORD(0),
        lpExeName,
        ctypes.byref(lpdwSize)
    )

    res = str(lpExeName)

    if res:
        res = os.path.split(res)[-1]
    return res


def MessageBox():
    def do():
        hWnd = NULL
        lpText = ctypes.create_unicode_buffer('UNITTESTS')
        lpCaption = ctypes.create_unicode_buffer('UNITTESTS')
        uType = UINT(MB_OK)
        _MessageBoxW(hWnd, lpText, lpCaption, uType)

    threading.Thread(target=do).start()


def PostMessage(hwnd, msg, wparam, lparam):
    if not isinstance(wparam, WPARAM):
        wparam = WPARAM(wparam)

    if not isinstance(lparam, LPARAM):
        lparam = LPARAM(lparam)

    return _PostMessage(hwnd, msg, wparam, lparam)


def Close():
    for hwnd in EnumWindows():
        if GetWindowText(hwnd) == 'UNITTESTS':
            PostMessage(hwnd, WM_CLOSE, 0, 0)
            break



pyWinVirtualDesktop = None
desktop_ids = []
desktops = []
new_desktop = None
hwnds = []
windows = []
new_window = None


class TestpyWinVirtualDesktop(unittest.TestCase):


    @classmethod
    def tearDownClass(cls):
        pass

    @classmethod
    def setUpClass(cls):
        pass

    def test_000_import(self):
        sys.path.insert(0, IMPORT_PATH)
        import pyWinVirtualDesktop as _pyWinVirtualDesktop
        global pyWinVirtualDesktop

        pyWinVirtualDesktop = _pyWinVirtualDesktop

    def test_010_desktop_ids(self):
        for id in pyWinVirtualDesktop.desktop_ids:
            print(str(id))
            desktop_ids.append(str(id))

    def test_020_enumerate_desktops(self):
        for desktop in pyWinVirtualDesktop:
            if str(desktop.id) not in desktop_ids:
                self.fail()
            desktops.append(desktop)

    def test_030_desktop_singleton(self):
        for desktop in pyWinVirtualDesktop:
            if desktop not in desktops:
                self.fail()

    def test_040_create_desktop(self):
        desktop = pyWinVirtualDesktop.create_desktop()
        desktop_guids = list(str(id) for id in pyWinVirtualDesktop.desktop_ids)

        if sorted(desktop_guids) != sorted(desktop_ids + [desktop.id]):
            self.fail()

        global new_desktop
        new_desktop = desktop

    def test_050_get_active_desktop(self):
        for desktop in pyWinVirtualDesktop:
            print(desktop.is_active)
            if desktop.is_active:
                break
        else:
            self.fail('No active desktop found.')

    def test_060_set_active_desktop(self):
        new_desktop.activate()

        if new_desktop != pyWinVirtualDesktop.current_desktop:
            self.fail('Desktops do not match.')

    def test_300_enumerate_windows(self):
        global hwnds

        hwnds = list(hwnd for hwnd in EnumWindows())

        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                if window.id not in hwnds:
                    self.fail()

                windows.append(window)

    def test_310_window_singleton(self):
        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                if window not in windows:
                    self.fail()

    def test_320_window_create_window(self):
        MessageBox()
        import time
        time.sleep(1.0)

        for hwnd in EnumWindows():
            if GetWindowText(hwnd) == 'UNITTESTS':
                print('Found Unittests message box')
                print(pyWinVirtualDesktop.current_desktop.add_window(hwnd))
                break

        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                if window.text == 'UNITTESTS':
                    global new_window
                    new_window = window
                    break
            else:
                continue

            break

        else:
            self.fail()

    def test_330_move_window(self):
        for desktop in pyWinVirtualDesktop:
            if not desktop.is_active:
                desktop.add_window(new_window)

                if new_window.desktop != desktop:
                    self.fail('Desktops do not match.')
                break
        else:
            self.fail()

    def test_340_window_on_active_desktop(self):
        if new_window.is_on_active_desktop is not False:
            self.fail()

        new_window.desktop.activate()
        if new_window.is_on_active_desktop is not True:
            self.fail()

    def test_999_end_test(self):
        Close()
        new_desktop.destroy()


if __name__ == '__main__':
    sys.argv.append('-v')
    unittest.main()
