# -*- coding: utf-8 -*-

from __future__ import print_function
import sys
import os
import unittest
import threading
import ctypes
from ctypes.wintypes import BOOL, LPARAM, INT, HWND, DWORD, HANDLE, UINT


BASE_PATH = os.path.dirname(__file__)

if not BASE_PATH:
    BASE_PATH = os.path.dirname(sys.argv[0])

IMPORT_PATH = os.path.abspath(os.path.join(BASE_PATH, '..'))

MB_OK = 0x00000000
PROCESS_QUERY_INFORMATION = 0x0400
MAX_PATH = 260
NULL = None

user32 = ctypes.windll.User32
kernel32 = ctypes.windll.Kernel32

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


pyWinVirtualDesktop = None


class TestpyWinVirtualDesktop(unittest.TestCase):
    desktop_ids = []
    desktops = []
    desktop = None
    hwnds = []
    windows = []
    window = None

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
            TestpyWinVirtualDesktop.desktop_ids += [str(id)]

    def test_020_enumerate_desktops(self):
        for desktop in pyWinVirtualDesktop:
            self.assertIn(str(desktop.id), TestpyWinVirtualDesktop.desktop_ids)
            TestpyWinVirtualDesktop.desktops += [desktop]

    def test_030_desktop_singleton(self):
        for desktop in pyWinVirtualDesktop:
            self.assertIn(desktop, TestpyWinVirtualDesktop.desktops)

    def test_040_create_desktop(self):
        desktop = pyWinVirtualDesktop.create_desktop()
        desktop_ids = list(str(id) for id in pyWinVirtualDesktop.desktop_ids)

        self.assertListEqual(
            sorted(desktop_ids),
            sorted(TestpyWinVirtualDesktop.desktop_ids + [desktop.id])
        )

        self.desktop = desktop

    def test_050_get_active_desktop(self):
        for desktop in pyWinVirtualDesktop:
            if desktop.is_active:
                break
        else:
            self.fail()

        self.assertEqual(desktop, TestpyWinVirtualDesktop.desktop)

    def test_060_set_active_desktop(self):
        TestpyWinVirtualDesktop.desktop.activate()

        self.assertEqual(pyWinVirtualDesktop.current_desktop, TestpyWinVirtualDesktop.desktop)

    def test_300_enumerate_windows(self):
        TestpyWinVirtualDesktop.hwnds = list(hwnd for hwnd in EnumWindows())

        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                self.assertIn(window.id, TestpyWinVirtualDesktop.hwnds)
                TestpyWinVirtualDesktop.windows += [window]

    def test_310_window_singleton(self):
        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                self.assertIn(window, TestpyWinVirtualDesktop.windows)

    def test_320_window_create_window(self):
        MessageBox()
        for desktop in pyWinVirtualDesktop:
            for window in desktop:
                if window.caption == 'UNITTESTS':
                    TestpyWinVirtualDesktop.window = window
                    break
        else:
            self.fail()

    def test_330_move_window(self):
        for desktop in pyWinVirtualDesktop:
            if not desktop.is_active:
                desktop.add_window(TestpyWinVirtualDesktop.window)
                self.assertEqual(TestpyWinVirtualDesktop.window.desktop, desktop)
                break
        else:
            self.fail()

    def test_340_window_on_active_desktop(self):
        self.assertEqual(TestpyWinVirtualDesktop.window.is_on_active_desktop, False)
        self.window.desktop.activate()
        self.assertEqual(TestpyWinVirtualDesktop.window.is_on_active_desktop, True)

    def test_999_end_test(self):
        TestpyWinVirtualDesktop.window.destroy()
        TestpyWinVirtualDesktop.desktop.destroy()
