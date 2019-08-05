# -*- coding: utf-8 -*-
from __future__ import print_function
import os
import sys
import ctypes
from ctypes.wintypes import (
    BOOL,
    LPARAM,
    INT,
    HWND,
    DWORD,
    HANDLE,
    RECT,
    WPARAM
)

user32 = ctypes.windll.User32
kernel32 = ctypes.windll.Kernel32

NULL = None
MAX_PATH = 260
PROCESS_QUERY_INFORMATION = 0x0400
SW_FORCEMINIMIZE = 11
SW_HIDE = 0
SW_MAXIMIZE = 3
SW_MINIMIZE = 6
SW_RESTORE = 9
SW_SHOW = 5
SW_SHOWDEFAULT = 10
SW_SHOWMAXIMIZED = 3
SW_SHOWMINIMIZED = 2
SW_SHOWMINNOACTIVE = 7
SW_SHOWNA = 8
SW_SHOWNOACTIVATE = 4
SW_SHOWNORMAL = 1
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_SHOWWINDOW = 0x0040

HWND_TOP = 0

WM_CLOSE = 0x0010


# BOOL BringWindowToTop(
#   HWND hWnd
# );
BringWindowToTop = user32.BringWindowToTop
BringWindowToTop.restype = BOOL

# HWND GetFocus();
GetFocus = user32.GetFocus
GetFocus.restype = HWND

# HWND SetFocus(
#   HWND hWnd
# );
SetFocus = user32.SetFocus
SetFocus.restype = HWND

# BOOL PostMessageW(
#   HWND   hWnd,
#   UINT   Msg,
#   WPARAM wParam,
#   LPARAM lParam
# );

_PostMessage = user32.PostMessageW
_PostMessage.restype = BOOL

# BOOL SetWindowPos(
#   HWND hWnd,
#   HWND hWndInsertAfter,
#   int  X,
#   int  Y,
#   int  cx,
#   int  cy,
#   UINT uFlags
# );
_SetWindowPos = user32.SetWindowPos
_SetWindowPos.restype = BOOL

# BOOL IsIconic(
#   HWND hWnd
# );
IsIconic = user32.IsIconic
IsIconic.restype = BOOL

# BOOL IsZoomed(
#   HWND hWnd
# );
IsZoomed = user32.IsZoomed
IsZoomed.restype = BOOL

# BOOL IsWindowVisible(
#   HWND hWnd
# );
IsWindowVisible = user32.IsWindowVisible
IsWindowVisible.restype = BOOL

# BOOL CALLBACK EnumWindowsProc(
#   _In_ HWND   hwnd,
#   _In_ LPARAM lParam
# );
_EnumWindows = user32.EnumWindows
_EnumWindows.restype = BOOL

# BOOL DestroyWindow(
#   HWND hWnd
# );
DestroyWindow = user32.DestroyWindow
DestroyWindow.restype = BOOL

# BOOL GetWindowRect(
#   HWND   hWnd,
#   LPRECT lpRect
# );

_GetWindowRect = user32.GetWindowRect
_GetWindowRect.restype = BOOL

# BOOL ShowWindow(
#   HWND hWnd,
#   int  nCmdShow
# );
ShowWindow = user32.ShowWindow
ShowWindow.restype = BOOL

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

    if sys.version_info[0] == 2:
        res = ''
    else:
        res = b''

    for i in range(260):
        if sys.version_info[0] == 2:
            if lpExeName[i] == '\x00':
                continue
        else:
            if lpExeName[i] == b'\x00':
                continue

        res += lpExeName[i]

    if res:
        res = os.path.split(res)[-1]
    return res


def PostMessage(hwnd, msg, wparam, lparam):
    if not isinstance(wparam, WPARAM):
        wparam = WPARAM(wparam)

    if not isinstance(lparam, LPARAM):
        lparam = LPARAM(lparam)

    return _PostMessage(hwnd, msg, wparam, lparam)


def SetWindowPos(hwnd, pos=None, size=None):
    if pos is None:
        uFlags = SWP_NOMOVE
        width, height = size
        x = INT(0)
        y = INT(0)
        cx = INT(width)
        cy = INT(height)
    else:
        uFlags = SWP_NOSIZE
        x, y = pos
        x = INT(x)
        y = INT(y)
        cx = INT(0)
        cy = INT(0)

    hWndInsertAfter = HWND_TOP

    if not IsWindowVisible(hwnd):
        uFlags |= SWP_SHOWWINDOW

    return _SetWindowPos(hwnd, hWndInsertAfter, x, y, cx, cy, uFlags)


def GetWindowRect(hwnd):

    rect = RECT()

    _GetWindowRect(hwnd, ctypes.byref(rect))
    return rect


if __name__ == '__main__':
    for hwnd in EnumWindows():
        name = GetProcessName(hwnd)

        print(name, repr(name))
