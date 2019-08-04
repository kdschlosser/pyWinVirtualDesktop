# -*- coding: utf-8 -*-
from __future__ import print_function
import os
import sys
import ctypes
from ctypes.wintypes import BOOL, LPARAM, INT, HWND, DWORD, HANDLE

user32 = ctypes.windll.User32
NULL = None

# BOOL CALLBACK EnumWindowsProc(
#   _In_ HWND   hwnd,
#   _In_ LPARAM lParam
# );


_EnumWindows = user32.EnumWindows
_EnumWindows.restype = BOOL

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


def GetWindowText(hwnd):
    if not isinstance(hwnd, HWND):
        hwnd = HWND(hwnd)

    nMaxCount = _GetWindowTextLength(hwnd) + 1
    lpString = ctypes.create_unicode_buffer(nMaxCount)

    _GetWindowText(hwnd, lpString, nMaxCount)

    return lpString.value


# DWORD GetWindowThreadProcessId(
#   HWND    hWnd,
#   LPDWORD lpdwProcessId
# );


_GetWindowThreadProcessId = user32.GetWindowThreadProcessId
_GetWindowThreadProcessId.restype = DWORD

MAX_PATH = 260

kernel32 = ctypes.windll.Kernel32

# HANDLE OpenProcess(
#   DWORD dwDesiredAccess,
#   BOOL  bInheritHandle,
#   DWORD dwProcessId
# );

_OpenProcess = kernel32.OpenProcess
_OpenProcess.restype = HANDLE

PROCESS_QUERY_INFORMATION = 0x0400

# BOOL QueryFullProcessImageNameW(
#   HANDLE hProcess,
#   DWORD  dwFlags,
#   LPWSTR lpExeName,
#   PDWORD lpdwSize
# );


_QueryFullProcessImageName = kernel32.QueryFullProcessImageNameW
_QueryFullProcessImageName.restype = BOOL


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


if __name__ == '__main__':
    for hwnd in EnumWindows():
        name = GetProcessName(hwnd)

        print(name, repr(name))
