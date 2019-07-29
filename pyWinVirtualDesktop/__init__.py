# -*- coding: utf-8 -*-

import sys
import comtypes
import ctypes
import six
from ctypes import POINTER

from .winuser import EnumWindows, IsWindow, GetWindowText, GetProcessName
from .servprov import IServiceProvider, CLSID_ImmersiveShell
from .shobjidl_core import (
    CLSID_VirtualDesktopManagerInternal,
    IID_IVirtualDesktopManagerInternal,
    IVirtualDesktopManagerInternal,
    IVirtualDesktopManager,
    IID_IVirtualDesktop,
    IVirtualDesktop,
    AdjacentDesktop
)


class Module(object):

    def __init__(self):
        mod = sys.modules[__name__]
        sys.modules[__name__] = self
        self.__original_module__ = mod

        comtypes.CoInitialize()

        self.__pServiceProvider = comtypes.CoCreateInstance(
            CLSID_ImmersiveShell,
            IServiceProvider,
            comtypes.CLSCTX_LOCAL_SERVER,
        )

        self.__pDesktopManagerInternal = comtypes.cast(
            self.__pServiceProvider.QueryService(
                CLSID_VirtualDesktopManagerInternal,
                IID_IVirtualDesktopManagerInternal
            ),
            ctypes.POINTER(IVirtualDesktopManagerInternal)
        )

        self.__pDesktopManager = self.__pServiceProvider.QueryInterface(
            IVirtualDesktopManager
        )

    @property
    def desktop_ids(self):
        desktop_ids = []
        pObjectArray = self.__pDesktopManagerInternal.GetDesktops()

        for i in range(pObjectArray.GetCount()):
            pDesktop = POINTER(IVirtualDesktop)()
            pObjectArray.GetAt(i,
                IID_IVirtualDesktop.ctypes.byref(IVirtualDesktop))

            id = pDesktop.GetID()
            desktop_ids += [id]

        return desktop_ids

    def create_desktop(self):
        ppNewDesktop = self.__pDesktopManagerInternal.CreateDesktopW()
        return Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            ppNewDesktop.GetId()
        )

    def __iter__(self):
        for id in self.desktop_ids:
            yield Desktop(
                self.__pDesktopManagerInternal,
                self.__pDesktopManager,
                id
            )


class InstanceSingleton(type):
    _instances = {}

    def __new__(mcs, name, bases, dct):
        cls = super(InstanceSingleton, mcs).__new__(name, bases, dct)
        cls._instances = {}

        return cls

    def __call__(cls, pDesktopManagerInternal, pDesktopManager, id):

        key = (pDesktopManagerInternal, pDesktopManager, str(id))

        if key not in cls._instances:
            cls._instances[key] = (
                super(InstanceSingleton, cls).__call__(
                    pDesktopManagerInternal,
                    pDesktopManager,
                    id
                )
            )

        return cls._instances[key]


@six.with_metaclass(InstanceSingleton)
class Window(object):

    def __init__(self, pDesktopManagerInternal, pDesktopManager, hwnd):
        self.__pDesktopManagerInternal = pDesktopManagerInternal
        self.__pDesktopManager = pDesktopManager
        self.__hwnd = hwnd

    @property
    def id(self):
        return self.__hwnd

    @property
    def text(self):
        if self.is_ok:
            return GetWindowText(self.__hwnd)

        return ''

    @property
    def is_ok(self):
        return IsWindow(self.__hwnd)

    @property
    def desktop(self):
        if self.is_ok:
            desktop = self.__pDesktopManager.GetWindowDesktopId(self.__hwnd)

            return Desktop(
                self.__pDesktopManagerInternal,
                self.__pDesktopManager,
                desktop.GetId()
            )

    @property
    def is_on_active_desktop(self):
        if self.is_ok:
            return self.__pDesktopManager.IsWindowOnCurrentVirtualDesktop(
                self.__hwnd
            )

    @property
    def process_name(self):
        if self.is_ok:
            return GetProcessName(self.__hwnd)

    def move_to_desktop(self, desktop):
        if isinstance(desktop, Desktop):
            desktop = desktop.id

        if self.is_ok:
            self.__pDesktopManager.MoveWindowToDesktop(
                self.__hwnd,
                desktop
            )


@six.with_metaclass(InstanceSingleton)
class Desktop(object):

    def __init__(self, pDesktopManagerInternal, pDesktopManager, id):
        self.__pDesktopManagerInternal = pDesktopManagerInternal
        self.__pDesktopManager = pDesktopManager
        self._id = id

    @property
    def id(self):
        return self._id

    @property
    def desktop_to_left(self):
        desktop = self.__pDesktopManagerInternal.FindDesktop(
            ctypes.byref(self.id)
        )
        neighbor = self.__pDesktopManagerInternal.GetAdjacentDesktop(
            desktop,
            AdjacentDesktop.LeftDirection
        )

        return Desktop(self.__pDesktopManagerInternal, neighbor.GetId())

    @property
    def desktop_to_right(self):
        desktop = self.__pDesktopManagerInternal.FindDesktop(
            ctypes.byref(self.id)
        )
        neighbor = self.__pDesktopManagerInternal.GetAdjacentDesktop(
            desktop,
            AdjacentDesktop.RightDirection
        )

        return Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            neighbor.GetId()
        )

    def add_window(self, window):
        if not isinstance(window, Window):
            window = Window(
                self.__pDesktopManagerInternal,
                self.__pDesktopManager,
                window
            )

        window.move_to_desktop(self)

    def __iter__(self):

        for hwnd in EnumWindows():
            id = self.__pDesktopManager.GetWindowDesktopId(hwnd)
            if str(id) == str(self.id):
                yield Window(
                    self.__pDesktopManagerInternal,
                    self.__pDesktopManager,
                    hwnd
                )

    @property
    def is_active(self):
        current_desktop = self.__pDesktopManagerInternal.GetCurrentDesktop()
        return str(current_desktop.GetId()) == str(self.id)

    def activate(self):
        desktop = self.__pDesktopManagerInternal.FindDesktop(
            ctypes.byref(self.id)
        )
        self.__pDesktopManagerInternal.SwitchDesktop(desktop)

    def delete(self):
        pyWinVirtualDesktop = __import__(__name__.split('.')[0])

        for id in pyWinVirtualDesktop.desktop_ids:
            if str(id) != str(self.id):
                break
        else:
            raise RuntimeError(
                'You cannot delete the only Virtual Desktop.')

        pRemove = self.__pDesktopManagerInternal.FindDesktop(
            ctypes.byref(self.id)
        )
        pFallbackDesktop = self.__pDesktopManagerInternal.FindDesktop(
            ctypes.byref(id)
        )
        self.__pDesktopManagerInternal.RemoveDesktop(
            pRemove,
            pFallbackDesktop
        )

    def add_view(self, pView):

        if self.__pDesktopManagerInternal.CanViewMoveDesktops(pView):
            pDesktop = self.__pDesktopManagerInternal.FindDesktop(
                ctypes.byref(self.id)
            )

            self.__pDesktopManagerInternal.MoveViewToDesktop(
                pView,
                pDesktop
            )


desktop_ids = Module.desktop_ids
create_desktop = Module.create_desktop

Module()



