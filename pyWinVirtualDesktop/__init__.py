# -*- coding: utf-8 -*-
from __future__ import print_function
import sys
import six
import libWinVirtualDesktop
from .winuser import (
    GetWindowText,
    IsWindow,
    GetProcessName,
    EnumWindows
)


class Module(object):

    def __init__(self):
        mod = sys.modules[__name__]
        sys.modules[__name__] = self
        self.__original_module__ = mod

        self.Desktop = Desktop
        self.Window = Window

    @property
    def current_desktop(self):
        desktop_guid = (
            libWinVirtualDesktop.DesktopManagerInternalGetCurrentDesktop()
        )
        if desktop_guid:
            return Desktop(desktop_guid)

    @property
    def desktop_ids(self):
        desktop_ids = (
            libWinVirtualDesktop.DesktopManagerInternalGetDesktopIds()
        )
        return desktop_ids

    def create_desktop(self):
        desktop_guid = (
            libWinVirtualDesktop.DesktopManagerInternalCreateDesktop()
        )
        if desktop_guid:
            return Desktop(desktop_guid)

    def register_notification_callback(self, callback):

        try:
            if issubclass(callback, DesktopNotificationCallback):
                callback = callback()
        except TypeError:
            if not isinstance(callback, DesktopNotificationCallback):
                raise RuntimeError(
                    'callback needs to be an instance or a subclass of '
                    'VirtualDesktopNotificationCallback'
                )

    def unregister_notification_callback(self, cookie):
        pass

    def __len__(self):
        return libWinVirtualDesktop.DesktopManagerInternalGetCount()

    def __iter__(self):
        for desktop_guid in self.desktop_ids:
            yield Desktop(desktop_guid)


class InstanceSingleton(type):

    def __init__(cls, *args, **kwargs):
        """
        InstanceSingleton metaclass constructor.
        """
        super(InstanceSingleton, cls).__init__(*args, **kwargs)

        cls._instances = {}

    def __call__(cls, id, **kwargs):

        key = [id] + list(kwargs[k] for k in sorted(kwargs.keys()))
        if key not in cls._instances:
            cls._instances[key] = (
                super(InstanceSingleton, cls).__call__(id, **kwargs)
            )

        return cls._instances[key]


@six.add_metaclass(InstanceSingleton)
class View(object):

    def __init__(self, window_handle):
        self.__id = window_handle

    def move_to_desktop(self, desktop):
        if self.is_ok:
            if isinstance(desktop, Desktop):
                desktop = desktop.id

            can_move = libWinVirtualDesktop.DesktopManagerInternalCanViewMoveDesktops(self.id)
            if can_move > 0:
                libWinVirtualDesktop.DesktopManagerInternalMoveViewToDesktop(desktop, self.id)

    @property
    def id(self):
        return self.__id

    @property
    def is_ok(self):
        return IsWindow(self.id)

    @property
    def thumbnail_handle(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetThumbnailWindow(self.id)

    @property
    def has_focus(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetFocus(self.id)

    def set_focus(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewSetFocus(self.id)

    def flash(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewFlash(self.id)

    def activate(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewSwitchTo(self.id)

    @property
    def is_mirrored(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewIsMirrored(self.id)

    @property
    def is_splash_screen_presented(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewIsSplashScreenPresented(self.id)

    @property
    def is_tray(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewIsTray(self.id)

    @property
    def can_receive_input(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewCanReceiveInput(self.id)

    @property
    def scale_factor(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetScaleFactor(self.id)

    @property
    def show_in_switchers(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetShowInSwitchers(self.id)

    @show_in_switchers.setter
    def show_in_switchers(self, value):
        if self.is_ok:
            libWinVirtualDesktop.ApplicationViewSetShowInSwitchers(self.id, value)

    @property
    def desktop(self):
        if self.is_ok:
            desktop_guid = libWinVirtualDesktop.ApplicationViewGetVirtualDesktopId(self.id)
            if desktop_guid:
                return Desktop(desktop_guid)

    @property
    def state(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetViewState(self.id)

    @state.setter
    def state(self, value):
        if self.is_ok:
            libWinVirtualDesktop.ApplicationViewSetViewState(self.id, value)

    @property
    def is_visible(self):
        if self.is_ok:
            return libWinVirtualDesktop.ApplicationViewGetVisibility(self.id)

    @property
    def size(self):
        rect = libWinVirtualDesktop.ApplicationViewGetExtendedFramePosition(self.id)
        if rect:
            return rect['right'] - rect['left'], rect['bottom'] - rect['top']
        return None, None

    @property
    def position(self):
        rect = libWinVirtualDesktop.ApplicationViewGetExtendedFramePosition(self.id)

        if rect:
            return rect['left'], rect['top']
        return None, None

    @property
    def pinned(self):
        if self.is_ok:
            desktop_guid = (
                libWinVirtualDesktop.DesktopManagerGetWindowDesktopId(
                    self.id
                )
            )
            if desktop_guid:
                return libWinVirtualDesktop.DesktopPinnedAppsIsViewPinned(desktop_guid, self.id)

    @pinned.setter
    def pinned(self, value):
        if self.is_ok:
            desktop_guid = (
                libWinVirtualDesktop.DesktopManagerGetWindowDesktopId(
                    self.id
                )
            )
            if desktop_guid:
                if value:
                    libWinVirtualDesktop.DesktopPinnedAppsPinView(desktop_guid, self.id)
                else:
                    libWinVirtualDesktop.DesktopPinnedAppsUnpinView(desktop_guid, self.id)


@six.add_metaclass(InstanceSingleton)
class Window(object):

    def __init__(self, hwnd):
        self.__hwnd = hwnd

    @property
    def id(self):
        return self.__hwnd

    @property
    def text(self):
        if self.is_ok:
            return GetWindowText(self.id)

        return ''

    @property
    def is_ok(self):
        return IsWindow(self.id)

    @property
    def desktop(self):
        if self.is_ok:
            desktop_guid = (
                libWinVirtualDesktop.DesktopManagerGetWindowDesktopId(
                    self.__hwnd
                )
            )

            return Desktop(desktop_guid)

    @property
    def is_on_active_desktop(self):
        if self.is_ok:
            return libWinVirtualDesktop.DesktopManagerIsWindowOnCurrentVirtualDesktop(self.id)

    @property
    def process_name(self):
        if self.is_ok:
            return GetProcessName(self.id)

    @property
    def view(self):
        return View(self.id)

    def move_to_desktop(self, desktop):
        if isinstance(desktop, Desktop):
            desktop = desktop.id

        if self.is_ok:
            return libWinVirtualDesktop.DesktopManagerMoveWindowToDesktop(desktop, self.id)


@six.add_metaclass(InstanceSingleton)
class Desktop(object):

    def __init__(self, desktop_guid):
        self._id = desktop_guid

    @property
    def number(self):
        return libWinVirtualDesktop.GetDesktopNumberFromId(self.id)

    @property
    def id(self):
        return self._id

    @property
    def desktop_to_left(self):
        desktop_guid = libWinVirtualDesktop.DesktopManagerInternalGetAdjacentDesktop(self.id, 3)
        if desktop_guid:
            return Desktop(desktop_guid)

    @property
    def desktop_to_right(self):
        desktop_guid = libWinVirtualDesktop.DesktopManagerInternalGetAdjacentDesktop(self.id, 4)
        if desktop_guid:
            return Desktop(desktop_guid)

    def add_window(self, window):
        if not isinstance(window, Window):
            window = Window(window)

        window.move_to_desktop(self)

    def __iter__(self):
        for hwnd in EnumWindows():
            desktop_guid = (
                libWinVirtualDesktop.DesktopManagerGetWindowDesktopId(hwnd)
            )

            if desktop_guid == self.id:
                yield Window(hwnd)

    @property
    def is_active(self):
        current_desktop = pyWinVirtualDesktop.current_desktop
        return current_desktop == self

    def activate(self):
        return libWinVirtualDesktop.DesktopManagerInternalSwitchDesktop(self.id)

    def delete(self):
        for desktop_id in pyWinVirtualDesktop.desktop_ids:
            if desktop_id != self.id:
                break
        else:
            raise RuntimeError(
                'You cannot delete the only Virtual Desktop.'
            )

        return libWinVirtualDesktop.DesktopManagerInternalRemoveDesktop(self.id, desktop_id)


class DesktopNotificationCallback(object):

    def change(self, old, new):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.change method to '
            'not see this message.'
        )
        print('OLD:', old.id)
        print('NEW:', new.id)

    def create(self, new):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.create method to '
            'not see this message.'
        )
        print('NEW:', new.id)

    def destroy_begin(self, destroyed, fallback):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroy_begin method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed.id)
        print('FALLBACK:', fallback.id)

    def destroy_failed(self, destroyed, fallback):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroy_failed method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed.id)
        print('FALLBACK:', fallback.id)

    def destroy(self, destroyed, fallback):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroyed method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed.id)
        print('FALLBACK:', fallback.id)

    def view_changed(self, view):
        pass


S_OK = 0x00000000


class VirtualDesktopNotification(object):

    def __init__(self, callback):
        self.__callback = callback

    def CurrentVirtualDesktopChanged(self, pDesktopOld, pDesktopNew):
        desktop_old = Desktop(pDesktopOld)

        desktop_new = Desktop(pDesktopNew)

        self.__callback.change(desktop_old, desktop_new)

        return S_OK

    def VirtualDesktopCreated(self, pDesktop):
        desktop = Desktop(pDesktop)

        self.__callback.create(desktop)

        return S_OK

    def VirtualDesktopDestroyBegin(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(pDesktopDestroyed)
        desktop_fallback = Desktop(pDesktopFallback)

        self.__callback.destroy_begin(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def VirtualDesktopDestroyFailed(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(pDesktopDestroyed)
        desktop_fallback = Desktop(pDesktopFallback)

        self.__callback.destroy_failed(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def VirtualDesktopDestroyed(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(pDesktopDestroyed)

        desktop_fallback = Desktop(pDesktopFallback)

        self.__callback.destroy(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def ViewVirtualDesktopChanged(self, pView):
        return S_OK


desktop_ids = Module.desktop_ids

pyWinVirtualDesktop = Module()
create_desktop = pyWinVirtualDesktop.create_desktop
