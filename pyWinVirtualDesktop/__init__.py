    # -*- coding: utf-8 -*-
from __future__ import print_function
import sys
import six
import traceback
from .config import Config
import uuid

import libWinVirtualDesktop as _libWinVirtualDesktop
from .winuser import (
    GetWindowText,
    IsWindow,
    GetProcessName,
    EnumWindows,
    DestroyWindow,
    SW_HIDE,
    SW_MAXIMIZE,
    SW_MINIMIZE,
    SW_RESTORE,
    SW_SHOW,
    SW_SHOWNORMAL,
    SW_SHOWMINIMIZED,
    SW_SHOWMAXIMIZED,
    ShowWindow,
    IsWindowVisible,
    IsIconic,
    IsZoomed,
    SetWindowPos,
    GetWindowRect,
    WM_CLOSE,
    PostMessage,
    GetFocus,
    SetFocus,
    BringWindowToTop
)

VIRTUAL_DESKTOP_CREATED = 5
VIRTUAL_DESKTOP_DESTROY_BEGIN = 4
VIRTUAL_DESKTOP_DESTROY_FAILED = 3
VIRTUAL_DESKTOP_DESTROYED = 2
VIRTUAL_DESKTOP_VIEW_CHANGED = 1
VIRTUAL_DESKTOP_CURRENT_CHANGED = 0


class libWinVirtualDesktop(object):

    def __getattr__(self, item):
        func = getattr(_libWinVirtualDesktop, item)

        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except:
                print(traceback.format_exc())

        return wrapper


libWinVirtualDesktop = libWinVirtualDesktop()


class Module(object):

    def __init__(self):
        mod = sys.modules[__name__]
        sys.modules[__name__] = self
        self.__original_module__ = mod

        self.Desktop = Desktop
        self.Window = Window
        self.DesktopNotificationCallback = DesktopNotificationCallback
        self.config = Config
        self.__callbacks = {}

    def __notification_callback(self, notif_type, desktop1_id, dekstop2_id, thumb_hwnd):
        if notif_type == VIRTUAL_DESKTOP_CREATED:
            desktop = Desktop(desktop1_id)
            for callback in list(self.__callbacks.values())[:]:
                callback.create(desktop)

        elif notif_type == VIRTUAL_DESKTOP_VIEW_CHANGED:
            desktop = Desktop(desktop1_id)
            for window in desktop:
                if window.view.thumbnail_handle == thumb_hwnd:
                    break
            else:
                window = None

            for callback in list(self.__callbacks.values())[:]:
                callback.view_changed(desktop, window)

        else:
            mapping = {
                VIRTUAL_DESKTOP_DESTROY_BEGIN: 'destroy_begin',
                VIRTUAL_DESKTOP_DESTROY_FAILED: 'destroy_failed',
                VIRTUAL_DESKTOP_DESTROYED: 'destroy',
                VIRTUAL_DESKTOP_CURRENT_CHANGED: 'change'
            }

            if desktop1_id:
                desktop1 = Desktop(desktop1_id)
            else:
                desktop1 = None

            if dekstop2_id:
                desktop2 = Desktop(dekstop2_id)
            else:
                desktop2 = None

            for callback in list(self.__callbacks.values())[:]:
                getattr(callback, mapping[notif_type])(desktop1, desktop2)

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

        if not self.__callbacks:
            libWinVirtualDesktop.RegisterDesktopNotifications(
                self.__notification_callback
            )

        guid = uuid.uuid4()
        self.__callbacks[guid] = callback
        return guid

    def unregister_notification_callback(self, cookie):
        if cookie in self.__callbacks:
            del self.__callbacks[cookie]

        if not self.__callbacks:
            libWinVirtualDesktop.UnregisterDesktopNotifications()

    def __contains__(self, desktop):
        if isinstance(desktop, Desktop):
            desktop = desktop.id

        return desktop in self.desktop_ids

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
        key = tuple(key)

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
            res = libWinVirtualDesktop.ApplicationViewGetFocus(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

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
            res = libWinVirtualDesktop.ApplicationViewIsMirrored(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

    @property
    def is_splash_screen_presented(self):
        if self.is_ok:
            res = libWinVirtualDesktop.ApplicationViewIsSplashScreenPresented(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

    @property
    def is_tray(self):
        if self.is_ok:
            res = libWinVirtualDesktop.ApplicationViewIsTray(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

    @property
    def can_receive_input(self):
        if self.is_ok:
            res = libWinVirtualDesktop.ApplicationViewCanReceiveInput(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

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
            res = libWinVirtualDesktop.ApplicationViewGetVisibility(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

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
                res = libWinVirtualDesktop.DesktopPinnedAppsIsViewPinned(desktop_guid, self.id)
                if res == 1:
                    return True
                elif res == 0:
                    return False

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

    def bring_to_top(self):
        if self.is_ok:
            return bool(BringWindowToTop(self.id))

        return False

    def restore(self):
        if self.is_ok:
            if self.is_shown:
                return bool(ShowWindow(self.id, SW_RESTORE))
            else:
                return bool(ShowWindow(self.id, SW_SHOW))

        return False

    def post_message(self, msg, wparam=0, lparam=0):
        if self.is_ok:
            return PostMessage(self.id, msg, wparam, lparam)
        return -1

    def close(self):
        if self.is_ok:
            return bool(self.post_message(WM_CLOSE))
        return False

    @property
    def has_focus(self):
        if self.is_ok:
            return GetFocus() == self.id
        return False

    def focus(self):
        if self.is_ok:
            return SetFocus(self.id) == self.id
        return False

    @property
    def is_minimized(self):
        if self.is_ok and self.is_shown:
            return bool(IsIconic(self.id))
        return False

    @is_minimized.setter
    def is_minimized(self, value):
        if self.is_ok:
            if self.is_shown:
                if value:
                    ShowWindow(self.id, SW_MINIMIZE)
                else:
                    ShowWindow(self.id, SW_RESTORE)
            else:
                if value:
                    ShowWindow(self.id, SW_SHOWMINIMIZED)
                else:
                    ShowWindow(self.id, SW_RESTORE)

    @property
    def is_maximized(self):
        if self.is_ok and self.is_shown:
            return bool(IsZoomed(self.id))
        return False

    @is_maximized.setter
    def is_maximized(self, value):
        if self.is_ok:
            if self.is_shown:
                if value:
                    ShowWindow(self.id, SW_MAXIMIZE)
                else:
                    ShowWindow(self.id, SW_RESTORE)
            else:
                if value:
                    ShowWindow(self.id, SW_SHOWMAXIMIZED)
                else:
                    ShowWindow(self.id, SW_RESTORE)

    @property
    def is_shown(self):
        if self.is_ok:
            return bool(IsWindowVisible(self.id))
        return False

    @is_shown.setter
    def is_shown(self, value):
        if self.is_ok:
            if not value and self.is_shown:
                ShowWindow(self.id, SW_HIDE)
            elif value and not self.is_shown:
                ShowWindow(self.id, SW_SHOW)

    @property
    def size(self):
        if self.is_ok:
            rect = GetWindowRect(self.id)
            return rect.right - rect.left, rect.bottom - rect.top
        return None, None

    @size.setter
    def size(self, value):
        if self.is_ok:
            SetWindowPos(self.id, size=value)

    @property
    def position(self):
        if self.is_ok:
            rect = GetWindowRect(self.id)
            return rect.left, rect.top
        return None, None

    @position.setter
    def position(self, value):
        if self.is_ok:
            SetWindowPos(self.id, pos=value)

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
            res = libWinVirtualDesktop.DesktopManagerIsWindowOnCurrentVirtualDesktop(self.id)
            if res == 1:
                return True
            elif res == 0:
                return False

    @property
    def process_name(self):
        if self.is_ok:
            return GetProcessName(self.id)

    @property
    def view(self):
        return View(self.id)

    def destroy(self):
        if self.is_ok:
            return DestroyWindow(self.id)

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
    def name(self):
        return pyWinVirtualDesktop.config.get_name(self.id)

    @name.setter
    def name(self, value):
        pyWinVirtualDesktop.config.set_name(self.id, value)

    @property
    def number(self):
        return pyWinVirtualDesktop.desktop_ids.index(self.id) + 1

    @property
    def id(self):
        return self._id

    @property
    def desktop_to_left(self):
        desktop_ids = pyWinVirtualDesktop.desktop_ids
        index = desktop_ids.index(self.id)

        try:
            return Desktop(desktop_ids[index - 1])
        except IndexError:
            pass

    @property
    def desktop_to_right(self):
        desktop_ids = pyWinVirtualDesktop.desktop_ids
        index = desktop_ids.index(self.id)

        try:
            return Desktop(desktop_ids[index + 1])
        except IndexError:
            pass

    def add_window(self, window):
        if not isinstance(window, Window):
            window = Window(window)

        window.move_to_desktop(self)

    def __contains__(self, window):
        if isinstance(window, Window):
            window = window.id

        desktop_guid = (
            libWinVirtualDesktop.DesktopManagerGetWindowDesktopId(window)
        )
        return desktop_guid == self.id

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

    def destroy(self):
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

        if old is None:
            old_id = 'None'
        else:
            old_id = old.id

        if new is None:
            new_id = 'None'
        else:
            new_id = new.id

        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.change method to '
            'not see this message.'
        )
        print('OLD:', old_id)
        print('NEW:', new_id)

    def create(self, new):
        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.create method to '
            'not see this message.'
        )
        print('NEW:', new.id)

    def destroy_begin(self, destroyed, fallback):

        if destroyed is None:
            destroyed_id = 'None'
        else:
            destroyed_id = destroyed.id

        if fallback is None:
            fallback_id = 'None'
        else:
            fallback_id = fallback.id

        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroy_begin method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed_id)
        print('FALLBACK:', fallback_id)

    def destroy_failed(self, destroyed, fallback):
        if destroyed is None:
            destroyed_id = 'None'
        else:
            destroyed_id = destroyed.id

        if fallback is None:
            fallback_id = 'None'
        else:
            fallback_id = fallback.id

        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroy_failed method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed_id)
        print('FALLBACK:', fallback_id)

    def destroy(self, destroyed, fallback):
        if destroyed is None:
            destroyed_id = 'None'
        else:
            destroyed_id = destroyed.id

        if fallback is None:
            fallback_id = 'None'
        else:
            fallback_id = fallback.id

        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.destroyed method to '
            'not see this message.'
        )
        print('DESTROYED:', destroyed_id)
        print('FALLBACK:', fallback_id)

    def view_changed(self, desktop, window):
        if window is None:
            window_id = 'None'
        else:
            window_id = window.id

        print(
            'You need to override the '
            'VirtualDesktopNotificationCallback.view_changed method to '
            'not see this message.'
        )
        print('DESKTOP:', desktop.id)
        print('WINDOW:', window_id)


desktop_ids = Module.desktop_ids

pyWinVirtualDesktop = Module()
create_desktop = pyWinVirtualDesktop.create_desktop
register_notification_callback = pyWinVirtualDesktop.register_notification_callback
unregister_notification_callback = pyWinVirtualDesktop.unregister_notification_callback
