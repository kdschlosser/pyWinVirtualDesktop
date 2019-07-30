# -*- coding: utf-8 -*-
from __future__ import print_function
import sys
import comtypes
import ctypes
import six
from ctypes import POINTER

from .winuser import EnumWindows, IsWindow, GetWindowText, GetProcessName
from .servprov import IServiceProvider, CLSID_ImmersiveShell
from .shobjidl_core import (
    CLSID_VirtualDesktopNotificationService,
    IID_IVirtualDesktopNotificationService,
    IVirtualDesktopNotificationService,

    CLSID_VirtualDesktopManagerInternal,
    IID_IVirtualDesktopManagerInternal,
    IVirtualDesktopManagerInternal,

    CLSID_VirtualDesktopPinnedApps,
    IID_IVirtualDesktopPinnedApps,
    IVirtualDesktopPinnedApps,

    CLSID_VirtualDesktopManager,
    IID_IVirtualDesktopManager,
    IVirtualDesktopManager,

    IID_IVirtualDesktop,
    IVirtualDesktop,
    IVirtualDesktopNotification,
    AdjacentDesktop
)

from .windows_ui_viewmanagement import (
    IID_IApplicationViewCollection,
    IApplicationViewCollection,
    IApplicationView9,
    IID_IApplicationView9,
    ApplicationViewOrientation,
    ApplicationViewBoundsMode,
    FullScreenSystemOverlayMode,
    ApplicationViewMode,
    ViewSizePreference
)


#
# ApplicationViewOrientation.ApplicationViewOrientation_Landscape
# ApplicationViewOrientation.ApplicationViewOrientation_Portrait
# ApplicationViewBoundsMode.ApplicationViewBoundsMode_UseVisible
# ApplicationViewBoundsMode.ApplicationViewBoundsMode_UseCoreWindow
# FullScreenSystemOverlayMode.FullScreenSystemOverlayMode_Standard
# FullScreenSystemOverlayMode.FullScreenSystemOverlayMode_Minimal
# ApplicationViewMode.ApplicationViewMode_Default
# ApplicationViewMode.ApplicationViewMode_CompactOverlay
# ViewSizePreference.ViewSizePreference_Default
# ViewSizePreference.ViewSizePreference_UseLess
# ViewSizePreference.ViewSizePreference_UseHalf
# ViewSizePreference.ViewSizePreference_UseMore
# ViewSizePreference.ViewSizePreference_UseMinimum
# ViewSizePreference.ViewSizePreference_UseNone
#
#
# bool = get_SuppressSystemOverlays()
# put_SuppressSystemOverlays(bool)
# Rect = get_VisibleBounds()
# cookie = add_VisibleBoundsChanged(IInspectable)
# remove_VisibleBoundsChanged(cookie)
# bool = SetDesiredBoundsMode(ApplicationViewBoundsMode)
# ApplicationViewBoundsMode = get_DesiredBoundsMode()
# IApplicationViewTitleBar = get_TitleBar()
# FullScreenSystemOverlayMode = get_FullScreenSystemOverlayMode()
# put_FullScreenSystemOverlayMode(FullScreenSystemOverlayMode)
# bool = get_IsFullScreenMode()
# bool = TryEnterFullScreenMode()
# ExitFullScreenMode() # if in fullscreen mode
# ShowStandardSystemOverlays() # if in fullscreen mode
# bool = TryResizeView(ctypes.byref(Size(width, height)))
# SetPreferredMinSize(ctypes.byref(Size(width, height)))
# ApplicationViewMode = get_ViewMode()
# bool = IsViewModeSupported(ApplicationViewMode)
# bool = TryEnterViewModeAsync(ApplicationViewMode)
# bool = TryEnterViewModeWithPreferencesAsync(ApplicationViewMode, IViewModePreferences)
# bool = TryConsolidateAsync() # closes the appview
# some_string = get_PersistedStateId()
# put_PersistedStateId(some_string)
# IWindowingEnvironment = get_WindowingEnvironment()
# IDisplayRegion = GetDisplayRegions()
#
# ApplicationViewOrientation = get_Orientation()
#
# bool = get_AdjacentToLeftDisplayEdge()
# bool = get_AdjacentToRightDisplayEdge()
# bool = get_IsFullScreen()
# bool = get_IsOnLockScreen()
# bool = get_IsScreenCaptureEnabled()
# put_IsScreenCaptureEnabled(bool)
# some_string = get_Title()
# put_Title(some_string)
#
# id = get_Id()
#
# cookie = add_Consolidated(IApplicationViewConsolidatedEventArgs)
# remove_Consolidated(cookie)
#

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
        self.__pNotificationService = comtypes.cast(
            self.__pServiceProvider.QueryService(
                CLSID_VirtualDesktopNotificationService,
                IID_IVirtualDesktopNotificationService
            ),
            ctypes.POINTER(IVirtualDesktopNotificationService)
        )
        self.__pPinnedApps = comtypes.cast(
            self.__pServiceProvider.QueryService(
                CLSID_VirtualDesktopPinnedApps,
                IID_IVirtualDesktopPinnedApps
            ),
            ctypes.POINTER(IVirtualDesktopPinnedApps)
        )

        self.__pDesktopManager = comtypes.CoCreateInstance(
            CLSID_VirtualDesktopManager,
            IVirtualDesktopManager,
        )

        self.__pViewCollection = comtypes.CoCreateInstance(
            IID_IApplicationViewCollection,
            IApplicationViewCollection,
        )

        # pObjectArray = self.__pViewCollection.GetViews()

        # for i in range(pObjectArray.GetCount()):
        #     ppView = comtypes.cast(
        #         pObjectArray.GetAt(i, IID_IApplicationView9),
        #         POINTER(IApplicationView9)
        #     )
        #
        # ppView = comtypes.cast(
        #     self.__pViewCollection.GetViewForHwnd(hwnd),
        #     POINTER(IApplicationView9)
        # )

        # self.__pViewCollection.RefreshCollection()

        self.DesktopNotificationCallback = (
            DesktopNotificationCallback
        )

        self.Desktop = Desktop
        self.Window = Window

    @property
    def desktop_ids(self):
        desktop_ids = []
        pObjectArray = self.__pDesktopManagerInternal.GetDesktops()

        for i in range(pObjectArray.GetCount()):
            pDesktop = ctypes.POINTER(IVirtualDesktop)()
            pObjectArray.GetAt(
                i,
                IID_IVirtualDesktop,
                ctypes.byref(pDesktop)
            )

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

        cls = VirtualDesktopNotification(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            callback
        )

        return self.__pNotificationService.Register(cls)

    def unregister_notification_callback(self, cookie):
        self.__pNotificationService.Unregister(cookie)

    def __iter__(self):
        for id in self.desktop_ids:
            yield Desktop(
                self.__pDesktopManagerInternal,
                self.__pDesktopManager,
                id
            )


class InstanceSingleton(type):

    def __init__(cls, *args, **kwargs):
        """
        InstanceSingleton metaclass constructor.
        """
        super(InstanceSingleton, cls).__init__(*args, **kwargs)

        cls._instances = {}

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


@six.add_metaclass(InstanceSingleton)
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


@six.add_metaclass(InstanceSingleton)
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

        return Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            neighbor.GetId()
        )

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


class VirtualDesktopNotification(comtypes.COMObject):
    _com_interfaces_ = [IVirtualDesktopNotification]

    def __init__(self, pDesktopManagerInternal, pDesktopManager, callback):
        self.__pDesktopManagerInternal = pDesktopManagerInternal
        self.__pDesktopManager = pDesktopManager
        self.__callback = callback
        comtypes.COMObject.__init__(self)

    def CurrentVirtualDesktopChanged(self, pDesktopOld, pDesktopNew):
        desktop_old = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopOld.GetId()
        )

        desktop_new = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopNew.GetId()
        )

        self.__callback.change(desktop_old, desktop_new)

        return S_OK

    def VirtualDesktopCreated(self, pDesktop):
        desktop = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktop.GetId()
        )

        self.__callback.create(desktop)

        return S_OK

    def VirtualDesktopDestroyBegin(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopDestroyed.GetId()
        )

        desktop_fallback = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopFallback.GetId()
        )

        self.__callback.destroy_begin(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def VirtualDesktopDestroyFailed(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopDestroyed.GetId()
        )

        desktop_fallback = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopFallback.GetId()
        )

        self.__callback.destroy_failed(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def VirtualDesktopDestroyed(self, pDesktopDestroyed, pDesktopFallback):
        desktop_destroyed = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopDestroyed.GetId()
        )

        desktop_fallback = Desktop(
            self.__pDesktopManagerInternal,
            self.__pDesktopManager,
            pDesktopFallback.GetId()
        )

        self.__callback.destroy(
            desktop_destroyed,
            desktop_fallback
        )

        return S_OK

    def ViewVirtualDesktopChanged(self, pView):
        return S_OK




desktop_ids = Module.desktop_ids
create_desktop = Module.create_desktop

Module()



