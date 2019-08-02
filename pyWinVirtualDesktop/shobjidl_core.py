# -*- coding: utf-8 -*-

from ctypes.wintypes import HWND, BOOL, DWORD, INT, UINT, WCHAR, LPVOID
from ctypes import HRESULT, POINTER

import comtypes
import ctypes

from comtypes import helpstring, COMMETHOD
from comtypes.GUID import GUID

from .windows_ui_viewmanagement import IApplicationView
from .objectarray import IObjectArray
from .iids import (
    IID_IVirtualDesktop,
    IID_IVirtualDesktopNotification,
    IID_IVirtualDesktopNotificationService,
    IID_IVirtualDesktopManagerInternal,
    IID_IVirtualDesktopManager,
    IID_IVirtualDesktopPinnedApps
)

REFGUID = POINTER(GUID)
REFIID = REFGUID
ENUM = INT
IID = GUID
INT32 = ctypes.c_int32
INT64 = ctypes.c_int64
PCWSTR = POINTER(WCHAR)


CLSID_VirtualDesktopNotificationService = GUID(
    '{A501FDEC-4A09-464C-AE4E-1B9C21B84918}'
)
CLSID_VirtualDesktopManagerInternal = GUID(
    '{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}'
)
CLSID_VirtualDesktopManager = GUID(
    '{AA509086-5CA9-4C25-8f95-589D3C07B48A}'
)
CLSID_VirtualDesktopPinnedApps = GUID(
    '{B5A399E7-1C87-46B8-88E9-FC5747B171BD}'
)

class AdjacentDesktop(ENUM):
    LeftDirection = 3
    RightDirection = 4


class IVirtualDesktopPinnedApps(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopPinnedApps
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('IsAppIdPinned')],
            HRESULT,
            'IsAppIdPinned',
            (['in'], PCWSTR, 'pAppId'),
            (['out', 'retval'], POINTER(BOOL), 'pPinned'),
        ),
        COMMETHOD(
            [helpstring('PinAppID')],
            HRESULT,
            'PinAppID',
            (['in'], PCWSTR, 'pAppId'),
        ),
        COMMETHOD(
            [helpstring('UnpinAppID')],
            HRESULT,
            'UnpinAppID',
            (['in'], PCWSTR, 'pAppId'),
        ),
        COMMETHOD(
            [helpstring('IsViewPinned')],
            HRESULT,
            'IsViewPinned',
            (['in'], POINTER(IApplicationView), 'pView'),
            (['out', 'retval'], POINTER(BOOL), 'pPinned'),
        ),
        COMMETHOD(
            [helpstring('PinView')],
            HRESULT,
            'PinView',
            (['in'], POINTER(IApplicationView), 'pView'),
        ),
        COMMETHOD(
            [helpstring('UnpinView')],
            HRESULT,
            'UnpinView',
            (['in'], POINTER(IApplicationView), 'pView'),
        ),
    ]


class IVirtualDesktop(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktop
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Is the desktop visible')],
            HRESULT,
            'IsViewVisible',
            (['in'], POINTER(IApplicationView), 'pView'),
            (['out', 'retval'], POINTER(INT), 'pfVisible'),
        ),
        COMMETHOD(
            [helpstring('Gets the id of the desktop')],
            HRESULT,
            'GetID',
            (['out', 'retval'], POINTER(GUID), 'pGuid'),
        )
    ]


class IVirtualDesktopManager(comtypes.IUnknown):
    """
    Exposes methods that enable an application to interact with groups of
    windows that form virtual workspaces.
    """
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopManager
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Indicates whether the provided window is '
                'on the currently active virtual desktop.'
            )],
            HRESULT,
            'IsWindowOnCurrentVirtualDesktop',
            (['in'], HWND, 'topLevelWindow'),
            (['out', 'retval'], POINTER(BOOL), 'onCurrentDesktop'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the identifier for the virtual desktop hosting '
                'the provided top-level window.'
            )],
            HRESULT,
            'GetWindowDesktopId',
            (['in'], HWND, 'topLevelWindow'),
            (['out', 'retval'], POINTER(GUID), 'desktopId'),
        ),
        COMMETHOD(
            [helpstring('Moves a window to the specified virtual desktop.')],
            HRESULT,
            'MoveWindowToDesktop',
            (['in'], HWND, 'topLevelWindow'),
            (['in'], REFGUID, 'desktopId'),
        ),
    ]


class IVirtualDesktopManagerInternal(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopManagerInternal
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method GetCount')],
            HRESULT,
            'GetCount',
            (['out', 'retval'], POINTER(UINT), 'pCount'),
        ),
        COMMETHOD(
            [helpstring('Method MoveViewToDesktop')],
            HRESULT,
            'MoveViewToDesktop',
            (['in'], POINTER(IApplicationView), 'pView'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method CanViewMoveDesktops')],
            HRESULT,
            'CanViewMoveDesktops',
            (['in'], POINTER(IApplicationView), 'pView'),
            (['out', 'retval'], POINTER(BOOL), 'pfCanViewMoveDesktops'),
        ),
        COMMETHOD(
            [helpstring('Method GetCurrentDesktop')],
            HRESULT,
            'GetCurrentDesktop',
            (['out', 'retval'], POINTER(POINTER(IVirtualDesktop)), 'desktop'),
        ),
        COMMETHOD(
            [helpstring('Method GetDesktops')],
            HRESULT,
            'GetDesktops',
            (['out', 'retval'], POINTER(POINTER(IObjectArray)), 'ppDesktops'),
        ),
        COMMETHOD(
            [helpstring('Method GetAdjacentDesktop')],
            HRESULT,
            'GetAdjacentDesktop',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopReference'),
            (['in'], AdjacentDesktop, 'uDirection'),
            (['out', 'retval'], POINTER(POINTER(IVirtualDesktop)), 'ppAdjacentDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method SwitchDesktop')],
            HRESULT,
            'SwitchDesktop',
            (['in'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method CreateDesktopW')],
            HRESULT,
            'CreateDesktopW',
            (
                ['out', 'retval'],
                POINTER(POINTER(IVirtualDesktop)),
                'ppNewDesktop'
            ),
        ),
        COMMETHOD(
            [helpstring('Method RemoveDesktop')],
            HRESULT,
            'RemoveDesktop',
            (['in'], POINTER(IVirtualDesktop), 'pRemove'),
            (['in'], POINTER(IVirtualDesktop), 'pFallbackDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method FindDesktop')],
            HRESULT,
            'FindDesktop',
            (['in'], POINTER(GUID), 'desktopId'),
            (
                ['out', 'retval'],
                POINTER(POINTER(IVirtualDesktop)),
                'ppDesktop'
            ),
        ),
    ]


class IVirtualDesktopNotification(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopNotification
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method VirtualDesktopCreated')],
            HRESULT,
            'VirtualDesktopCreated',
            (['in'], POINTER(IVirtualDesktop), 'pDesktop'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyBegin')],
            HRESULT,
            'VirtualDesktopDestroyBegin',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyFailed')],
            HRESULT,
            'VirtualDesktopDestroyFailed',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method VirtualDesktopDestroyed')],
            HRESULT,
            'VirtualDesktopDestroyed',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopDestroyed'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopFallback'),
        ),
        COMMETHOD(
            [helpstring('Method ViewVirtualDesktopChanged')],
            HRESULT,
            'ViewVirtualDesktopChanged',
            (['in'], POINTER(IApplicationView), 'pView'),
        ),
        COMMETHOD(
            [helpstring('Method CurrentVirtualDesktopChanged')],
            HRESULT,
            'CurrentVirtualDesktopChanged',
            (['in'], POINTER(IVirtualDesktop), 'pDesktopOld'),
            (['in'], POINTER(IVirtualDesktop), 'pDesktopNew'),
        ),
    ]


class IVirtualDesktopNotificationService(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IVirtualDesktopNotificationService
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method Register')],
            HRESULT,
            'Register',
            (['in'], POINTER(IVirtualDesktopNotification), 'pNotification'),
            (['out'], POINTER(DWORD), 'pdwCookie'),
        ),

        COMMETHOD(
            [helpstring('Method Unregister')],
            HRESULT,
            'Unregister',
            (['in'], DWORD, 'dwCookie'),
        ),
    ]

