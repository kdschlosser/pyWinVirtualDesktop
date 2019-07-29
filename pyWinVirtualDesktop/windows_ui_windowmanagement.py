# -*- coding: utf-8 -*-


from .inspectable import IInspectable, HSTRING
from .eventtoken import EventRegistrationToken
from .windows_foundation import Point, Size

from ctypes import HRESULT, POINTER
from ctypes.wintypes import BOOL, INT
from comtypes import GUID, COMMETHOD, helpstring


ENUM = INT

IID_IWindowingEnvironment = GUID(
    '{264363C0-2A49-5417-B3AE-48A71C63A3BD}'
)
IID_IDisplayRegion = GUID(
    '{DB50C3A2-4094-5F47-8CB1-EA01DDAFAA94}'
)
IID_IWindowingEnvironmentChangedEventArgs = GUID(
    '{4160CFC6-023D-5E9A-B431-350E67DC978A}'
)


class WindowingEnvironmentKind(ENUM):
    WindowingEnvironmentKind = 0


class IDisplayRegion(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IDisplayRegion
    _idlflags_ = []


class IWindowingEnvironmentChangedEventArgs(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IWindowingEnvironmentChangedEventArgs
    _idlflags_ = []


class IWindowingEnvironment(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IWindowingEnvironment
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('get_IsEnabled')],
            HRESULT,
            'get_IsEnabled',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('get_Kind')],
            HRESULT,
            'get_Kind',
            (['retval', 'out'], POINTER(WindowingEnvironmentKind), 'value'),
        ),
        COMMETHOD(
            [helpstring('GetDisplayRegions')],
            HRESULT,
            'GetDisplayRegions',
            (['retval', 'out'], POINTER(POINTER(IDisplayRegion)), 'result'),
        ),
        COMMETHOD(
            [helpstring('add_Changed')],
            HRESULT,
            'add_Changed',
            (
                ['in'],
                POINTER(IWindowingEnvironmentChangedEventArgs),
                'handler'
            ),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring('remove_Changed')],
            HRESULT,
            'remove_Changed',
            (['in'], EventRegistrationToken, 'token'),
        ),
    ]


IDisplayRegion._methods_ = [
        COMMETHOD(
            [helpstring('get_DisplayMonitorDeviceId')],
            HRESULT,
            'get_DisplayMonitorDeviceId',
            (['retval', 'out'], POINTER(HSTRING), 'value'),
        ),
        COMMETHOD(
            [helpstring('get_IsVisible')],
            HRESULT,
            'get_IsVisible',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('get_WorkAreaOffset')],
            HRESULT,
            'get_WorkAreaOffset',
            (['retval', 'out'], POINTER(Point), 'value'),
        ),
        COMMETHOD(
            [helpstring('get_WorkAreaSize')],
            HRESULT,
            'get_WorkAreaSize',
            (['retval', 'out'], POINTER(Size), 'value'),
        ),
        COMMETHOD(
            [helpstring('get_WindowingEnvironment')],
            HRESULT,
            'get_WindowingEnvironment',
            (['retval', 'out'], POINTER(IWindowingEnvironment), 'value'),
        ),
        COMMETHOD(
            [helpstring('add_Changed')],
            HRESULT,
            'add_Changed',
            (['in'], POINTER(IDisplayRegion), 'handler'),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring('remove_Changed')],
            HRESULT,
            'remove_Changed',
            (['in'], EventRegistrationToken, 'token'),
        ),
    ]
