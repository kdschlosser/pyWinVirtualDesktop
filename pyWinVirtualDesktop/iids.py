# -*- coding: utf-8 -*-
from __future__ import print_function
from comtypes import GUID
import sys

x64 = sys.maxsize > 2**32

try:
    import _winreg
except ImportError:
    _winreg = __import__('winreg')

IID_IVirtualDesktop = GUID(
    '{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}'
)
IID_IVirtualDesktopNotification = GUID(
    '{C179334C-4295-40D3-BEA1-C654D965605A}'
)
IID_IVirtualDesktopNotificationService = GUID(
    '{0CD45E71-D927-4F15-8B0A-8FEF525337BF}'
)
IID_IVirtualDesktopManagerInternal = GUID(
    '{F31574D6-B682-4CDC-BD56-1827860ABEC6}'
)
IID_IVirtualDesktopManager = GUID(
    '{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}'
)
IID_IVirtualDesktopPinnedApps = GUID(
    '{4CE81583-1E4C-4632-A621-07A53543148F}'
)
IID_IApplicationView = GUID(
    '{372E1D3B-38D3-42E4-A15B-8AB2B178F513}'
)
IID_IApplicationView2 = GUID(
    '{E876B196-A545-40DC-B594-450CBA68CC00}'
)
IID_IApplicationView3 = GUID(
    '{903C9CE5-793A-4FDF-A2B2-AF1AC21E3108}'
)
IID_IApplicationView4 = GUID(
    '{15E5CBEC-9E0F-46B5-BC3F-9BF653E74B5E}'
)
IID_IApplicationView7 = GUID(
    '{A0369647-5FAF-5AA6-9C38-BEFBB12A071E}'
)
IID_IApplicationView9 = GUID(
    '{9C6516F9-021A-5F01-93E5-9BDAD2647574}'
)
IID_IApplicationViewTitleBar = GUID(
    '{00924AC0-932B-4A6B-9C4B-DC38C82478CE}'
)
IID_IViewModePreferences = GUID(
    '{878FCD3A-0B99-42C9-84D0-D3F1D403554B}'
)
IID_IApplicationViewCollection = GUID(
    '{1841C6D7-4F9D-42C0-AF41-8747538F10E5}'
)
IID_IApplicationViewConsolidatedEventArgs = GUID(
    '{514449EC-7EA2-4DE7-A6A6-7DFBAAEBB6FB}'
)
IID_IWindowingEnvironment = GUID(
    '{264363C0-2A49-5417-B3AE-48A71C63A3BD}'
)
IID_IDisplayRegion = GUID(
    '{DB50C3A2-4094-5F47-8CB1-EA01DDAFAA94}'
)
IID_IWindowingEnvironmentChangedEventArgs = GUID(
    '{4160CFC6-023D-5E9A-B431-350E67DC978A}'
)

IID_IApplicationViewCollectionManagement = None
IID_IApplicationViewChangeListener = None
IID_IImmersiveApplication = None

iids = [
    'IVirtualDesktopNotificationService',
    'IVirtualDesktopNotification',
    'IVirtualDesktopManagerInternal',
    'IVirtualDesktopManager',
    'IVirtualDesktopPinnedApps',
    'IVirtualDesktop',
    'IApplicationView2',
    'IApplicationView3',
    'IApplicationView4',
    'IApplicationView7',
    'IApplicationView9',
    'IApplicationViewCollection',
    'IApplicationViewConsolidatedEventArgs',
    'IApplicationViewTitleBar',
    'IApplicationViewCollectionManagement',
    'IApplicationViewChangeListener',
    'IApplicationView',
    'IViewModePreferences',
    'IDisplayRegion',
    'IWindowingEnvironmentChangedEventArgs',
    'IWindowingEnvironment',
    'IImmersiveApplication'
]


if x64:
    handle = _winreg.OpenKeyEx(
        _winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\\Classes\\Interface',
        0,
        _winreg.KEY_READ | _winreg.KEY_WOW64_64KEY
    )
else:
    handle = _winreg.OpenKeyEx(
        _winreg.HKEY_LOCAL_MACHINE,
        'SOFTWARE\\Classes\\Interface'
    )

res = {}

for i in range(_winreg.QueryInfoKey(handle)[0]):
    try:
        name = _winreg.EnumKey(handle, i)
        value = _winreg.QueryValue(handle, name)
    except WindowsError:
        continue

    for item in iids:
        if item in value:
            res['IID_' + item] = GUID(name.upper())

_winreg.CloseKey(handle)

mod = sys.modules[__name__]

if __name__ == '__main__':
    for key, value in mod.__dict__.items():
        if key.startswith('IID_'):
            if key in res:
                print('OLD:', key, ':', str(value))
                print('NEW:', key, ':', str(res[key]))
            else:
                print('NOT FOUND:', key)
else:
    mod.__dict__.update(res)
