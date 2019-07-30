# -*- coding: utf-8 -*-

from .inspectable import IInspectable, HSTRING
from .eventtoken import EventRegistrationToken
from .windows_ui import CColor
from .windows_foundation import Rect, Size
from .windows_ui_windowmanagement import IWindowingEnvironment, IDisplayRegion
from .objectarray import IObjectArray

from ctypes.wintypes import BOOL, INT, WCHAR, UINT, HWND
from ctypes import HRESULT, POINTER
import comtypes
import ctypes
from comtypes import helpstring, COMMETHOD
from comtypes.GUID import GUID


REFGUID = POINTER(GUID)
REFIID = REFGUID
ENUM = INT
IID = GUID
INT32 = ctypes.c_int32
INT64 = ctypes.c_int64
PCWSTR = POINTER(WCHAR)


IID_IApplicationView = GUID('{D222D519-4361-451E-96C4-60F4F9742DB0}')
IID_IApplicationView2 = GUID('{E876B196-A545-40DC-B594-450CBA68CC00}')
IID_IApplicationView3 = GUID('{903C9CE5-793A-4FDF-A2B2-AF1AC21E3108}')
IID_IApplicationView4 = GUID('{15E5CBEC-9E0F-46B5-BC3F-9BF653E74B5E}')
IID_IApplicationView7 = GUID('{A0369647-5FAF-5AA6-9C38-BEFBB12A071E}')
IID_IApplicationView9 = GUID('{9C6516F9-021A-5F01-93E5-9BDAD2647574}')

IID_IApplicationViewTitleBar = GUID('{00924AC0-932B-4A6B-9C4B-DC38C82478CE}')
IID_IViewModePreferences = GUID('{878FCD3A-0B99-42C9-84D0-D3F1D403554B}')

IID_IApplicationViewCollection = GUID(
    '{1841C6D7-4F9D-42C0-AF41-8747538F10E5}'
)
IID_IApplicationViewConsolidatedEventArgs = GUID(
    '{514449EC-7EA2-4DE7-A6A6-7DFBAAEBB6FB}'
)

IImmersiveApplication = UINT
IApplicationViewChangeListener = UINT


class ApplicationViewOrientation(ENUM):
    ApplicationViewOrientation_Landscape = 0
    ApplicationViewOrientation_Portrait = 1


class ApplicationViewBoundsMode(ENUM):
    ApplicationViewBoundsMode_UseVisible = 0
    ApplicationViewBoundsMode_UseCoreWindow = 1


class FullScreenSystemOverlayMode(ENUM):
    FullScreenSystemOverlayMode_Standard = 0
    FullScreenSystemOverlayMode_Minimal = 1


class ApplicationViewMode(ENUM):
    ApplicationViewMode_Default = 0
    ApplicationViewMode_CompactOverlay = 1


class ViewSizePreference(ENUM):
    ViewSizePreference_Default = 0
    ViewSizePreference_UseLess = 1
    ViewSizePreference_UseHalf = 2
    ViewSizePreference_UseMore = 3
    ViewSizePreference_UseMinimum = 4
    ViewSizePreference_UseNone = 5


class IApplicationViewConsolidatedEventArgs(IInspectable):
    _case_insensitive_ = True
    _iid_ = IID_IApplicationViewConsolidatedEventArgs
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Method get_IsUserInitiated')],
            HRESULT,
            'get_IsUserInitiated',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
    ]


class IApplicationViewTitleBar(IInspectable):
    """
    Represents the title bar of an app.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationViewTitleBar
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Sets the color of the title bar foreground.')],
            HRESULT,
            'put_ForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring('Gets the color of the title bar foreground.')],
            HRESULT,
            'get_ForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring('Sets the color of the title bar background.')],
            HRESULT,
            'put_BackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring('Gets the color of the title bar background.')],
            HRESULT,
            'get_BackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the foreground color of the title bar buttons.'
            )],
            HRESULT,
            'put_ButtonForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the foreground color of the title bar buttons.'
            )],
            HRESULT,
            'get_ButtonForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the background color of the title bar buttons.'
            )],
            HRESULT,
            'put_ButtonBackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the background color of the title bar buttons.'
            )],
            HRESULT,
            'get_ButtonBackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the foreground color of a title bar '
                'button when the pointer is over it.'
            )],
            HRESULT,
            'put_ButtonHoverForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the foreground color of a title '
                'bar button when the pointer is over it.'
            )],
            HRESULT,
            'get_ButtonHoverForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the background color of a title bar '
                'button when the pointer is over it.'
            )],
            HRESULT,
            'put_ButtonHoverBackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the background color of a title bar '
                'button when the pointer is over it.'
            )],
            HRESULT,
            'get_ButtonHoverBackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the foreground color of a title '
                'bar button when it\'s pressed.'
            )],
            HRESULT,
            'put_ButtonPressedForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the foreground color of a title '
                'bar button when it\'s pressed.'
            )],
            HRESULT,
            'get_ButtonPressedForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the background color of a title '
                'bar button when it\'s pressed.'
            )],
            HRESULT,
            'put_ButtonPressedBackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the background color of a title '
                'bar button when it\'s pressed.'
            )],
            HRESULT,
            'get_ButtonPressedBackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the color of the title bar '
                'foreground when it\'s inactive.'
            )],
            HRESULT,
            'put_InactiveForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the color of the title bar '
                'foreground when it\'s inactive.'
            )],
            HRESULT,
            'get_InactiveForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the color of the title bar '
                'background when it\'s inactive.'
            )],
            HRESULT,
            'put_InactiveBackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the color of the title bar '
                'background when it\'s inactive.'
            )],
            HRESULT,
            'get_InactiveBackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the foreground color of a title '
                'bar button when it\'s inactive.'
            )],
            HRESULT,
            'put_ButtonInactiveForegroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the foreground color of a title '
                'bar button when it\'s inactive.'
            )],
            HRESULT,
            'get_ButtonInactiveForegroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Sets the background color of a title '
                'bar button when it\'s inactive.'
            )],
            HRESULT,
            'put_ButtonInactiveBackgroundColor',
            (['in'], POINTER(CColor), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the background color of a title '
                'bar button when it\'s inactive.'
            )],
            HRESULT,
            'get_ButtonInactiveBackgroundColor',
            (['retval', 'out'], POINTER(POINTER(CColor)), 'value'),
        ),
    ]


class IViewModePreferences(IInspectable):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IViewModePreferences
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Gets the preferred size of the app window.'
            )],
            HRESULT,
            'get_ViewSizePreference',
            (['retval', 'out'], POINTER(ViewSizePreference), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the preferred size of the app window.'
            )],
            HRESULT,
            'put_ViewSizePreference',
            (['in'], ViewSizePreference, 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a custom preferred size for the app window.'
            )],
            HRESULT,
            'get_CustomSize',
            (['retval', 'out'], POINTER(Size), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets a custom preferred size for the app window.'
            )],
            HRESULT,
            'put_CustomSize',
            (['in'], Size, 'value'),
        ),
    ]


class IApplicationView(IInspectable):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Gets the current orientation (landscape or portrait) of '
                'the window (app view) with respect to the display.'
            )],
            HRESULT,
            'get_Orientation',
            (['retval', 'out'], POINTER(ApplicationViewOrientation), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value that indicates whether the current window '
                'is in close proximity to the left edge of the screen.'
            )],
            HRESULT,
            'get_AdjacentToLeftDisplayEdge',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value that indicates whether the current window is '
                'in close proximity to the right edge of the screen.'
            )],
            HRESULT,
            'get_AdjacentToRightDisplayEdge',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value that indicates whether the window '
                'touches both the left and right sides of the display.'
            )],
            HRESULT,
            'get_IsFullScreen',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets whether the window (app view) '
                'is on the Windows lock screen.'
            )],
            HRESULT,
            'get_IsOnLockScreen',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets whether screen capture is '
                'enabled for the window (app view).'
            )],
            HRESULT,
            'get_IsScreenCaptureEnabled',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets whether screen capture is '
                'enabled for the window (app view).'
            )],
            HRESULT,
            'put_IsScreenCaptureEnabled',
            (['in'], BOOL, 'value'),
        ),
        COMMETHOD(
            [helpstring('Sets the displayed title of the window.')],
            HRESULT,
            'put_Title',
            (['in'], HSTRING, 'value'),
        ),
        COMMETHOD(
            [helpstring('Gets the displayed title of the window.')],
            HRESULT,
            'get_Title',
            (['retval', 'out'], POINTER(HSTRING), 'value'),
        ),
        COMMETHOD(
            [helpstring('Gets the ID of the window (app view).')],
            HRESULT,
            'get_Id',
            (['retval', 'out'], POINTER(INT32), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Add event listener: occurs when the window is removed from '
                'the list of recently used apps, or if the user executes a '
                'close gesture on it.'
            )],
            HRESULT,
            'add_Consolidated',
            (
                ['in'],
                POINTER(IApplicationViewConsolidatedEventArgs),
                'handler'
            ),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring(
                'Remove event listener: occurs when the window is removed '
                'from the list of recently used apps, or if the user executes '
                'a close gesture on it.'
            )],
            HRESULT,
            'remove_Consolidated',
            (['in', ], EventRegistrationToken, 'token'),
        ),
    ]


class IApplicationView2(IApplicationView):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView2
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Gets a value indicating whether or not system overlays '
                '(such as overlay applications or the soft steering wheel) '
                'should be shown.'
            )],
            HRESULT,
            'get_SuppressSystemOverlays',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets a value indicating whether or not system overlays '
                '(such as overlay applications or the soft steering wheel) '
                'should be shown.'
            )],
            HRESULT,
            'put_SuppressSystemOverlays',
            (['in'], BOOL, 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Gets the visible region of the window (app view). The '
                'visible region is the region not occluded by chrome such as '
                'the status bar and app bar.'
            )],
            HRESULT,
            'get_VisibleBounds',
            (['retval', 'out'], POINTER(Rect), 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Add event listener: raised when the value of VisibleBounds '
                'changes, typically as a result of the status bar, app bar, '
                'or other chrome being shown or hidden.'
            )],
            HRESULT,
            'add_VisibleBoundsChanged',
            (['in'], POINTER(IInspectable), 'handler'),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring(
                'Remove event listener: raised when the value of '
                'VisibleBounds changes, typically as a result of the status '
                'bar, app bar, or other chrome being shown or hidden.'
            )],
            HRESULT,
            'remove_VisibleBoundsChanged',
            (['in'], EventRegistrationToken, 'token'),
        ),

        COMMETHOD(
            [helpstring(
                'Sets a value indicating the bounds used by the framework to '
                'lay out the contents of the window (app view).'
            )],
            HRESULT,
            'SetDesiredBoundsMode',
            (['in'], ApplicationViewBoundsMode, 'boundsMode'),
            (['retval', 'out'], POINTER(BOOL), 'success'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value indicating the bounds used by the framework to '
                'lay out the contents of the window (app view).'
            )],
            HRESULT,
            'get_DesiredBoundsMode',
            (['retval', 'out'], POINTER(ApplicationViewBoundsMode), 'value'),
        ),
    ]


class IApplicationView3(IApplicationView2):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView3
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Gets the title bar of the app.')],
            HRESULT,
            'get_TitleBar',
            (
                ['retval', 'out'],
                POINTER(POINTER(IApplicationViewTitleBar)),
                'value'
            ),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value that indicates how an app in full-screen '
                'mode responds to edge swipe actions.'
            )],
            HRESULT,
            'get_FullScreenSystemOverlayMode',
            (['retval', 'out'], POINTER(FullScreenSystemOverlayMode), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets a value that indicates how an app in '
                'full-screen mode responds to edge swipe actions.'
            )],
            HRESULT,
            'put_FullScreenSystemOverlayMode',
            (['in'], FullScreenSystemOverlayMode, 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets a value that indicates whether the '
                'app is running in full-screen mode.'
            )],
            HRESULT,
            'get_IsFullScreenMode',
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),
        COMMETHOD(
            [helpstring('Attempts to place the app in full-screen mode.')],
            HRESULT,
            'TryEnterFullScreenMode',
            (['retval', 'out'], POINTER(BOOL), 'success'),
        ),
        COMMETHOD(
            [helpstring('Takes the app out of full-screen mode.')],
            HRESULT,
            'ExitFullScreenMode',
            (),
        ),
        COMMETHOD(
            [helpstring(
                'Shows system UI elements, like the title '
                'bar, over a full-screen app.'
            )],
            HRESULT,
            'ShowStandardSystemOverlays',
            (),
        ),
        COMMETHOD(
            [helpstring(
                'Attempts to change the size of the view '
                'to the specified size in effective pixels.'
            )],
            HRESULT,
            'TryResizeView',
            (['in'], Size, 'value'),
            (['retval', 'out'], POINTER(BOOL), 'success'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets the smallest size, in effective '
                'pixels, allowed for the app window.'
            )],
            HRESULT,
            'SetPreferredMinSize',
            (['in'], Size, 'minSize'),
        ),
    ]


class IApplicationView4(IApplicationView3):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView4
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Gets the app view mode for the current view.')],
            HRESULT,
            'get_ViewMode',
            (['retval', 'out'], POINTER(ApplicationViewMode), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Determines whether the specified view '
                'mode is supported on the current device.'
            )],
            HRESULT,
            'IsViewModeSupported',
            (['in'], ApplicationViewMode, 'viewMode'),
            (['retval', 'out'], POINTER(BOOL), 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Attempts to change the app view to the specified view mode.'
            )],
            HRESULT,
            'TryEnterViewModeAsync',
            (['in'], ApplicationViewMode, 'viewMode'),
            (['retval', 'out'], POINTER(POINTER(BOOL)), 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Attempts to change the app view to the specified '
                'view mode using the specified options.'
            )],
            HRESULT,
            'TryEnterViewModeWithPreferencesAsync',
            (['in'], ApplicationViewMode, 'viewMode'),
            (['in'], POINTER(IViewModePreferences), 'viewModePreferences'),
            (['retval', 'out'], POINTER(POINTER(BOOL)), 'value'),
        ),

        COMMETHOD(
            [helpstring(
                'Tries to close the current app view. This method is a '
                'programmatic equivalent to a user initiating a close '
                'gesture for the app view.'
            )],
            HRESULT,
            'TryConsolidateAsync',
            (['retval', 'out'], POINTER(POINTER(BOOL)), 'value'),
        ),
    ]


class IApplicationView7(IApplicationView4):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView7
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Gets a string that identifies this '
                'view for tracking and saving state.'
            )],
            HRESULT,
            'get_PersistedStateId',
            (['retval', 'out'], POINTER(HSTRING), 'value'),
        ),
        COMMETHOD(
            [helpstring(
                'Sets a string that identifies this '
                'view for tracking and saving state.'
            )],
            HRESULT,
            'put_PersistedStateId',
            (['in'], HSTRING, 'value'),
        ),
    ]


class IApplicationView9(IApplicationView7):
    """
    Represents the active application view and associated states and behaviors.
    """
    _case_insensitive_ = True
    _iid_ = IID_IApplicationView9
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('Gets the windowing environment for the view.')],
            HRESULT,
            'get_WindowingEnvironment',
            (
                ['retval', 'out'],
                POINTER(POINTER(IWindowingEnvironment)),
                'value'
            ),
        ),
        COMMETHOD(
            [helpstring(
                'Returns the collection of display '
                'regions available for the view.'
            )],
            HRESULT,
            'GetDisplayRegions',
            (['retval', 'out'], POINTER(POINTER(IDisplayRegion)), 'result'),
        ),
    ]


class IApplicationViewCollection(comtypes.IUnknown):
    _case_insensitive_ = True
    _iid_ = IID_IApplicationViewCollection
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [helpstring('GetViews')],
            HRESULT,
            'GetViews',
            (['out', 'retval'], POINTER(POINTER(IObjectArray)), 'ppViews'),
        ),
        COMMETHOD(
            [helpstring('GetViewsByZOrder')],
            HRESULT,
            'GetViewsByZOrder',
            (['out', 'retval'], POINTER(POINTER(IObjectArray)), 'ppViews'),
        ),
        COMMETHOD(
            [helpstring('GetViewsByAppUserModelId')],
            HRESULT,
            'GetViewsByAppUserModelId',
            (['in'], PCWSTR, 'pUserModeId'),
            (['out', 'retval'], POINTER(POINTER(IObjectArray)), 'ppViews'),
        ),
        COMMETHOD(
            [helpstring('GetViewForHwnd')],
            HRESULT,
            'GetViewForHwnd',
            (['in'], HWND, 'hwnd'),
            (['out', 'retval'], POINTER(POINTER(IApplicationView)), 'ppView'),
        ),
        COMMETHOD(
            [helpstring('GetViewForApplication')],
            HRESULT,
            'GetViewForApplication',
            (['in'], POINTER(IImmersiveApplication), 'pApplication'),
            (['out', 'retval'], POINTER(POINTER(IApplicationView)), 'ppView'),
        ),
        COMMETHOD(
            [helpstring('GetViewForAppUserModelId')],
            HRESULT,
            'GetViewForAppUserModelId',
            (['in'], PCWSTR, 'pUserModeId'),
            (['out', 'retval'], POINTER(POINTER(IApplicationView)), 'ppView'),
        ),
        COMMETHOD(
            [helpstring('GetViewInFocus')],
            HRESULT,
            'GetViewInFocus',
            (['out', 'retval'], POINTER(POINTER(IApplicationView)), 'ppView'),
        ),
        COMMETHOD(
            [helpstring('Unknown1')],
            HRESULT,
            'Unknown1',
            (['out', 'retval'], POINTER(POINTER(IApplicationView)), 'ppView'),
        ),
        COMMETHOD(
            [helpstring('RefreshCollection')],
            HRESULT,
            'RefreshCollection',
            (),
        ),
        COMMETHOD(
            [helpstring('RegisterForApplicationViewChanges')],
            HRESULT,
            'RegisterForApplicationViewChanges',
            (['in'], POINTER(IApplicationViewChangeListener), 'pListener'),
            (['retval', 'out'], POINTER(EventRegistrationToken), 'token'),
        ),
        COMMETHOD(
            [helpstring('UnregisterForApplicationViewChanges')],
            HRESULT,
            'UnregisterForApplicationViewChanges',
            (['in'], EventRegistrationToken, 'token'),
        ),
    ]
