# -*- coding: utf-8 -*-
from __future__ import print_function
import pyWinVirtualDesktop

import sys
appveyor_window = None

for desktop in pyWinVirtualDesktop:
    print('DESKTOP ID:', desktop.id)

    desktop_to_left = desktop.desktop_to_left
    desktop_to_right = desktop.desktop_to_right

    if desktop_to_left is None:
        print('DESKTOP TO LEFT: None')
    else:
        print('DESKTOP TO LEFT:', desktop_to_left.id)

    if desktop_to_right is None:
        print('DESKTOP TO RIGHT: None')
    else:
        print('DESKTOP TO RIGHT:', desktop_to_right.id)

    print('DESKTOP IS ACTIVE:', desktop.is_active)

    print('DESKTOP WINDOWS:')
    for window in desktop:
        if sys.version_info[0] == 2:
            if window.process_name == 'Appveyor.BuildAgent.Interactive.exe':
                appveyor_window = window
        else:
            if window.process_name == b'Appveyor.BuildAgent.Interactive.exe':
                appveyor_window = window

        print('    HANDLE:', window.id)
        print('    CAPTION:', repr(window.text))
        print('    PROCESS NAME:', repr(window.process_name))
        print('    ON ACTIVE DESKTOP:', window.is_on_active_desktop)
        print('\n')

if appveyor_window is not None:
    print('FOUND APPVEYOR')
    new_desktop = pyWinVirtualDesktop.create_desktop()
    print('NEW DESKTOP:', new_desktop.id)
    print('IS ACTIVE:', new_desktop.is_active)
    new_desktop.activate()
    print('IS ACTIVE:', new_desktop.is_active)
    print('DEKSTOP NUMBER:', new_desktop.number)
    print('WINDOWS:', list(w.id for w in new_desktop))
    new_desktop.add_window(appveyor_window)
    print('WINDOWS:', list(w.id for w in new_desktop))

    for desktop in pyWinVirtualDesktop:
        print('DESKTOP ID:', desktop.id)

        desktop_to_left = desktop.desktop_to_left
        desktop_to_right = desktop.desktop_to_right

        if desktop_to_left is None:
            print('DESKTOP TO LEFT: None')
        else:
            print('DESKTOP TO LEFT:', desktop_to_left.id)

        if desktop_to_right is None:
            print('DESKTOP TO RIGHT: None')
        else:
            print('DESKTOP TO RIGHT:', desktop_to_right.id)

    print('WINDOW ON ACTIVE:', appveyor_window.is_on_active_desktop)

    view = appveyor_window.view

    print('CAN RECEIVE INPUT:', view.can_receive_input)
    print('SCALE FACTOR:', view.scale_factor)
    print('SHOW IN SWITCHERS:', view.show_in_switchers)
    print('STATE:', view.state)
    print('SIZE:', view.size)
    print('POSITION:', view.position)
    print('IS VISIBLE:', view.is_visible)
    print('THUMBNAIL HANDLE:', view.thumbnail_handle)
    print('IS MIRRORED:', view.is_mirrored)
    print('SPLASH SCREEN:', view.is_splash_screen_presented)
    print('IS TRAY:', view.is_tray)
    print('HAS FOCUS:', view.has_focus)
    view.set_focus()
    print('HAS FOCUS:', view.has_focus)
    print('PINNED:', view.pinned)
    view.pinned = not view.pinned
    print('PINNED:', view.pinned)
    view.flash()
    view.activate()
