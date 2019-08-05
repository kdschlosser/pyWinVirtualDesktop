# -*- coding: utf-8 -*-
from __future__ import print_function
import time

print('Importing pyWinVirtualDesktop')
time.sleep(2.0)

import pyWinVirtualDesktop

print('Import done!')
time.sleep(2.0)

import sys
appveyor_window = None

for desktop in pyWinVirtualDesktop:
    print('DESKTOP ID:', desktop.id)
    time.sleep(0.1)

    desktop_to_left = desktop.desktop_to_left
    desktop_to_right = desktop.desktop_to_right

    if desktop_to_left is None:
        print('DESKTOP TO LEFT: None')
    else:
        print('DESKTOP TO LEFT:', desktop_to_left.id)
    time.sleep(0.1)
    if desktop_to_right is None:
        print('DESKTOP TO RIGHT: None')
    else:
        print('DESKTOP TO RIGHT:', desktop_to_right.id)
    time.sleep(0.1)
    print('DESKTOP IS ACTIVE:', desktop.is_active)
    time.sleep(0.1)
    print('DESKTOP WINDOWS:')
    time.sleep(0.1)
    for window in desktop:
        if sys.version_info[0] == 2:
            if window.process_name == 'Appveyor.BuildAgent.Interactive.exe':
                appveyor_window = window
        else:
            if window.process_name == b'Appveyor.BuildAgent.Interactive.exe':
                appveyor_window = window

        print('    HANDLE:', window.id)
        time.sleep(0.1)
        print('    CAPTION:', repr(window.text))
        time.sleep(0.1)
        print('    PROCESS NAME:', repr(window.process_name))
        time.sleep(0.1)
        print('    ON ACTIVE DESKTOP:', window.is_on_active_desktop)
        time.sleep(0.1)
        print('\n')
        time.sleep(0.1)

if appveyor_window is not None:
    print('FOUND APPVEYOR')
    time.sleep(0.1)
    new_desktop = pyWinVirtualDesktop.create_desktop()
    print('NEW DESKTOP:', new_desktop.id)
    time.sleep(0.1)
    print('IS ACTIVE:', new_desktop.is_active)
    time.sleep(0.1)
    new_desktop.activate()
    print('IS ACTIVE:', new_desktop.is_active)
    time.sleep(0.1)
    print('DEKSTOP NUMBER:', new_desktop.number)
    time.sleep(0.1)
    print('WINDOWS:', list(w.id for w in new_desktop))
    time.sleep(0.1)
    new_desktop.add_window(appveyor_window)
    print('WINDOWS:', list(w.id for w in new_desktop))
    time.sleep(0.1)

    for desktop in pyWinVirtualDesktop:
        print('DESKTOP ID:', desktop.id)
        time.sleep(0.1)

        desktop_to_left = desktop.desktop_to_left
        desktop_to_right = desktop.desktop_to_right

        if desktop_to_left is None:
            print('DESKTOP TO LEFT: None')
        else:
            print('DESKTOP TO LEFT:', desktop_to_left.id)
        time.sleep(0.1)

        if desktop_to_right is None:
            print('DESKTOP TO RIGHT: None')
        else:
            print('DESKTOP TO RIGHT:', desktop_to_right.id)

        time.sleep(0.1)

    print('WINDOW ON ACTIVE:', appveyor_window.is_on_active_desktop)
    time.sleep(0.1)

    view = appveyor_window.view

    print('CAN RECEIVE INPUT:', view.can_receive_input)
    time.sleep(0.1)
    print('SCALE FACTOR:', view.scale_factor)
    time.sleep(0.1)
    print('SHOW IN SWITCHERS:', view.show_in_switchers)
    time.sleep(0.1)
    print('STATE:', view.state)
    time.sleep(0.1)
    print('SIZE:', view.size)
    time.sleep(0.1)
    print('POSITION:', view.position)
    time.sleep(0.1)
    print('IS VISIBLE:', view.is_visible)
    time.sleep(0.1)
    print('THUMBNAIL HANDLE:', view.thumbnail_handle)
    time.sleep(0.1)
    print('IS MIRRORED:', view.is_mirrored)
    time.sleep(0.1)
    print('SPLASH SCREEN:', view.is_splash_screen_presented)
    time.sleep(0.1)
    print('IS TRAY:', view.is_tray)
    time.sleep(0.1)
    print('HAS FOCUS:', view.has_focus)
    time.sleep(0.1)
    view.set_focus()
    print('HAS FOCUS:', view.has_focus)
    time.sleep(0.1)
    print('PINNED:', view.pinned)
    time.sleep(0.1)
    view.pinned = not view.pinned
    print('PINNED:', view.pinned)
    time.sleep(0.1)
    view.flash()
    view.activate()
