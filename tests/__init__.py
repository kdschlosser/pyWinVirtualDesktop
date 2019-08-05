# -*- coding: utf-8 -*-
from __future__ import print_function
import sys
import time

import pyWinVirtualDesktop

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

    new_desktop.add_window(appveyor_window)

    for desktop in pyWinVirtualDesktop:
        if desktop != new_desktop:
            desktop.activate()
        print('DESKTOP ID:', desktop.id)

        print('IS ACTIVE:', desktop.is_active)
        print('DESKTOP NUMBER:', desktop.number)
        print('DESKTOP NAME:', desktop.name)

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

        for window in desktop:
            print('    HANDLE:', window.id)
            print('    CAPTION:', repr(window.text))
            print('    PROCESS NAME:', repr(window.process_name))
            print('    ON ACTIVE DESKTOP:', window.is_on_active_desktop)
            print('\n')

    # view = appveyor_window.view
    #
    # print('CAN RECEIVE INPUT:', view.can_receive_input)
    # sys.stdout.flush()
    #
    # print('SCALE FACTOR:', view.scale_factor)
    # sys.stdout.flush()
    #
    # print('SHOW IN SWITCHERS:', view.show_in_switchers)
    # sys.stdout.flush()
    #
    # print('STATE:', view.state)
    # sys.stdout.flush()
    #
    # print('SIZE:', view.size)
    # sys.stdout.flush()
    #
    # print('POSITION:', view.position)
    # sys.stdout.flush()
    #
    # print('IS VISIBLE:', view.is_visible)
    # sys.stdout.flush()
    #
    # print('THUMBNAIL HANDLE:', view.thumbnail_handle)
    # sys.stdout.flush()
    #
    # print('IS MIRRORED:', view.is_mirrored)
    # sys.stdout.flush()
    #
    # print('SPLASH SCREEN:', view.is_splash_screen_presented)
    # sys.stdout.flush()
    #
    # print('IS TRAY:', view.is_tray)
    # sys.stdout.flush()
    #
    # print('HAS FOCUS:', view.has_focus)
    # sys.stdout.flush()
    #
    # view.set_focus()
    # print('HAS FOCUS:', view.has_focus)
    # sys.stdout.flush()
    #
    # print('PINNED:', view.pinned)
    # sys.stdout.flush()
    #
    # view.pinned = not view.pinned
    # print('PINNED:', view.pinned)
    # sys.stdout.flush()
    #
    # view.flash()
    # view.activate()
