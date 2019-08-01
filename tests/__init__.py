# -*- coding: utf-8 -*-
from __future__ import print_function
import pyWinVirtualDesktop

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
        print('    HANDLE:', window.id)
        print('    CAPTION:', window.text)
        print('    PROCESS NAME:', window.process_name)
        print('    ON ACTIVE DESKTOP:', window.is_on_active_desktop)
        print('\n')
