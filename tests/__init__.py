# -*- coding: utf-8 -*-
from __future__ import print_function
import pyWinVirtualDesktop


for desktop in pyWinVirtualDesktop:
    print('DESKTOP ID:', desktop.id)
    print('DESKTOP TO LEFT:', desktop.desktop_to_left.id)
    print('DESKTOP TO RIGHT:', desktop.desktop_to_right.id)
    print('DESKTOP IS ACTIVE:', desktop.is_active)

    print('DESKTOP WINDOWS:')
    for window in desktop:
        print('    HANDLE:', window.id)
        print('    CAPTION:', window.text)
        print('    PROCESS NAME:', window.process_name)
        print('    ON ACTIVE DESKTOP:', window.is_on_active_desktop)
        print('\n')
