# pyWinVirtualDesktop
Windows 10 Virtual Desktop Management

Requirements

* comtypes
* six

This is still a work in progress. But the basic use is as follows.

```python
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
```

Here is the list of available properties/methods

pyWinVirtualDesktop


* Properties

    * desktop_ids: a list of all of the desktop ids

* Functions

    * create_desktop: creates a new desktop and returns a Desktop instance

Desktop

* Properties

    * id: returns comtypes.GUID instance
    * desktop_to_left: return Desktop instance
    * desktop_to_right: return Desktop instance
    * is_active: return `True`/`False`

* Methods

    * add_window(window): adds a window to the desktop, window can be wither a window handle or a Window class instance
    * add_view:(view): under construction
    * delete():: deletes the window


Window

* Properties

    * id: window handle
    * text: window caption
    * is_ok: if the window is still valid
    * desktop: Desktop instance the window is on
    * process_name: the executable name
    * is_on_active_desktop: if the windows is on the desktop that is being viewed

* Methods

    * move_to_desktop(desktop): move the window to a new desktop. desktop can be either the desktop id or a Desktop instance
