# pyWinVirtualDesktop
Windows 10 Virtual Desktop Management

## *Just Added*

* callback notifications

## Requirements

* comtypes
* six

<br></br>

## Installation

### *Install*

```
python setup.py install
```

### *Build*

```
python setup.py build
```

<br></br>

## Basic Use

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

<br></br>

## Properties/Methods

### *pyWinVirtualDesktop*


* Properties

    * desktop_ids: a list of all of the desktop ids

* Functions

    * create_desktop: creates a new desktop and returns a Desktop instance

### *pyWinVirtualDesktop.Desktop*

* Properties

    * id: returns comtypes.GUID instance
    * desktop_to_left: return Desktop instance
    * desktop_to_right: return Desktop instance
    * is_active: return `True`/`False`

* Methods

    * add_window(window): adds a window to the desktop, window can be wither a window handle or a Window class instance
    * add_view:(view): under construction
    * delete():: deletes the window


### *pyWinVirtualDesktop.Window*

* Properties

    * id: window handle
    * text: window caption
    * is_ok: if the window is still valid
    * desktop: Desktop instance the window is on
    * process_name: the executable name
    * is_on_active_desktop: if the windows is on the desktop that is being viewed

* Methods

    * move_to_desktop(desktop): move the window to a new desktop. desktop can be either the desktop id or a Desktop instance


<br></br>

## Callback Notifications

In order to register for notifications you need to either create a
subclass or an instance of `pyWinVirtualDesktop.DesktopNotificationCallback`
and override several methods.

* `create(new)`
* `change(old, new)`
* `destroy_begin(destroyed, fallback)`
* `destroy_failed(destroyed, fallback)`
* `destroy(destroyed, fallback)`

There are 2 ways this can be done.
It's easier to explain through the use of code.

### *Way 1*

```python
import pyWinVirtualDesktop

class NotificationCallback(pyWinVirtualDesktop.DesktopNotificationCallback):
    def create(self, new):
        pass

    def change(self, old, new):
        pass

    def destroy_begin(self, destroyed, fallback):
        pass

    def destroy_failed(self, destroyed, fallback):
        pass

    def destroy(self, destroyed, fallback):
        pass

cookie = pyWinVirtualDesktop.register_notification_callback(NotificationCallback)

```

you can also pass an instance of the subclass if you like.
```python
callback = NotificationCallback()
cookie = pyWinVirtualDesktop.register_notification_callback(callback)
```

### *Way 2*

```Python
import pyWinVirtualDesktop

def create(new):
    pass

def change(old, new):
    pass

def destroy_begin(destroyed, fallback):
    pass

def destroy_failed(destroyed, fallback):
    pass

def destroy(destroyed, fallback):
    pass

callback = pyWinVirtualDesktop.DesktopNotificationCallback()

callback.create = create
callback.change = change
callback.destroy_begin = destroy_begin
callback.destroy_failed = destroy_failed
callback.destroy = destroy

cookie = pyWinVirtualDesktop.register_notification_callback(callback)

```

Either way will work. But make sure you hold a reference to the cookie
that is returned from calling  `register_notification_callback`

This cookie is what you use to unregister

```python
pyWinVirtualDesktop.unregister_notification_callback(cookie)

```

in the methods you override you will place whatever code it i you want
to run. The parameters that are passed are `Desktop` instances.
