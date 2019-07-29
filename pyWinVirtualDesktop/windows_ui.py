# -*- coding: utf-8 -*-

import ctypes

from ctypes.wintypes import BYTE


class Color(ctypes.Structure):
    _fields_ = [
        ('A', BYTE),
        ('R', BYTE),
        ('G', BYTE),
        ('B', BYTE)
    ]


CColor = Color
