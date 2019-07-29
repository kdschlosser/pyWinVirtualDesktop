# -*- coding: utf-8 -*-

import ctypes
from ctypes.wintypes import FLOAT


class Rect(ctypes.Structure):
    _fields_ = [
        ('X', FLOAT),
        ('Y', FLOAT),
        ('Width', FLOAT),
        ('Height', FLOAT)
    ]


class Size(ctypes.Structure):
    _fields_ = [
        ('Width', FLOAT),
        ('Height', FLOAT)
    ]


class Point(ctypes.Structure):
    _fields_ = [
        ('X', FLOAT),
        ('Y', FLOAT),
    ]
