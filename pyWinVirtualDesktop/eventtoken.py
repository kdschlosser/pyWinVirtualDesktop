# -*- coding: utf-8 -*-

import ctypes

INT64 = ctypes.c_int64


class EventRegistrationToken(ctypes.Structure):
    _fields_ = [
        ('value', INT64)
    ]
