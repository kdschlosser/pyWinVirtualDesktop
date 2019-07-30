# -*- coding: utf-8 -*-


from ctypes.wintypes import INT, LPVOID
from ctypes import HRESULT, POINTER

import comtypes
import ctypes
from comtypes import helpstring, COMMETHOD
from comtypes.GUID import GUID


REFGUID = POINTER(GUID)
REFIID = REFGUID
ENUM = INT
IID = GUID
INT32 = ctypes.c_int32
INT64 = ctypes.c_int64


IID_IServiceProvider = GUID(
    '{6D5140C1-7436-11CE-8034-00AA006009FA}'
)

CLSID_ImmersiveShell = GUID(
    '{C2F03A33-21F5-47FA-B4BB-156362A2F239}'
)


class IServiceProvider(comtypes.IUnknown):
    """
    Defines a mechanism for retrieving a service object; that is,
    an object that provides custom support to other objects.
    """
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IServiceProvider
    _methods_ = [
        COMMETHOD(
            [helpstring('Gets the service object of the specified type.')],
            HRESULT,
            'QueryService',
            (['in'], REFGUID, 'guidService'),
            (['in'], REFIID, 'riid'),
            (['out'], POINTER(LPVOID), 'ppvObject'),
        ),
        COMMETHOD(
            [helpstring('Gets the service object of the specified type.')],
            HRESULT,
            'RemoteQueryService',
            (['in'], REFGUID, 'guidService'),
            (['out'], POINTER(POINTER(comtypes.IUnknown)), 'ppvObject'),
        ),
    ]
