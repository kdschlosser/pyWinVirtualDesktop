# -*- coding: utf-8 -*-


from ctypes.wintypes import INT, UINT, LPVOID
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


IID_IObjectArray = GUID(
    "{92CA9DCD-5622-4BBA-A805-5E9F541BD8C9}"
)


class IObjectArray(comtypes.IUnknown):
    """
    Exposes methods that enable clients to access items in a collection of
    objects that support IUnknown.
    """
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IObjectArray

    _methods_ = [
        COMMETHOD(
            [helpstring('Provides a count of the objects in the collection.')],
            HRESULT,
            'GetCount',
            (['out'], POINTER(UINT), 'pcObjects'),
        ),
        COMMETHOD(
            [helpstring(
                'Provides a pointer to a specified object\'s interface. '
                'The object and interface are specified by index and '
                'interface ID.'
            )],
            HRESULT,
            'GetAt',
            (['in'], UINT, 'uiIndex'),
            (['in'], REFIID, 'riid'),
            (['out', 'iid_is(riid)'], POINTER(LPVOID), 'ppv'),
        ),
    ]
