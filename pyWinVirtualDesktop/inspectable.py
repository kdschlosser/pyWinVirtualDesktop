# -*- coding: utf-8 -*-

from ctypes.wintypes import INT, ULONG
from ctypes import HRESULT, POINTER

import comtypes
import ctypes

from comtypes import helpstring, COMMETHOD
from comtypes.GUID import GUID


ENUM = INT
IID = GUID


# noinspection PyPep8Naming
class HSTRING__(ctypes.Structure):
    _fields_ = [
        ('unused', INT),
    ]


HSTRING = POINTER(HSTRING__)


class TrustLevel(ENUM):
    BaseTrust = 0
    PartialTrust = BaseTrust + 1
    FullTrust = PartialTrust + 1


IID_IInspectable = GUID(
    '{AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90}'
)


class IInspectable(comtypes.IUnknown):
    """
    Provides functionality required for all Windows Runtime classes.

    IInspectable methods have no effect on COM apartments and are safe to call
    from user interface threads.
    """
    _case_insensitive_ = True
    _idlflags_ = []
    _iid_ = IID_IInspectable
    _methods_ = [
        COMMETHOD(
            [helpstring(
                'Gets the interfaces that are implemented by '
                'the current Windows Runtime class.'
            )],
            HRESULT,
            'GetIids',
            (['out'], POINTER(ULONG), 'iidCount'),
            (['out'], POINTER(POINTER(IID)), 'iids'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the fully qualified name of the '
                'current Windows Runtime object.'
            )],
            HRESULT,
            'GetRuntimeClassName',
            (['out'], POINTER(HSTRING), 'className'),
        ),
        COMMETHOD(
            [helpstring(
                'Gets the trust level of the '
                'current Windows Runtime object.'
            )],
            HRESULT,
            'GetTrustLevel',
            (['out'], POINTER(TrustLevel), 'trustLevel'),
        ),
    ]
