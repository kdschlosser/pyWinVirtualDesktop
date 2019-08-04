#pragma once
#include "stdafx.h"
#include "VirtualDesktop.h"
#include <Rpc.h>
#include <Python.h>
#include <string.h>

#define string std::string

#define VDA_VirtualDesktopCreated 5
#define VDA_VirtualDesktopDestroyBegin 4
#define VDA_VirtualDesktopDestroyFailed 3
#define VDA_VirtualDesktopDestroyed 2
#define VDA_ViewVirtualDesktopChanged 1
#define VDA_CurrentVirtualDesktopChanged 0

#define VDA_IS_NORMAL 1
#define VDA_IS_MINIMIZED 2
#define VDA_IS_MAXIMIZED 3

std::map<HWND, int> listeners;
IServiceProvider* pServiceProvider = nullptr;
IVirtualDesktopManagerInternal *pDesktopManagerInternal = nullptr;
IVirtualDesktopManager *pDesktopManager = nullptr;
IApplicationViewCollection *viewCollection = nullptr;
IVirtualDesktopPinnedApps *pinnedApps = nullptr;
IVirtualDesktopNotificationService* pDesktopNotificationService = nullptr;
BOOL registeredForNotifications = FALSE;

DWORD idNotificationService = 0;

struct TempWindowEntry {
    HWND hwnd;
    ULONGLONG lastActivationTimestamp;
};


struct ChangeDesktopAction {
    GUID newDesktopGuid;
    GUID oldDesktopGuid;
};


struct ShowWindowOnDesktopAction {
    int desktopNumber;
    int cmdShow;
};



// common functions

HWND _ConvertPyHwndToHwnd(PyObject* pyHwnd) {
    HWND hwnd;
    if (PyInt_Check(pyHwnd)) {
        hwnd = (HWND)(INT32) PyInt_AsLong(pyHwnd);
    } else {
        hwnd = 0;
    }
    return hwnd;
}


GUID _ConvertPyGuidToGuid(PyObject* pGuid) {
    GUID guid = {0};

    char* sGuid = PyString_AsString(pGuid);

    ::UuidFromStringA((RPC_CSTR) sGuid, &guid);
    return guid;
}


static PyObject* _ConvertGuidToPyGuid(GUID guid) {
    /*
    std::stringstream stream;
    stream << std::hex << your_int;
    std::string result( stream.str() );
    */

    wchar_t* pWCBuffer;
    ::StringFromCLSID((const IID) guid, &pWCBuffer);

    size_t count;
    char *pMBBuffer = (char *)malloc(39);

    ::wcstombs_s(&count, pMBBuffer, (size_t)39, pWCBuffer, (size_t)39);

    PyObject* res = Py_BuildValue("s", pMBBuffer);

    if (pMBBuffer) {
        free(pMBBuffer);
    }

    return res;
}


IApplicationView* _GetApplicationViewFromHwnd(HWND hwnd) {
    if (hwnd == 0)
        return nullptr;
    IApplicationView* view = nullptr;
    viewCollection->GetViewForHwnd(hwnd, &view);
    return view;
}


IApplicationView* _GetViewFromPyWindowHwnd(HWND hwnd) {
    IApplicationView* view = nullptr;

    if (hwnd != 0) {
        view = _GetApplicationViewFromHwnd(hwnd);
    }
    return view;
}


int _GetDesktopNumberFromId(GUID desktopId) {
    IObjectArray *pObjectArray = nullptr;
    HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr)) {
        UINT count;
        hr = pObjectArray->GetCount(&count);

        if (!SUCCEEDED(hr)) {
            pObjectArray->Release();
            return -1;
        }

        for (UINT i = 0; i < count; i++) {
            IVirtualDesktop *pDesktop = nullptr;
            pObjectArray->GetAt(
                i,
                __uuidof(IVirtualDesktop),
                (void**)&pDesktop
            );

            if (pDesktop == nullptr) {
                continue;
            }

            GUID id = { 0 };

            if (SUCCEEDED(pDesktop->GetID(&id)) && id == desktopId) {
                pDesktop->Release();
                pObjectArray->Release();
                return i;
            }
            pDesktop->Release();
        }
    }
    pObjectArray->Release();
    return -1;
}


int _GetDesktopNumber(IVirtualDesktop *pDesktop) {
    if (pDesktop == nullptr) {
        return -1;
    }

    GUID guid;

    if (SUCCEEDED(pDesktop->GetID(&guid))) {
        return _GetDesktopNumberFromId(guid);
    }

    return -1;
}


IVirtualDesktop* _GetDesktopFromNumber(int number) {
    IObjectArray *pObjectArray = nullptr;
    IVirtualDesktop *pDesktop = nullptr;
    HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr)) {
        pObjectArray->GetAt(
            number,
            __uuidof(IVirtualDesktop),
            (void**)&pDesktop
        );
    }

    return pDesktop;
}


GUID _GetDesktopIdFromNumber(int number) {
    GUID id = {0};

    IVirtualDesktop* pDesktop = _GetDesktopFromNumber(number);
    if (pDesktop != nullptr) {
        pDesktop->GetID(&id);
        pDesktop->Release();
    }

    return id;
}

/*
LPWSTR _GetApplicationIdFromHwnd(HWND hwnd) {
    // TODO: This should not return a pointer, it should take in a pointer, or return either wstring or std::string

    if (hwnd == 0)
        return nullptr;
    IApplicationView* app = _GetApplicationViewFromHwnd(hwnd);
    if (app != nullptr) {
        LPWSTR appId = new TCHAR[1024];
        app->GetAppUserModelId(&appId);
        app->Release();
        return appId;
    }
    return nullptr;
}
*/

// IApplicationView -------------------------------------------------

static PyObject* ApplicationViewSetFocus(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    if (view != nullptr) {
        view->SetFocus();
        view->Release();
        return Py_BuildValue("i", 1);
    }
    view->Release();
    return Py_BuildValue("i", 0);
}


static PyObject* ApplicationViewGetFocus(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    IApplicationView* current_view;
    viewCollection->GetViewInFocus(&current_view);



    if ((current_view == nullptr) || (view == nullptr)) {
        view->Release();
        current_view->Release();
        return Py_BuildValue("i", -1);
    }

    HWND ret1 = 0;
    HWND ret2 = 0;

    view->GetThumbnailWindow(&ret1);
    current_view->GetThumbnailWindow(&ret2);

    view->Release();
    current_view->Release();

    if ((ret1 == 0) || (ret2 == 0)) {
        return Py_BuildValue("i", -1);
    }

    return Py_BuildValue("i", ret1 == ret2);
}


static PyObject* ApplicationViewSwitchTo(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    if (view != nullptr) {
        view->SwitchTo();
        view->Release();
        return Py_BuildValue("i", 1);
    }
    view->Release();
    return Py_BuildValue("i", 0);
}

static PyObject* ApplicationViewGetThumbnailWindow(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    HWND thumbHwnd;
    if (view != nullptr) {
        view->GetThumbnailWindow(&thumbHwnd);
        view->Release();
        return Py_BuildValue("l", thumbHwnd);
    }
    view->Release();
    return Py_BuildValue("l", -1);
}


static PyObject* ApplicationViewGetVisibility(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    int visibility;
    if (view != nullptr) {
        view->GetVisibility(&visibility);
        view->Release();
        return Py_BuildValue("i", visibility);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* ApplicationViewSetCloak(PyObject* self, PyObject* args) {
    HWND hwnd;
    APPLICATION_VIEW_CLOAK_TYPE cloakType;
    int unknown;
    PyArg_ParseTuple(args, "iii", &hwnd, &cloakType, &unknown);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    HRESULT res;
    if (view != nullptr) {
        res = view->SetCloak(cloakType, unknown);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* ApplicationViewGetExtendedFramePosition(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    RECT rect;
    PyObject* res = PyDict_New();
    if (view != nullptr) {
        view->GetExtendedFramePosition(&rect);
        PyDict_SetItem(
            res,
            Py_BuildValue("s", "left"),
            Py_BuildValue("i", rect.left)
        );
        PyDict_SetItem(
            res,
            Py_BuildValue("s", "top"),
            Py_BuildValue("i", rect.top)
        );
        PyDict_SetItem(
            res,
            Py_BuildValue("s", "right"),
            Py_BuildValue("i", rect.right)
        );
        PyDict_SetItem(
            res,
            Py_BuildValue("s", "bottom"),
            Py_BuildValue("i", rect.bottom)
        );
    }
    view->Release();
    return res;
}



static PyObject* ApplicationViewGetViewState(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        UINT state;
        view->GetViewState(&state);
        view->Release();
        return Py_BuildValue("i", state);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewSetViewState(PyObject* self, PyObject* args) {
    HWND hwnd;
    UINT state;
    PyArg_ParseTuple(args, "ii", &hwnd, &state);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        HRESULT res;
        res = view->SetViewState(state);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewGetVirtualDesktopId(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        GUID guid;
        view->GetVirtualDesktopId(&guid);
        view->Release();
        return _ConvertGuidToPyGuid(guid);
    }
    view->Release();
    return Py_BuildValue("s", "");
}



static PyObject* ApplicationViewSetVirtualDesktopId(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyObject* sGuid;
    PyArg_ParseTuple(args, "iO", &hwnd, &sGuid);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        GUID guid = _ConvertPyGuidToGuid(sGuid);
        if (guid.Data1 != 0) {
            HRESULT res;
            res = view->SetVirtualDesktopId(guid);
            view->Release();
            return Py_BuildValue("i", res);
        }
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewGetShowInSwitchers(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        int switcher;
        view->GetShowInSwitchers(&switcher);
        view->Release();
        return Py_BuildValue("i", switcher);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* ApplicationViewSetShowInSwitchers(PyObject* self, PyObject* args) {
    HWND hwnd;
    int switcher;
    PyArg_ParseTuple(args, "ii", &hwnd, &switcher);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        HRESULT res;
        res = view->SetShowInSwitchers(switcher);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* ApplicationViewGetScaleFactor(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        int factor;
        view->GetScaleFactor(&factor);
        view->Release();
        return Py_BuildValue("i", factor);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* ApplicationViewCanReceiveInput(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        BOOL res;
        view->CanReceiveInput(&res);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewIsTray(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        BOOL res;
        view->IsTray(&res);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewIsSplashScreenPresented(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        BOOL res;
        view->IsSplashScreenPresented(&res);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}



static PyObject* ApplicationViewFlash(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        HRESULT res;
        res = view->Flash();
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* ApplicationViewIsMirrored(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        BOOL res;
        view->IsMirrored(&res);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}



// IVirtualDesktopPinnedApps -------------------------------------------------
static PyObject* DesktopPinnedAppsIsAppIdPinned(PyObject* self, PyObject* args) {
    char* appId;
    PyArg_ParseTuple(args, "s", &appId);
    BOOL res;
    pinnedApps->IsAppIdPinned((PCWSTR)appId, &res);
    return Py_BuildValue("i", res);

}

static PyObject* DesktopPinnedAppsPinAppID(PyObject* self, PyObject* args) {
    char* appId;
    PyArg_ParseTuple(args, "s", &appId);
    HRESULT res;
    res = pinnedApps->PinAppID((PCWSTR)appId);
    return Py_BuildValue("i", res);

}

static PyObject* DesktopPinnedAppsUnpinAppID(PyObject* self, PyObject* args) {
    char* appId;
    PyArg_ParseTuple(args, "s", &appId);
    HRESULT res;
    res = pinnedApps->UnpinAppID((PCWSTR)appId);
    return Py_BuildValue("i", res);

}

static PyObject* DesktopPinnedAppsIsViewPinned(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        BOOL res;
        pinnedApps->IsViewPinned(view, &res);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* DesktopPinnedAppsPinView(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        HRESULT res;
        res = pinnedApps->PinView(view);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}

static PyObject* DesktopPinnedAppsUnpinView(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        HRESULT res;
        res = pinnedApps->UnpinView(view);
        view->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}


/*
IApplicationViewCollection
GetViews(&IObjectArray)
GetViewsByZOrder(&IObjectArray)
GetViewsByAppUserModelId(PCWSTR, &IObjectArray)
GetViewForHwnd(HWND, &IApplicationView)
GetViewForApplication(&IImmersiveApplication, &IApplicationView)
GetViewForAppUserModelId(PCWSTR, &IApplicationView)
GetViewInFocus(&IApplicationView)
RefreshCollection()
*/


// IVirtualDesktop -----------------------------------------------------------

static PyObject* DesktopIsViewVisible(PyObject* self, PyObject* args) {
    PyObject* sGuid;
    HWND hwnd;
    PyArg_ParseTuple(args, "Oi", &sGuid, &hwnd);

    GUID guid = _ConvertPyGuidToGuid(sGuid);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if ((view == nullptr) || (guid.Data1 == 0)) {
        view->Release();
        return Py_BuildValue("i", -1);
    }

    IVirtualDesktop* desktop = nullptr;

    pDesktopManagerInternal->FindDesktop(&guid, &desktop);
    if (desktop == nullptr) {
        view->Release();
        desktop->Release();
        return Py_BuildValue("i", -1);
    }

    int res;
    desktop->IsViewVisible(view, &res);
    view->Release();
    desktop->Release();
    return Py_BuildValue("i", res);
}


// IVirtualDesktopManagerInternal ---------------------------------------
static PyObject* DesktopManagerInternalGetCount(PyObject* self) {
    UINT count;
    pDesktopManagerInternal->GetCount(&count);
    return Py_BuildValue("i", count);
}


static PyObject* DesktopManagerInternalMoveViewToDesktop(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyObject* sGuid;
    PyArg_ParseTuple(args, "Oi", &sGuid, &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        GUID guid = _ConvertPyGuidToGuid(sGuid);
        if (guid.Data1 == 0) {
            view->Release();
            return Py_BuildValue("i", -1);
        }

        IVirtualDesktop* desktop = nullptr;

        pDesktopManagerInternal->FindDesktop(&guid, &desktop);
        if (desktop == nullptr) {
            view->Release();
            desktop->Release();
            return Py_BuildValue("i", -1);
        }

        HRESULT res;
        res = pDesktopManagerInternal->MoveViewToDesktop(view, desktop);
        view->Release();
        desktop->Release();
        return Py_BuildValue("i", res);
    }
    view->Release();
    return Py_BuildValue("i", -1);
}



static PyObject* DesktopManagerInternalCanViewMoveDesktops(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if (view != nullptr) {
        int res;
        res = pDesktopManagerInternal->CanViewMoveDesktops(view, &res);
        view->Release();
        return Py_BuildValue("i", res);
    }

    view->Release();
    return Py_BuildValue("i", -1);
}


static PyObject* DesktopManagerInternalGetCurrentDesktop(PyObject* self) {
    IVirtualDesktop* desktop = nullptr;
    pDesktopManagerInternal->GetCurrentDesktop(&desktop);
    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("s", "");
    }
    GUID guid;
    desktop->GetID(&guid);

    desktop->Release();
    return _ConvertGuidToPyGuid(guid);
}


static PyObject* DesktopManagerInternalGetDesktopIds(PyObject* self) {
    IObjectArray *pObjectArray = nullptr;
    HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);
    int found = -1;

    if (!SUCCEEDED(hr)) {
        pObjectArray->Release();
        return PyList_New(0);
    }

    UINT count;
    hr = pObjectArray->GetCount(&count);

    if (!SUCCEEDED(hr)) {
        pObjectArray->Release();
        return PyList_New(0);
    }

    PyObject* pyList = PyList_New(count);

    for (UINT i = 0; i < count; i++) {
        IVirtualDesktop *pDesktop = nullptr;

        if (FAILED(pObjectArray->GetAt(i, __uuidof(IVirtualDesktop), (void**)&pDesktop))) {
            continue;
        }

        GUID guid = { 0 };
        if (SUCCEEDED(pDesktop->GetID(&guid))) {
            PyList_SetItem(
                pyList,
                i,
                _ConvertGuidToPyGuid(guid)
            );
        }

        pDesktop->Release();
    }
    pObjectArray->Release();

    return pyList;
}


static PyObject* DesktopManagerInternalGetAdjacentDesktop(PyObject* self, PyObject* args) {
    PyObject* sGuid;
    int direction;
    PyArg_ParseTuple(args, "Oi", &sGuid, &direction);

    GUID guid = _ConvertPyGuidToGuid(sGuid);
    if (guid.Data1 == 0) {
        return Py_BuildValue("s", "");
    }

    IVirtualDesktop* desktop = nullptr;

    pDesktopManagerInternal->FindDesktop(&guid, &desktop);
    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("s", "");
    }
    IVirtualDesktop* neighbor = nullptr;

    pDesktopManagerInternal->GetAdjacentDesktop(desktop, (AdjacentDesktop)direction, &neighbor);

    if (neighbor == nullptr) {
        desktop->Release();
        neighbor->Release();
        return Py_BuildValue("s", "");
    }

    GUID neighbor_guid;
    neighbor->GetID(&neighbor_guid);

    desktop->Release();
    neighbor->Release();

    return _ConvertGuidToPyGuid(neighbor_guid);
}


static PyObject* DesktopManagerInternalSwitchDesktop(PyObject* self, PyObject* args) {
    PyObject* sGuid;
    PyArg_ParseTuple(args, "O", &sGuid);

    GUID guid = _ConvertPyGuidToGuid(sGuid);
    if (guid.Data1 == 0) {
        return Py_BuildValue("l", -1);
    }

    IVirtualDesktop* desktop = nullptr;

    pDesktopManagerInternal->FindDesktop(&guid, &desktop);
    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("l", -1);
    }

    HRESULT res;
    res = pDesktopManagerInternal->SwitchDesktop(desktop);
    desktop->Release();
    return Py_BuildValue("l", res);
}



static PyObject* DesktopManagerInternalCreateDesktop(PyObject* self) {
    IVirtualDesktop* desktop = nullptr;
    pDesktopManagerInternal->CreateDesktopW(&desktop);

    if (desktop == nullptr) {
        return Py_BuildValue("s", "");
    }

    GUID guid;
    desktop->GetID(&guid);
    desktop->Release();
    return _ConvertGuidToPyGuid(guid);
}


static PyObject* DesktopManagerInternalRemoveDesktop(PyObject* self, PyObject* args) {
    PyObject* sGuid1;
    PyObject* sGuid2;
    PyArg_ParseTuple(args, "OO", &sGuid1, &sGuid2);

    GUID guid1 = _ConvertPyGuidToGuid(sGuid1);
    GUID guid2 = _ConvertPyGuidToGuid(sGuid2);
    if ((guid1.Data1 == 0) || (guid2.Data1 == 0)) {
        return Py_BuildValue("l", -1);
    }

    IVirtualDesktop* desktop1 = nullptr;
    IVirtualDesktop* desktop2 = nullptr;

    pDesktopManagerInternal->FindDesktop(&guid1, &desktop1);
    pDesktopManagerInternal->FindDesktop(&guid2, &desktop2);

    if ((desktop1 == nullptr) || (desktop2 == nullptr)) {
        desktop1->Release();
        desktop2->Release();
        return Py_BuildValue("l", -1);
    }

    HRESULT res;
    res = pDesktopManagerInternal->RemoveDesktop(desktop1, desktop2);
    desktop1->Release();
    desktop2->Release();
    return Py_BuildValue("l", res);
}


//IVirtualDesktopManager

static PyObject* DesktopManagerIsWindowOnCurrentVirtualDesktop(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);

    if (hwnd == 0) {
        return Py_BuildValue("i", -1);
    }

    BOOL res;
    pDesktopManager->IsWindowOnCurrentVirtualDesktop(hwnd, &res);
    return Py_BuildValue("i", res);
}


static PyObject* DesktopManagerGetWindowDesktopId(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "i", &hwnd);

    if (hwnd == 0) {
        return Py_BuildValue("s", "");
    }

    GUID guid;
    pDesktopManager->GetWindowDesktopId(hwnd, &guid);
    return _ConvertGuidToPyGuid(guid);
}


static PyObject* DesktopManagerMoveWindowToDesktop(PyObject* self, PyObject* args) {
    PyObject* sGuid;
    HWND hwnd;
    PyArg_ParseTuple(args, "Oi", &sGuid, &hwnd);

    GUID guid = _ConvertPyGuidToGuid(sGuid);

    if ((hwnd == 0) || (guid.Data1 == 0)) {
        return Py_BuildValue("l", -1);
    }

    HRESULT res;
    res = pDesktopManager->MoveWindowToDesktop(hwnd, guid);
    return Py_BuildValue("l", res);
}


// Misc Functions ------------------------------------------------


static PyObject* GetCurrentDesktopNumber(PyObject* self) {
    IVirtualDesktop* desktop = nullptr;

    pDesktopManagerInternal->GetCurrentDesktop(&desktop);
    if (desktop == nullptr) {
        return Py_BuildValue("i", -1);
    }
    int number = _GetDesktopNumber(desktop);
    desktop->Release();
    return Py_BuildValue("i", number);
}



static PyObject* GetDesktopNumberFromId(PyObject* self, PyObject* args) {
    PyObject* sGuid;
    PyArg_ParseTuple(args, "O", &sGuid);

    GUID guid = _ConvertPyGuidToGuid(sGuid);

    if (guid.Data1 == 0) {
        return Py_BuildValue("l", -1);
    }

    IVirtualDesktop* desktop = nullptr;
    pDesktopManagerInternal->FindDesktop(&guid, &desktop);
    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("l", -1);
    }

    int number = _GetDesktopNumber(desktop);
    desktop->Release();
    return Py_BuildValue("i", number);
}


static PyObject* GetDesktopIdFromNumber(PyObject* self, PyObject* args) {
    int number;

    PyArg_ParseTuple(args, "i", &number);
    GUID guid = _GetDesktopIdFromNumber(number);

    if (guid.Data1 == 0) {
        return Py_BuildValue("s", "");
    }

    return _ConvertGuidToPyGuid(guid);
}



// Notifications ------------------------------------



void _PostMessageToListeners(int msgOffset, WPARAM wParam, LPARAM lParam) {
    for each (std::pair<HWND, int> listener in listeners) {
        PostMessage(listener.first, listener.second + msgOffset, wParam, lParam);
    }
}



class _Notifications : public IVirtualDesktopNotification {
private:
    ULONG _referenceCount;
public:
    // Inherited via IVirtualDesktopNotification
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void ** ppvObject) override
    {
        // Always set out parameter to NULL, validating it first.
        if (!ppvObject)
            return E_INVALIDARG;
        *ppvObject = NULL;

        if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification)
        {
            // Increment the reference count and return the pointer.
            *ppvObject = (LPVOID)this;
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    virtual ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&_referenceCount);
    }

    virtual ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG result = InterlockedDecrement(&_referenceCount);
        if (result == 0)
        {
            delete this;
        }
        return 0;
    }
    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopCreated(IVirtualDesktop * pDesktop) override
    {
        _PostMessageToListeners(VDA_VirtualDesktopCreated, _GetDesktopNumber(pDesktop), 0);
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyBegin(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
    {
        _PostMessageToListeners(VDA_VirtualDesktopDestroyBegin, _GetDesktopNumber(pDesktopDestroyed), _GetDesktopNumber(pDesktopFallback));
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyFailed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
    {
        _PostMessageToListeners(VDA_VirtualDesktopDestroyFailed, _GetDesktopNumber(pDesktopDestroyed), _GetDesktopNumber(pDesktopFallback));
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
    {
        _PostMessageToListeners(VDA_VirtualDesktopDestroyed, _GetDesktopNumber(pDesktopDestroyed), _GetDesktopNumber(pDesktopFallback));
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE ViewVirtualDesktopChanged(IApplicationView * pView) override
    {
        _PostMessageToListeners(VDA_ViewVirtualDesktopChanged, 0, 0);
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE CurrentVirtualDesktopChanged(
        IVirtualDesktop *pDesktopOld,
        IVirtualDesktop *pDesktopNew) override
    {
        viewCollection->RefreshCollection();
        ChangeDesktopAction act;
        if (pDesktopOld != nullptr) {
            pDesktopOld->GetID(&act.oldDesktopGuid);
        }
        if (pDesktopNew != nullptr) {
            pDesktopNew->GetID(&act.newDesktopGuid);
        }

        _PostMessageToListeners(VDA_CurrentVirtualDesktopChanged, _GetDesktopNumberFromId(act.oldDesktopGuid), _GetDesktopNumberFromId(act.newDesktopGuid));
        return S_OK;
    }
};

void _RegisterDesktopNotifications() {
    if (pDesktopNotificationService == nullptr) {
        return;
    }
    if (registeredForNotifications) {
        return;
    }

    // TODO: This is never deleted
    _Notifications *nf = new _Notifications();
    HRESULT res = pDesktopNotificationService->Register(nf, &idNotificationService);
    if (SUCCEEDED(res)) {
        registeredForNotifications = TRUE;
    }
}


void RegisterPostMessageHook(HWND listener, int messageOffset) {
    listeners.insert(std::pair<HWND, int>(listener, messageOffset));
    if (listeners.size() != 1) {
        return;
    }
    _RegisterDesktopNotifications();
}

void UnregisterPostMessageHook(HWND hwnd) {
    listeners.erase(hwnd);
    if (listeners.size() != 0) {
        return;
    }

    if (pDesktopNotificationService == nullptr) {
        return;
    }

    if (idNotificationService > 0) {
        registeredForNotifications = TRUE;
        pDesktopNotificationService->Unregister(idNotificationService);
    }
}


// Module definitions ----------------------------------------------------

struct module_state {
    PyObject *error;
};


#if PY_MAJOR_VERSION >= 3
#define GETSTATE(m) ((struct module_state*)PyModule_GetState(m))
#else
#define GETSTATE(m) (&_state)
static struct module_state _state;
#endif


static PyObject* error_out(PyObject *m) {
    struct module_state *st = GETSTATE(m);
    PyErr_SetString(st->error, "something bad happened");
    return NULL;
}


static PyMethodDef module_methods[] = {
    {"ApplicationViewSetFocus", (PyCFunction)ApplicationViewSetFocus, METH_VARARGS, NULL},
    {"ApplicationViewGetFocus", (PyCFunction)ApplicationViewGetFocus, METH_VARARGS, NULL},
    {"ApplicationViewSwitchTo", (PyCFunction)ApplicationViewSwitchTo, METH_VARARGS, NULL},
    {"ApplicationViewGetThumbnailWindow", (PyCFunction)ApplicationViewGetThumbnailWindow, METH_VARARGS, NULL},
    {"ApplicationViewGetVisibility", (PyCFunction)ApplicationViewGetVisibility, METH_VARARGS, NULL},
    {"ApplicationViewSetCloak", (PyCFunction)ApplicationViewSetCloak, METH_VARARGS, NULL},
    {"ApplicationViewGetExtendedFramePosition", (PyCFunction)ApplicationViewGetExtendedFramePosition, METH_VARARGS, NULL},
    {"ApplicationViewGetViewState", (PyCFunction)ApplicationViewGetViewState, METH_VARARGS, NULL},
    {"ApplicationViewSetViewState", (PyCFunction)ApplicationViewSetViewState, METH_VARARGS, NULL},
    {"ApplicationViewGetVirtualDesktopId", (PyCFunction)ApplicationViewGetVirtualDesktopId, METH_VARARGS, NULL},
    {"ApplicationViewSetVirtualDesktopId", (PyCFunction)ApplicationViewSetVirtualDesktopId, METH_VARARGS, NULL},
    {"ApplicationViewGetShowInSwitchers", (PyCFunction)ApplicationViewGetShowInSwitchers, METH_VARARGS, NULL},
    {"ApplicationViewSetShowInSwitchers", (PyCFunction)ApplicationViewSetShowInSwitchers, METH_VARARGS, NULL},
    {"ApplicationViewGetScaleFactor", (PyCFunction)ApplicationViewGetScaleFactor, METH_VARARGS, NULL},
    {"ApplicationViewCanReceiveInput", (PyCFunction)ApplicationViewCanReceiveInput, METH_VARARGS, NULL},
    {"ApplicationViewIsTray", (PyCFunction)ApplicationViewIsTray, METH_VARARGS, NULL},
    {"ApplicationViewIsSplashScreenPresented", (PyCFunction)ApplicationViewIsSplashScreenPresented, METH_VARARGS, NULL},
    {"ApplicationViewFlash", (PyCFunction)ApplicationViewFlash, METH_VARARGS, NULL},
    {"ApplicationViewIsMirrored", (PyCFunction)ApplicationViewIsMirrored, METH_VARARGS, NULL},
    {"DesktopPinnedAppsIsAppIdPinned", (PyCFunction)DesktopPinnedAppsIsAppIdPinned, METH_VARARGS, NULL},
    {"DesktopPinnedAppsPinAppID", (PyCFunction)DesktopPinnedAppsPinAppID, METH_VARARGS, NULL},
    {"DesktopPinnedAppsUnpinAppID", (PyCFunction)DesktopPinnedAppsUnpinAppID, METH_VARARGS, NULL},
    {"DesktopPinnedAppsIsViewPinned", (PyCFunction)DesktopPinnedAppsIsViewPinned, METH_VARARGS, NULL},
    {"DesktopPinnedAppsPinView", (PyCFunction)DesktopPinnedAppsPinView, METH_VARARGS, NULL},
    {"DesktopPinnedAppsUnpinView", (PyCFunction)DesktopPinnedAppsUnpinView, METH_VARARGS, NULL},
    {"DesktopIsViewVisible", (PyCFunction)DesktopIsViewVisible, METH_VARARGS, NULL},
    {"DesktopManagerInternalGetCount", (PyCFunction)DesktopManagerInternalGetCount, METH_NOARGS, NULL},
    {"DesktopManagerInternalMoveViewToDesktop", (PyCFunction)DesktopManagerInternalMoveViewToDesktop, METH_VARARGS, NULL},
    {"DesktopManagerInternalCanViewMoveDesktops", (PyCFunction)DesktopManagerInternalCanViewMoveDesktops, METH_VARARGS, NULL},
    {"DesktopManagerInternalGetCurrentDesktop", (PyCFunction)DesktopManagerInternalGetCurrentDesktop, METH_NOARGS, NULL},
    {"DesktopManagerInternalGetDesktopIds", (PyCFunction)DesktopManagerInternalGetDesktopIds, METH_NOARGS, NULL},
    {"DesktopManagerInternalGetAdjacentDesktop", (PyCFunction)DesktopManagerInternalGetAdjacentDesktop, METH_VARARGS, NULL},
    {"DesktopManagerInternalSwitchDesktop", (PyCFunction)DesktopManagerInternalSwitchDesktop, METH_VARARGS, NULL},
    {"DesktopManagerInternalCreateDesktop", (PyCFunction)DesktopManagerInternalCreateDesktop, METH_NOARGS, NULL},
    {"DesktopManagerInternalRemoveDesktop", (PyCFunction)DesktopManagerInternalRemoveDesktop, METH_VARARGS, NULL},
    {"DesktopManagerIsWindowOnCurrentVirtualDesktop", (PyCFunction)DesktopManagerIsWindowOnCurrentVirtualDesktop, METH_VARARGS, NULL},
    {"DesktopManagerGetWindowDesktopId", (PyCFunction)DesktopManagerGetWindowDesktopId, METH_VARARGS, NULL},
    {"DesktopManagerMoveWindowToDesktop", (PyCFunction)DesktopManagerMoveWindowToDesktop, METH_VARARGS, NULL},
    {"GetDesktopNumberFromId", (PyCFunction)GetDesktopNumberFromId, METH_VARARGS, NULL},
    {"GetCurrentDesktopNumber", (PyCFunction)GetCurrentDesktopNumber, METH_NOARGS, NULL},
    {"GetDesktopIdFromNumber", (PyCFunction)GetDesktopIdFromNumber, METH_VARARGS, NULL},
    { NULL, NULL, 0, NULL }
};


#if PY_MAJOR_VERSION >= 3

static int myextension_traverse(PyObject *m, visitproc visit, void *arg) {
    Py_VISIT(GETSTATE(m)->error);
    return 0;
}

static int myextension_clear(PyObject *m) {
    Py_CLEAR(GETSTATE(m)->error);
    return 0;
}

#define INITERROR return NULL


static struct PyModuleDef libWinVirtualDesktop = {
    PyModuleDef_HEAD_INIT,
    "libWinVirtualDesktop", /* name of module */
    NULL, /* module documentation, may be NULL */
    -1,   /* size of per-interpreter state of the module, or -1 if the module keeps state in global variables. */
    module_methods,
    NULL,                /* m_reload */
    NULL,                /* m_traverse */
    NULL,                /* m_clear */
    NULL,                /* m_free */
};

PyMODINIT_FUNC PyInit_libWinVirtualDesktop(void)

#else
#define INITERROR return

void
initlibWinVirtualDesktop(void)
#endif
{

#if PY_MAJOR_VERSION >= 3
    PyObject *module =  PyModule_Create(&libWinVirtualDesktop);
#else
    PyObject *module = Py_InitModule("libWinVirtualDesktop", module_methods);
#endif

    if (module == NULL) {
        INITERROR;
    }

    PyObject* ExceptionBase = PyErr_NewExceptionWithDoc(
        "libWinVirtualDesktop.ExceptionBase", /* char *name */
        "Base exception class for the libWinVirtualDesktop module.", /* char *doc */
        NULL, /* PyObject *base */
        NULL /* PyObject *dict */
    );

    if (ExceptionBase == NULL) {
        Py_DECREF(module);
        INITERROR;
    }

    PyModule_AddObject(module, "ExceptionBase", ExceptionBase);

    PyObject* ServiceProviderError = PyErr_NewExceptionWithDoc(
        "libWinVirtualDesktop.ServiceProviderError", /* char *name */
        "ServiceProvider error.", /* char *doc */
        ExceptionBase, /* PyObject *base */
        NULL /* PyObject *dict */
    );

    if (ServiceProviderError == NULL) {
        Py_DECREF(module);
        INITERROR;
    }

    PyModule_AddObject(module, "ServiceProviderError", ServiceProviderError);

    PyObject* ViewCollectionError = PyErr_NewExceptionWithDoc(
        "libWinVirtualDesktop.ViewCollectionError", /* char *name */
        "ViewCollection error.", /* char *doc */
        ExceptionBase, /* PyObject *base */
        NULL /* PyObject *dict */
    );

    if (ViewCollectionError == NULL) {
        Py_DECREF(module);
        INITERROR;
    }

    PyModule_AddObject(module, "ViewCollectionError", ViewCollectionError);

    PyObject* DesktopManagerInternalError = PyErr_NewExceptionWithDoc(
        "libWinVirtualDesktop.DesktopManagerInternalError", /* char *name */
        "DesktopManagerInternal error.", /* char *doc */
        ExceptionBase, /* PyObject *base */
        NULL /* PyObject *dict */
    );

    if (DesktopManagerInternalError == NULL) {
        Py_DECREF(module);
        INITERROR;
    }

    PyModule_AddObject(module, "DesktopManagerInternalError", DesktopManagerInternalError);

    PyObject* NotificationServiceError = PyErr_NewExceptionWithDoc(
        "libWinVirtualDesktop.NotificationServiceError", /* char *name */
        "NotificationService error.", /* char *doc */
        ExceptionBase, /* PyObject *base */
        NULL /* PyObject *dict */
    );

    if (NotificationServiceError == NULL) {
        Py_DECREF(module);
        INITERROR;
    }

    PyModule_AddObject(module, "NotificationServiceError", NotificationServiceError);

    pServiceProvider = nullptr;
    pDesktopManagerInternal = nullptr;
    pDesktopManager = nullptr;
    viewCollection = nullptr;
    pinnedApps = nullptr;
    pDesktopNotificationService = nullptr;
    registeredForNotifications = FALSE;

    ::CoInitialize(NULL);
    ::CoCreateInstance(
        CLSID_ImmersiveShell,
        NULL,
        CLSCTX_LOCAL_SERVER,
        __uuidof(IServiceProvider),
        (PVOID*)&pServiceProvider
    );

    if (pServiceProvider == nullptr) {
        PyErr_SetString(ServiceProviderError, "FATAL ERROR");
        INITERROR;
    }

    pServiceProvider->QueryService(
        __uuidof(IApplicationViewCollection),
        &viewCollection
    );

    pServiceProvider->QueryService(
        __uuidof(IVirtualDesktopManager),
        &pDesktopManager
    );

    pServiceProvider->QueryService(
        CLSID_VirtualDesktopPinnedApps,
        __uuidof(IVirtualDesktopPinnedApps),
        (PVOID*)&pinnedApps
    );

    pServiceProvider->QueryService(
        CLSID_VirtualDesktopManagerInternal,
        __uuidof(IVirtualDesktopManagerInternal),
        (PVOID*)&pDesktopManagerInternal
    );

    // Notification service
    pServiceProvider->QueryService(
        CLSID_IVirtualNotificationService,
        __uuidof(IVirtualDesktopNotificationService),
        (PVOID*)&pDesktopNotificationService
    );

    if (viewCollection == nullptr) {
        IApplicationViewCollectionOlder *tmpViewCollection = nullptr;
        pServiceProvider->QueryService(
            __uuidof(IApplicationViewCollectionOlder),
            &tmpViewCollection
        );

        if (tmpViewCollection == nullptr) {
            PyErr_SetString(ViewCollectionError, "FATAL ERROR");
            INITERROR;
        }

        viewCollection = tmpViewCollection;
    }

    if (pDesktopManagerInternal == nullptr) {
        PyErr_SetString(DesktopManagerInternalError, "FATAL ERROR");
        INITERROR;
    }

    if (pDesktopNotificationService == nullptr) {
        PyErr_SetString(NotificationServiceError, "FATAL ERROR");
        INITERROR;
    }

#if PY_MAJOR_VERSION >= 3
    return module;
#endif
}



