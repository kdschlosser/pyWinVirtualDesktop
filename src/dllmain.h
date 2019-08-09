#pragma once
#include "stdafx.h"
#include "VirtualDesktop.h"
#include <Python.h>
#include <string.h>
#include <Rpc.h>

#define string std::string

#define VIRTUAL_DESKTOP_CREATED 5
#define VIRTUAL_DESKTOP_DESTROY_BEGIN 4
#define VIRTUAL_DESKTOP_DESTROY_FAILED 3
#define VIRTUAL_DESKTOP_DESTROYED 2
#define VIRTUAL_DESKTOP_VIEW_CHANGED 1
#define VIRTUAL_DESKTOP_CURRENT_CHANGED 0

#define VDA_IS_NORMAL 1
#define VDA_IS_MINIMIZED 2
#define VDA_IS_MAXIMIZED 3

IServiceProvider* pServiceProvider = nullptr;
IVirtualDesktopManagerInternal *pDesktopManagerInternal = nullptr;
IVirtualDesktopManager *pDesktopManager = nullptr;
IApplicationViewCollection *viewCollection = nullptr;
IVirtualDesktopPinnedApps *pinnedApps = nullptr;
IVirtualDesktopNotificationService* pDesktopNotificationService = nullptr;

static PyObject *notificationCallback = nullptr;
BOOL registeredForNotifications = FALSE;
DWORD idNotificationService = 0;


GUID _ConvertPyGuidToGuid(char* sGuid) {
    GUID guid = {0};
    ::UuidFromString((RPC_CSTR)sGuid, &guid);

    return guid;
}

void _ConvertGuidToCString(GUID guid, char* res) {
    wchar_t* pWCBuffer;
    ::StringFromCLSID((const IID) guid, &pWCBuffer);

    size_t count;
    char* pMBBuffer = (char *)malloc(39);

    ::wcstombs_s(&count, pMBBuffer, (size_t)39, pWCBuffer, (size_t)39);

    strcpy_s(res, (size_t)39, pMBBuffer);

    if (pMBBuffer) {
        free(pMBBuffer);
    }
}


static PyObject* _ConvertGuidToPyGuid(GUID guid) {
    char res[39] = {""};
    _ConvertGuidToCString(guid, res);
    return Py_BuildValue("s", res);
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


IVirtualDesktop* _GetDesktopFromStringId(char* guid) {
    IObjectArray *pObjectArray = nullptr;
    HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr)) {
        UINT count;
        hr = pObjectArray->GetCount(&count);

        if (!SUCCEEDED(hr)) {
            pObjectArray->Release();

            IVirtualDesktop* desktop = nullptr;
            return desktop;
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

            if (SUCCEEDED(pDesktop->GetID(&id))) {
                wchar_t* pWCBuffer;
                ::StringFromCLSID((const IID) id, &pWCBuffer);

                size_t count;
                char *pMBBuffer = (char *)malloc(39);

                ::wcstombs_s(&count, pMBBuffer, (size_t)39, pWCBuffer, (size_t)39);

                if (strcmp(guid, pMBBuffer) == 0) {
                    free(pMBBuffer);
                    pObjectArray->Release();
                    return pDesktop;

                }

                if (pMBBuffer) {
                    free(pMBBuffer);
                }
            }
            pDesktop->Release();
        }
    }
    pObjectArray->Release();

    IVirtualDesktop* desktop = nullptr;
    return desktop;
}


IVirtualDesktop* _GetDesktop(GUID guid) {
    IObjectArray *pObjectArray = nullptr;
    HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr)) {
        UINT count;
        hr = pObjectArray->GetCount(&count);

        if (!SUCCEEDED(hr)) {
            pObjectArray->Release();

            IVirtualDesktop* desktop = nullptr;
            return desktop;
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

            if (SUCCEEDED(pDesktop->GetID(&id)) && id == guid) {
                pObjectArray->Release();
                return pDesktop;
            }
            pDesktop->Release();
        }
    }
    pObjectArray->Release();

    IVirtualDesktop* desktop = nullptr;
    return desktop;
}


// IApplicationView -------------------------------------------------

static PyObject* ApplicationViewSetFocus(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "lii", &hwnd, &cloakType, &unknown);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "li", &hwnd, &state);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    char* sGuid;
    PyArg_ParseTuple(args, "ls", &hwnd, &sGuid);

    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

    if ((view == nullptr) || (desktop == nullptr)) {
        view->Release();
        desktop->Release();
        return Py_BuildValue("i", -1);
    }

    HRESULT res;
    GUID guid;

    desktop->GetID(&guid);

    res = view->SetVirtualDesktopId(guid);
    view->Release();
    desktop->Release();
    return Py_BuildValue("i", res);
}


static PyObject* ApplicationViewGetShowInSwitchers(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "li", &hwnd, &switcher);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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
    PyArg_ParseTuple(args, "l", &hwnd);
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


// IVirtualDesktop -----------------------------------------------------------

static PyObject* DesktopIsViewVisible(PyObject* self, PyObject* args) {
    char* sGuid;
    HWND hwnd;
    PyArg_ParseTuple(args, "sl", &sGuid, &hwnd);

    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);
    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);

    if ((view == nullptr) || (desktop == nullptr)) {
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
    char* sGuid;
    PyArg_ParseTuple(args, "sl", &sGuid, &hwnd);

    IApplicationView* view = _GetViewFromPyWindowHwnd(hwnd);
    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

    if ((view == nullptr) || (desktop == nullptr)) {
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


static PyObject* DesktopManagerInternalCanViewMoveDesktops(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "l", &hwnd);
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
    char* sGuid;
    AdjacentDesktop direction;
    PyArg_ParseTuple(args, "si", &sGuid, &direction);

    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("s", "");
    }

    IVirtualDesktop* neighbor = nullptr;

    pDesktopManagerInternal->GetAdjacentDesktop(desktop, direction, &neighbor);

    if (neighbor == nullptr) {
        desktop->Release();
        neighbor->Release();
        return Py_BuildValue("s", "");
    }

    GUID neighbor_guid;
    neighbor->GetID(&neighbor_guid);

    PyObject* res = _ConvertGuidToPyGuid(neighbor_guid);

    desktop->Release();
    neighbor->Release();
    return  res;
}


static PyObject* DesktopManagerInternalSwitchDesktop(PyObject* self, PyObject* args) {
    char* sGuid;
    PyArg_ParseTuple(args, "s", &sGuid);

    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

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
    char* sGuid1;
    char* sGuid2;
    PyArg_ParseTuple(args, "ss", &sGuid1, &sGuid2);

    IVirtualDesktop* desktop1 = _GetDesktopFromStringId(sGuid1);
    IVirtualDesktop* desktop2 = _GetDesktopFromStringId(sGuid2);

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
    PyArg_ParseTuple(args, "l", &hwnd);

    if (hwnd == 0) {
        return Py_BuildValue("i", -1);
    }

    BOOL res;
    pDesktopManager->IsWindowOnCurrentVirtualDesktop(hwnd, &res);
    return Py_BuildValue("i", res);
}


static PyObject* DesktopManagerGetWindowDesktopId(PyObject* self, PyObject* args) {
    HWND hwnd;
    PyArg_ParseTuple(args, "l", &hwnd);

    if (hwnd == 0) {
        return Py_BuildValue("s", "");
    }

    GUID guid;
    pDesktopManager->GetWindowDesktopId(hwnd, &guid);
    return _ConvertGuidToPyGuid(guid);
}


static PyObject* DesktopManagerMoveWindowToDesktop(PyObject* self, PyObject* args) {
    char* sGuid;
    HWND hwnd;
    PyArg_ParseTuple(args, "sl", &sGuid, &hwnd);

    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

    if (desktop == nullptr) {
        desktop->Release();
        return Py_BuildValue("l", -1);
    }

    HRESULT res;
    GUID guid;

    desktop->GetID(&guid);

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
    char* sGuid;
    PyArg_ParseTuple(args, "s", &sGuid);

    IVirtualDesktop* desktop = _GetDesktopFromStringId(sGuid);

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

void _PostMessageToListener(PyObject* args) {
    if (notificationCallback != nullptr) {

        PyGILState_STATE gstate;
        gstate = PyGILState_Ensure();

        PyObject * pInstance = PyObject_CallObject(notificationCallback, args);


        PyGILState_Release(gstate);
    }
}


class _Notifications : public IVirtualDesktopNotification {
private:
    ULONG _referenceCount;
public:
    // Inherited via IVirtualDesktopNotification
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void ** ppvObject) override {
        // Always set out parameter to NULL, validating it first.
        if (!ppvObject) {
            return E_INVALIDARG;
        }

        *ppvObject = NULL;

        if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification) {
            // Increment the reference count and return the pointer.
            *ppvObject = (LPVOID)this;
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    virtual ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&_referenceCount);
    }

    virtual ULONG STDMETHODCALLTYPE Release() override {
        ULONG result = InterlockedDecrement(&_referenceCount);
        if (result == 0) {
            delete this;
        }
        return 0;
    }

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopCreated(IVirtualDesktop * pDesktop) override
    {
        GUID guid;
        pDesktop->GetID(&guid);
        char sGuid[39] = {""};
        _ConvertGuidToCString(guid, sGuid);

        PyObject* args = Py_BuildValue("is", VIRTUAL_DESKTOP_CREATED, sGuid);
        _PostMessageToListener(args);
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyBegin(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback) override
    {
        GUID destroyedGuid;
        GUID fallbackGuid;

        pDesktopDestroyed->GetID(&destroyedGuid);
        pDesktopFallback->GetID(&fallbackGuid);

        char sDestroyedGuid[39] = {""};
        _ConvertGuidToCString(destroyedGuid, sDestroyedGuid);

        char sFallbackGuid[39] = {""};
        _ConvertGuidToCString(fallbackGuid, sFallbackGuid);

        PyObject* args = Py_BuildValue("iss", VIRTUAL_DESKTOP_DESTROY_BEGIN, destroyedGuid, fallbackGuid);
        _PostMessageToListener(args);
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyFailed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
    {
        GUID destroyedGuid;
        GUID fallbackGuid;

        pDesktopDestroyed->GetID(&destroyedGuid);
        pDesktopFallback->GetID(&fallbackGuid);

        char sDestroyedGuid[39] = {""};
        _ConvertGuidToCString(destroyedGuid, sDestroyedGuid);

        char sFallbackGuid[39] = {""};
        _ConvertGuidToCString(fallbackGuid, sFallbackGuid);

        PyObject* args = Py_BuildValue("iss", VIRTUAL_DESKTOP_DESTROY_FAILED, destroyedGuid, fallbackGuid);
        _PostMessageToListener(args);
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
    {
        GUID destroyedGuid;
        GUID fallbackGuid;

        pDesktopDestroyed->GetID(&destroyedGuid);
        pDesktopFallback->GetID(&fallbackGuid);

        char sDestroyedGuid[39] = {""};
        _ConvertGuidToCString(destroyedGuid, sDestroyedGuid);

        char sFallbackGuid[39] = {""};
        _ConvertGuidToCString(fallbackGuid, sFallbackGuid);

        PyObject* args = Py_BuildValue("iss", VIRTUAL_DESKTOP_DESTROYED, destroyedGuid, fallbackGuid);
        _PostMessageToListener(args);
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ViewVirtualDesktopChanged(IApplicationView * pView) override
    {
        GUID guid;
        pView->GetVirtualDesktopId(&guid);

        HWND thumbnailHwnd;
        pView->GetThumbnailWindow(&thumbnailHwnd);

        char sGuid[39] = {""};
        _ConvertGuidToCString(guid, sGuid);

        PyObject* args = Py_BuildValue("isi", VIRTUAL_DESKTOP_VIEW_CHANGED, sGuid, thumbnailHwnd);
        _PostMessageToListener(args);
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE CurrentVirtualDesktopChanged(IVirtualDesktop *pDesktopOld, IVirtualDesktop *pDesktopNew) override
    {

        viewCollection->RefreshCollection();
        GUID oldGuid = {0};
        GUID newGuid = {0};

        pDesktopOld->GetID(&oldGuid);
        pDesktopNew->GetID(&newGuid);

        char sOldGuid[39] = {""};
        _ConvertGuidToCString(oldGuid, sOldGuid);

        char sNewGuid[39] = {""};
        _ConvertGuidToCString(newGuid, sNewGuid);

        PyObject* args = Py_BuildValue("iss", VIRTUAL_DESKTOP_CURRENT_CHANGED, sOldGuid, sNewGuid);
        _PostMessageToListener(args);
        return S_OK;
    }
};


static PyObject *RegisterDesktopNotifications(PyObject *self, PyObject *args) {

    if (!registeredForNotifications) {
        _Notifications *nf = new _Notifications();

        HRESULT res = pDesktopNotificationService->Register(nf, &idNotificationService);

        if (!SUCCEEDED(res)) {
            return Py_BuildValue("i", -1);
         }

        PyArg_Parse(args, "O", &notificationCallback);
        Py_INCREF(notificationCallback);
        registeredForNotifications = TRUE;

        return Py_BuildValue("i", res);
    }
    return Py_BuildValue("i", -1);
}


static PyObject *UnregisterDesktopNotifications(PyObject *self, PyObject *args) {
    if (registeredForNotifications) {
        pDesktopNotificationService->Unregister(idNotificationService);

        registeredForNotifications = FALSE;
        Py_DECREF(notificationCallback);
        notificationCallback = nullptr;
        return Py_BuildValue("i", 1);
    }

    return Py_BuildValue("i", 0);
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
    {"RegisterDesktopNotifications", (PyCFunction)RegisterDesktopNotifications, METH_VARARGS, NULL},
    {"UnregisterDesktopNotifications", (PyCFunction)UnregisterDesktopNotifications, METH_NOARGS, NULL},
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

    if (!PyEval_ThreadsInitialized()) {
       PyEval_InitThreads();
    }

#if PY_MAJOR_VERSION >= 3
    return module;
#endif
}

VOID _OpenDllWindow(HINSTANCE injModule) {
}

