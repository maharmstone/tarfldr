#include "tarfldr.h"
#include "resource.h"
#include <strsafe.h>

using namespace std;

#define VERB_EXTRACTA "extract"
#define VERB_EXTRACTW u"extract"

HRESULT shell_context_menu::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
    else if (iid == IID_IShellExtInit)
        *ppv = static_cast<IShellExtInit*>(this);
    else {
        debug("shell_context_menu::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_context_menu::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_context_menu::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_context_menu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
    UINT cmd = idCmdFirst;
    MENUITEMINFOW mii;

    debug("shell_context_menu::QueryContextMenu({}, {}, {}, {}, {:#x})", (void*)hmenu, indexMenu, idCmdFirst,
          idCmdLast, uFlags);

    if (!(uFlags & CMF_DEFAULTONLY)) {
        WCHAR buf[256];

        mii.cbSize = sizeof(mii);
        mii.wID = cmd;

        if (LoadStringW(instance, IDS_EXTRACT_ALL, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
            return E_FAIL;

        mii.fMask = MIIM_FTYPE | MIIM_ID | MIIM_STRING;
        mii.fType = MFT_STRING;
        mii.dwTypeData = buf;

        InsertMenuItemW(hmenu, indexMenu, true, &mii);
        cmd++;
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

void shell_context_menu::extract_all(HWND hwnd) {
    MessageBoxA(hwnd, "FIXME - extract all", 0, 0); // FIXME
}

HRESULT shell_context_menu::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_context_menu::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    if (IS_INTRESOURCE(pici->lpVerb)) {
        if ((uintptr_t)pici->lpVerb == 0) {
            extract_all(pici->hwnd);
            return S_OK;
        }
    } else {
        if (!strcmp(pici->lpVerb, VERB_EXTRACTA)) {
            extract_all(pici->hwnd);
            return S_OK;
        }
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax) {
    debug("shell_context_menu::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    if (idCmd > 0)
        return E_INVALIDARG;

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            return StringCchCopyA(pszName, cchMax, VERB_EXTRACTA);

        case GCS_VERBW:
            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)VERB_EXTRACTW);
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID) {
    FORMATETC format = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    HRESULT hr;
    UINT num_files;
    HDROP hdrop;
    STGMEDIUM stgm;
    WCHAR path[MAX_PATH];

    if (pidlFolder || !pdtobj)
        return E_INVALIDARG;

    stgm.tymed = TYMED_HGLOBAL;

    hr = pdtobj->GetData(&format, &stgm);
    if (FAILED(hr))
        return hr;

    hdrop = (HDROP)GlobalLock(stgm.hGlobal);

    if (!hdrop) {
        ReleaseStgMedium(&stgm);
        return E_INVALIDARG;
    }

    num_files = DragQueryFileW((HDROP)stgm.hGlobal, 0xFFFFFFFF, nullptr, 0);

    for (unsigned int i = 0; i < num_files; i++) {
        if (DragQueryFileW((HDROP)stgm.hGlobal, i, path, sizeof(path) / sizeof(WCHAR)))
            files.emplace_back(path);
    }

    GlobalUnlock(stgm.hGlobal);
    ReleaseStgMedium(&stgm);

    return S_OK;
}
