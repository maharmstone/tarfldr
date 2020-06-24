#include "tarfldr.h"
#include <strsafe.h>
#include <shlobj.h>

#define OPEN_VERBA "Open"
#define OPEN_VERBW u"Open"

using namespace std;

shell_context_menu::shell_context_menu(PIDLIST_ABSOLUTE pidl, bool is_dir) : is_dir(is_dir) {
    this->pidl = ILCloneFull(pidl);
}

shell_context_menu::~shell_context_menu() {
    if (pidl)
        ILFree(pidl);
}

HRESULT shell_context_menu::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
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

HRESULT shell_context_menu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst,
                                             UINT idCmdLast, UINT uFlags) {
    UINT cmd = idCmdFirst;
    MENUITEMINFOW mii;

    debug("shell_context_menu::QueryContextMenu({}, {}, {}, {}, {:#x})", (void*)hmenu, indexMenu, idCmdFirst,
          idCmdLast, uFlags);

    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_FTYPE | MIIM_STATE | MIIM_ID | MIIM_STRING;
    mii.fType = MFT_STRING;
    mii.fState = MFS_DEFAULT;
    mii.wID = cmd;
    mii.dwTypeData = L"&Open"; // FIXME - get from resource file

    // FIXME - others: Extract, Cut, Copy, Paste, Properties (if CMF_DEFAULTONLY not set)

    InsertMenuItemW(hmenu, indexMenu, true, &mii);

    cmd++;

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

HRESULT shell_context_menu::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_context_menu::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((int)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    if ((IS_INTRESOURCE(pici->lpVerb) && pici->lpVerb == 0) || (!IS_INTRESOURCE(pici->lpVerb) && !strcmp(pici->lpVerb, OPEN_VERBA))) {
        if (is_dir) {
            SHELLEXECUTEINFOW sei;

            memset(&sei, 0, sizeof(sei));
            sei.cbSize = sizeof(sei);
            sei.fMask = SEE_MASK_IDLIST | SEE_MASK_CLASSNAME;
            sei.lpIDList = pidl;
            sei.lpClass = L"Folder";
            sei.hwnd = pici->hwnd;
            sei.nShow = SW_SHOWNORMAL;
            sei.lpVerb = L"open";
            ShellExecuteExW(&sei);
        } else {
            // FIXME - extract and open using normal handler
            MessageBoxW(pici->hwnd, L"FIXME - open file", L"Error", MB_ICONERROR); // FIXME
        }

        return S_OK;
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved,
                                             CHAR* pszName, UINT cchMax) {
    debug("shell_context_menu::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    if (idCmd != 0)
        return E_INVALIDARG;

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            return StringCchCopyA(pszName, cchMax, OPEN_VERBA);

        case GCS_VERBW:
            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)OPEN_VERBW);
    }

    return E_INVALIDARG;
}