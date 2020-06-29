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
        for (unsigned int i = 0; i < items.size(); i++) {
            WCHAR buf[256];

            mii.cbSize = sizeof(mii);
            mii.wID = cmd;

            if (LoadStringW(instance, items[i].res_num, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
                return E_FAIL;

            mii.fMask = MIIM_FTYPE | MIIM_ID | MIIM_STRING;
            mii.fType = MFT_STRING;
            mii.dwTypeData = buf;

            InsertMenuItemW(hmenu, indexMenu + i, true, &mii);
            cmd++;
        }
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

void shell_context_menu::extract_all(CMINVOKECOMMANDINFO* pici) {
    BROWSEINFOW bi;
    WCHAR msg[256], buf[MAX_PATH], path[MAX_PATH];

    // FIXME - allow extracting to arbitrary PIDL, not just FS?

    if (LoadStringW(instance, IDS_EXTRACT_TEXT, msg, sizeof(msg) / sizeof(WCHAR)) <= 0)
        return;

    bi.hwndOwner = pici->hwnd;
    bi.pidlRoot = nullptr;
    bi.pszDisplayName = buf;
    bi.lpszTitle = msg;
    bi.ulFlags = BIF_RETURNONLYFSDIRS;
    bi.lpfn = nullptr;

    auto dest_pidl = SHBrowseForFolderW(&bi);

    if (!dest_pidl)
        return;

    if (!SHGetPathFromIDListW(dest_pidl, path)) {
        debug("SHGetPathFromIDListW failed");
        ILFree(dest_pidl);
        return;
    }

    ILFree(dest_pidl);

    // FIXME - do extraction
    MessageBoxW(pici->hwnd, path, L"FIXME", 0); // FIXME
}

HRESULT shell_context_menu::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_context_menu::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    if (IS_INTRESOURCE(pici->lpVerb)) {
        if ((uintptr_t)pici->lpVerb >= items.size())
            return E_INVALIDARG;

        items[(uintptr_t)pici->lpVerb].cmd(this, pici);

        return S_OK;
    }

    for (const auto& item : items) {
        if (item.verba == pici->lpVerb) {
            item.cmd(this, pici);

            return S_OK;
        }
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax) {
    debug("shell_context_menu::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    if (!IS_INTRESOURCE(idCmd)) {
        bool found = false;

        for (unsigned int i = 0; i < items.size(); i++) {
            if (items[i].verba == (char*)idCmd) {
                idCmd = i;
                found = true;
                break;
            }
        }

        if (!found)
            return E_INVALIDARG;
    } else {
        if (idCmd >= items.size())
            return E_INVALIDARG;
    }

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            return StringCchCopyA(pszName, cchMax, items[idCmd].verba.c_str());

        case GCS_VERBW:
            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)items[idCmd].verbw.c_str());
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

    items.emplace_back(IDS_EXTRACT_ALL, "extract", u"extract", shell_context_menu::extract_all);

    return S_OK;
}
