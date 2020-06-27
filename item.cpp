#include "tarfldr.h"
#include <strsafe.h>
#include <shlobj.h>

#define OPEN_VERBA "Open"
#define OPEN_VERBW u"Open"

using namespace std;

shell_item::shell_item(PIDLIST_ABSOLUTE pidl, bool is_dir, const string_view& full_path,
                       const std::shared_ptr<tar_info>& tar) : is_dir(is_dir), full_path(full_path), tar(tar) {
    this->pidl = ILCloneFull(pidl);
}

shell_item::~shell_item() {
    if (pidl)
        ILFree(pidl);
}

HRESULT shell_item::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
    else if (iid == IID_IDataObject)
        *ppv = static_cast<IDataObject*>(this);
    else {
        debug("shell_item::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_item::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_item::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_item::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst,
                                             UINT idCmdLast, UINT uFlags) {
    UINT cmd = idCmdFirst;
    MENUITEMINFOW mii;

    debug("shell_item::QueryContextMenu({}, {}, {}, {}, {:#x})", (void*)hmenu, indexMenu, idCmdFirst,
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

static filesystem::path get_temp_file_name(const filesystem::path& dir, const u16string& prefix, unsigned int unique) {
    WCHAR tmpfn[MAX_PATH];

    if (GetTempFileNameW((WCHAR*)dir.u16string().c_str(), (WCHAR*)prefix.c_str(), unique, tmpfn) == 0)
        throw runtime_error("GetTempFileName failed.");

    return tmpfn;
}

HRESULT shell_item::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_item::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
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
            try {
                WCHAR temp_path[MAX_PATH];

                if (GetTempPathW(sizeof(temp_path) / sizeof(WCHAR), temp_path) == 0)
                    throw last_error("GetTempPath", GetLastError());

                filesystem::path fn = get_temp_file_name(temp_path, u"tar", 0);

                // replace extension with original one

                auto st = full_path.rfind(".");
                if (st != string::npos) {
                    string_view ext = string_view(full_path).substr(st + 1);
                    fn.replace_extension(ext);
                }

                tar->extract_file(full_path, fn);

                // open using normal handler

                auto ret = (intptr_t)ShellExecuteW(pici->hwnd, L"open", (WCHAR*)fn.u16string().c_str(), nullptr,
                                                   nullptr, SW_SHOW);
                if (ret <= 32)
                    throw formatted_error("ShellExecute returned {}.", ret);
            } catch (const exception& e) {
                MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
                return E_FAIL;
            }
        }

        return S_OK;
    }

    return E_INVALIDARG;
}

HRESULT shell_item::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved,
                                             CHAR* pszName, UINT cchMax) {
    debug("shell_item::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
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

HRESULT shell_item::GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium) {
    if (!pformatetcIn || !pmedium)
        return E_INVALIDARG;

    if (pformatetcIn->cfFormat != CF_HDROP || pformatetcIn->tymed != TYMED_HGLOBAL)
        return E_INVALIDARG;

    debug("shell_item::GetData(pformatetcIn = [cfFormat = {}, ptd = {}, dwAspect = {}, lindex = {}, tymed = {})], pmedium = [tymed = {}, hGlobal = {}])",
          pformatetcIn->cfFormat, (void*)pformatetcIn->ptd, pformatetcIn->dwAspect,
          pformatetcIn->lindex, pformatetcIn->tymed, pmedium->tymed, pmedium->hGlobal);

    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::QueryGetData(FORMATETC* pformatetc) {
    if (!pformatetc)
        return E_INVALIDARG;

    debug("shell_item::QueryGetData(cfFormat = {}, ptd = {}, dwAspect = {}, lindex = {}, tymed = {})",
          pformatetc->cfFormat, (void*)pformatetc->ptd, pformatetc->dwAspect, pformatetc->lindex,
          pformatetc->tymed);

    if (pformatetc->cfFormat != CF_HDROP || pformatetc->tymed != TYMED_HGLOBAL)
        return DV_E_TYMED;

    return S_OK;
}

HRESULT shell_item::GetCanonicalFormatEtc(FORMATETC* pformatectIn, FORMATETC* pformatetcOut) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, WINBOOL fRelease) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC* *ppenumFormatEtc) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::DAdvise(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink, DWORD* pdwConnection) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::DUnadvise(DWORD dwConnection) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::EnumDAdvise(IEnumSTATDATA* *ppenumAdvise) {
    UNIMPLEMENTED; // FIXME
}
