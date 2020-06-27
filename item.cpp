#include "tarfldr.h"
#include <strsafe.h>
#include <shlobj.h>
#include <span>
#include <functional>

using namespace std;

static const struct {
    char16_t* name;
    const char* verba;
    const char16_t* verbw;
    function<HRESULT(shell_item*, CMINVOKECOMMANDINFO*)> cmd;
} menu_items[] = {
    { u"&Open", "Open", u"Open", &shell_item::open_cmd },
    { u"&Copy", "Copy", u"Copy", &shell_item::copy_cmd }
};

// FIXME - others: Extract, Cut, Paste, Properties

shell_item::shell_item(PIDLIST_ABSOLUTE root_pidl, const shared_ptr<tar_info>& tar,
                       const vector<tar_item*>& itemlist) : tar(tar), itemlist(itemlist) {
    this->root_pidl = ILCloneFull(root_pidl);
    cf_shell_id_list = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
    cf_file_contents = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    cf_file_descriptor = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
}

shell_item::~shell_item() {
    if (root_pidl)
        ILFree(root_pidl);
}

HRESULT shell_item::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
    else if (iid == IID_IDataObject)
        *ppv = static_cast<IDataObject*>(this);
    else {
        if (iid == IID_IStdMarshalInfo)
            debug("shell_item::QueryInterface: unsupported interface IID_IStdMarshalInfo");
        else if (iid == IID_INoMarshal)
            debug("shell_item::QueryInterface: unsupported interface IID_INoMarshal");
        else if (iid == IID_IAgileObject)
            debug("shell_item::QueryInterface: unsupported interface IID_IAgileObject");
        else if (iid == IID_ICallFactory)
            debug("shell_item::QueryInterface: unsupported interface IID_ICallFactory");
        else if (iid == IID_IExternalConnection)
            debug("shell_item::QueryInterface: unsupported interface IID_IExternalConnection");
        else if (iid == IID_IMarshal)
            debug("shell_item::QueryInterface: unsupported interface IID_IMarshal");
        else if (iid == IID_IDataObjectAsyncCapability)
            debug("shell_item::QueryInterface: unsupported interface IID_IDataObjectAsyncCapability");
        else
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

    span mi = menu_items;

    for (unsigned int i = 0; i < mi.size(); i++) {
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_FTYPE | MIIM_STATE | MIIM_ID | MIIM_STRING;
        mii.fType = MFT_STRING;
        mii.fState = i == 0 ? MFS_DEFAULT : 0;
        mii.wID = cmd;
        mii.dwTypeData = (WCHAR*)mi[i].name; // FIXME - get from resource file

        InsertMenuItemW(hmenu, indexMenu + i, true, &mii);
        cmd++;

        if (uFlags & CMF_DEFAULTONLY)
            break;
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

static filesystem::path get_temp_file_name(const filesystem::path& dir, const u16string& prefix, unsigned int unique) {
    WCHAR tmpfn[MAX_PATH];

    if (GetTempFileNameW((WCHAR*)dir.u16string().c_str(), (WCHAR*)prefix.c_str(), unique, tmpfn) == 0)
        throw runtime_error("GetTempFileName failed.");

    return tmpfn;
}

HRESULT shell_item::open_cmd(CMINVOKECOMMANDINFO* pici) {
    for (auto item : itemlist) {
        if (item->dir) {
            SHELLEXECUTEINFOW sei;

            auto child_pidl = item->make_pidl_child();
            auto pidl = ILCombine(root_pidl, child_pidl);

            ILFree(child_pidl);

            memset(&sei, 0, sizeof(sei));
            sei.cbSize = sizeof(sei);
            sei.fMask = SEE_MASK_IDLIST | SEE_MASK_CLASSNAME;
            sei.lpIDList = pidl;
            sei.lpClass = L"Folder";
            sei.hwnd = pici->hwnd;
            sei.nShow = SW_SHOWNORMAL;
            sei.lpVerb = L"open";
            ShellExecuteExW(&sei);

            ILFree(pidl);
        } else {
            try {
                WCHAR temp_path[MAX_PATH];

                if (GetTempPathW(sizeof(temp_path) / sizeof(WCHAR), temp_path) == 0)
                    throw last_error("GetTempPath", GetLastError());

                filesystem::path fn = get_temp_file_name(temp_path, u"tar", 0);

                // replace extension with original one

                auto st = item->full_path.rfind(".");
                if (st != string::npos) {
                    string_view ext = string_view(item->full_path).substr(st + 1);
                    fn.replace_extension(ext);
                }

                tar->extract_file(item->full_path, fn);

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
    }

    return S_OK;
}

HRESULT shell_item::copy_cmd(CMINVOKECOMMANDINFO* pici) {
    HRESULT hr;
    IDataObject* dataobj;

    hr = QueryInterface(IID_IDataObject, (void**)&dataobj);
    if (FAILED(hr))
        return hr;

    OleSetClipboard(dataobj);

    dataobj->Release();

    return S_OK;
}

HRESULT shell_item::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_item::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    span mi = menu_items;

    if (IS_INTRESOURCE(pici->lpVerb)) {
        if ((int)pici->lpVerb >= mi.size())
            return E_INVALIDARG;

        return mi[(unsigned int)pici->lpVerb].cmd(this, pici);
    }

    for (const auto& mie : mi) {
        if (!strcmp(pici->lpVerb, mie.verba))
            return mie.cmd(this, pici);
    }

    return E_INVALIDARG;
}

HRESULT shell_item::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved,
                                     CHAR* pszName, UINT cchMax) {
    debug("shell_item::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    span mi = menu_items;

    if (idCmd >= mi.size())
        return E_INVALIDARG;

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            return StringCchCopyA(pszName, cchMax, mi[0].verba);

        case GCS_VERBW:
            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)mi[0].verbw);
    }

    return E_INVALIDARG;
}

HRESULT shell_item::GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium) {
    char16_t format[256];

    if (!pformatetcIn || !pmedium)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetcIn->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item::GetData(pformatetcIn = [cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})], pmedium = [tymed = {}, hGlobal = {}])",
          pformatetcIn->cfFormat, utf16_to_utf8(format), (void*)pformatetcIn->ptd, pformatetcIn->dwAspect,
          pformatetcIn->lindex, pformatetcIn->tymed, pmedium->tymed, pmedium->hGlobal);

    if (pformatetcIn->cfFormat == cf_shell_id_list && pformatetcIn->tymed == TYMED_HGLOBAL) {
        pmedium->tymed = TYMED_HGLOBAL;
        pmedium->hGlobal = make_shell_id_list();
        pmedium->pUnkForRelease = nullptr;

        return S_OK;
    }

    return E_INVALIDARG;
}

HGLOBAL shell_item::make_shell_id_list() {
    HGLOBAL hg;
    CIDA* cida;
    size_t size, root_pidl_size;
    uint8_t* ptr;
    UINT* off;

    root_pidl_size = ILGetSize(root_pidl);

    size = offsetof(CIDA, aoffset) + (sizeof(UINT) * (itemlist.size() + 1)) + root_pidl_size;

    for (auto item : itemlist) {
        auto child_pidl = item->make_pidl_child();

        size += ILGetSize(child_pidl);

        ILFree(child_pidl);
    }

    hg = GlobalAlloc(GHND | GMEM_SHARE, size);

    if (!hg)
        return nullptr;

    cida = (CIDA*)GlobalLock(hg);
    cida->cidl = itemlist.size();

    off = &cida->aoffset[0];
    ptr = (uint8_t*)cida + offsetof(CIDA, aoffset) + ((cida->cidl + 1) * sizeof(UINT));

    *off = ptr - (uint8_t*)cida;
    memcpy(ptr, &root_pidl, root_pidl_size);
    ptr += root_pidl_size;
    off++;

    for (auto item : itemlist) {
        auto child_pidl = item->make_pidl_child();
        size_t child_pidl_size = ILGetSize(child_pidl);

        *off = ptr - (uint8_t*)cida;
        memcpy(ptr, child_pidl, child_pidl_size);

        ILFree(child_pidl);

        ptr += child_pidl_size;
        off++;
    }

    GlobalUnlock(hg);

    return hg;
}

HRESULT shell_item::GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::QueryGetData(FORMATETC* pformatetc) {
    char16_t format[256];

    if (!pformatetc)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetc->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item::QueryGetData(cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})",
          pformatetc->cfFormat, utf16_to_utf8(format), (void*)pformatetc->ptd, pformatetc->dwAspect,
          pformatetc->lindex, pformatetc->tymed);

    if (pformatetc->cfFormat == cf_shell_id_list && pformatetc->tymed == TYMED_HGLOBAL)
        return S_OK;
    else if (pformatetc->cfFormat == cf_file_contents && pformatetc->tymed == TYMED_ISTREAM)
        return S_OK;
    else if (pformatetc->cfFormat == cf_file_descriptor && pformatetc->tymed == TYMED_HGLOBAL)
        return S_OK;

    return DV_E_TYMED;
}

HRESULT shell_item::GetCanonicalFormatEtc(FORMATETC* pformatectIn, FORMATETC* pformatetcOut) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, WINBOOL fRelease) {
    char16_t format[256];

    if (!pformatetc || !pmedium)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetc->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item::SetData(pformatetc = [cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})], pmedium = [tymed = {}, hGlobal = {}], fRelease = {})",
          pformatetc->cfFormat, utf16_to_utf8(format), (void*)pformatetc->ptd, pformatetc->dwAspect,
          pformatetc->lindex, pformatetc->tymed, pmedium->tymed, pmedium->hGlobal, fRelease);

    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc) {
    if (dwDirection == DATADIR_GET) {
        auto sief = new shell_item_enum_format(cf_shell_id_list, cf_file_contents, cf_file_descriptor);

        return sief->QueryInterface(IID_IEnumFORMATETC, (void**)ppenumFormatEtc);
    }

    return E_NOTIMPL;
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

shell_item_enum_format::shell_item_enum_format(CLIPFORMAT cf_shell_id_list, CLIPFORMAT cf_file_contents,
                                               CLIPFORMAT cf_file_descriptor) {
    formats.emplace_back(cf_shell_id_list, TYMED_HGLOBAL);
    formats.emplace_back(cf_file_contents, TYMED_ISTREAM);
    formats.emplace_back(cf_file_descriptor, TYMED_HGLOBAL);
}

HRESULT shell_item_enum_format::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IEnumFORMATETC)
        *ppv = static_cast<IEnumFORMATETC*>(this);
    else {
        debug("shell_item_enum_format::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_item_enum_format::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_item_enum_format::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_item_enum_format::Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched) {
    if (pceltFetched)
        *pceltFetched = 0;

    while (celt > 0 && index < formats.size()) {
        rgelt->cfFormat = formats[index].format;
        rgelt->dwAspect = DVASPECT_CONTENT;
        rgelt->ptd = nullptr;
        rgelt->tymed = formats[index].tymed;
        rgelt->lindex = -1;

        rgelt++;
        celt--;
        index++;

        if (pceltFetched)
            (*pceltFetched)++;
    }

    return celt == 0 ? S_OK : S_FALSE;
}

HRESULT shell_item_enum_format::Skip(ULONG celt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_enum_format::Reset() {
    index = 0;

    return S_OK;
}

HRESULT shell_item_enum_format::Clone(IEnumFORMATETC** ppenum) {
    UNIMPLEMENTED; // FIXME
}
