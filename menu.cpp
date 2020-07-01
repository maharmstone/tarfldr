#include "tarfldr.h"
#include "resource.h"
#include <strsafe.h>

using namespace std;

#define HIDA_GetPIDLFolder(pida) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[0])
#define HIDA_GetPIDLItem(pida, i) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[i+1])

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
    try {
        HRESULT hr;
        BROWSEINFOW bi;
        WCHAR msg[256], buf[MAX_PATH];
        com_object<IShellItem> dest;
        com_object<IFileOperation> ifo;

        // FIXME - can we preserve LXSS metadata when extracting?

        if (LoadStringW(instance, IDS_EXTRACT_TEXT, msg, sizeof(msg) / sizeof(WCHAR)) <= 0)
            throw last_error("LoadString", GetLastError());

        bi.hwndOwner = pici->hwnd;
        bi.pidlRoot = nullptr;
        bi.pszDisplayName = buf;
        bi.lpszTitle = msg;
        bi.ulFlags = 0;
        bi.lpfn = nullptr;

        auto dest_pidl = SHBrowseForFolderW(&bi);

        if (!dest_pidl)
            return;

        {
            IShellItem* si;

            hr = SHCreateItemFromIDList(dest_pidl, IID_IShellItem, (void**)&si);
            if (FAILED(hr)) {
                ILFree(dest_pidl);
                throw formatted_error("SHCreateItemFromIDList returned {:08x}.", (uint32_t)hr);
            }

            dest.reset(si);
        }

        ILFree(dest_pidl);

        {
            IFileOperation* fo;

            hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_ALL, IID_IFileOperation, (void**)&fo);
            if (FAILED(hr))
                throw formatted_error("CoCreateInstance returned {:08x} for CLSID_FileOperation.", (uint32_t)hr);

            ifo.reset(fo);
        }

        vector<shell_item> shell_items;

        for (const auto& file : files) {
            if ((int)get<1>(file) & (int)archive_type::tarball) {
                vector<tar_item*> itemlist;
                WCHAR path[MAX_PATH];

                if (!SHGetPathFromIDListW((ITEMIDLIST*)get<0>(file).data(), path))
                    throw runtime_error("SHGetPathFromIDList failed");

                shared_ptr<tar_info> ti{new tar_info(path)};

                for (auto& item : ti->root.children) {
                    itemlist.push_back(&item);
                }

                shell_items.emplace_back((ITEMIDLIST*)get<0>(file).data(), ti, itemlist, &ti->root, false);
            }
        }

        for (auto& si : shell_items) {
            IUnknown* unk;

            hr = si.QueryInterface(IID_IUnknown, (void**)&unk);
            if (FAILED(hr))
                throw formatted_error("shell_item::QueryInterface returned {:08x}.", (uint32_t)hr);

            hr = ifo->CopyItems(unk, dest.get());
            if (FAILED(hr))
                throw formatted_error("IFileOperation::CopyItems returned {:08x}.", (uint32_t)hr);
        }

        hr = ifo->PerformOperations();
        if (FAILED(hr))
            throw formatted_error("IFileOperation::PerformOperations returned {:08x}.", (uint32_t)hr);
    } catch (const exception& e) {
        MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
    }
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
    CLIPFORMAT cf = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
    FORMATETC format = { cf, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    HRESULT hr;
    UINT num_files;
    CIDA* cida;
    STGMEDIUM stgm;
    WCHAR path[MAX_PATH];
    bool show_extract_all = false;

    if (pidlFolder || !pdtobj)
        return E_INVALIDARG;

    stgm.tymed = TYMED_HGLOBAL;

    hr = pdtobj->GetData(&format, &stgm);
    if (FAILED(hr))
        return hr;

    cida = (CIDA*)GlobalLock(stgm.hGlobal);

    if (!cida) {
        ReleaseStgMedium(&stgm);
        return E_INVALIDARG;
    }

    for (unsigned int i = 0; i < cida->cidl; i++) {
        auto pidl = ILCombine(HIDA_GetPIDLFolder(cida), HIDA_GetPIDLItem(cida, i));

        files.emplace_back(string_view((char*)pidl, ILGetSize(pidl)), archive_type::unknown);

        ILFree(pidl);
    }

    GlobalUnlock(stgm.hGlobal);
    ReleaseStgMedium(&stgm);

    // get names of files

    vector<const ITEMIDLIST*> pidls;
    pidls.reserve(files.size());

    for (const auto& f : files) {
        pidls.push_back((ITEMIDLIST*)get<0>(f).data());
    }

    com_object<IShellItemArray> isia;

    {
        IShellItemArray* tmp;

        hr = SHCreateShellItemArrayFromIDLists(pidls.size(), pidls.data(), &tmp);
        if (FAILED(hr))
            return hr;

        isia.reset(tmp);
    }

    for (unsigned int i = 0; i < files.size(); i++) {
        com_object<IShellItem> isi;
        WCHAR* buf;

        {
            IShellItem* tmp;
            hr = isia->GetItemAt(i, &tmp);
            if (FAILED(hr))
                return hr;

            isi.reset(tmp);
        }

        hr = isi->GetDisplayName(SIGDN_NORMALDISPLAY, &buf);
        if (FAILED(hr))
            return hr;

        get<1>(files[i]) = identify_file_type((char16_t*)buf);

        if ((int)get<1>(files[i]) & (int)archive_type::tarball)
            show_extract_all = true;

        CoTaskMemFree(buf);
    }

    if (show_extract_all)
        items.emplace_back(IDS_EXTRACT_ALL, "extract", u"extract", shell_context_menu::extract_all);

    return S_OK;
}
