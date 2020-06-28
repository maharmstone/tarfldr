#include "tarfldr.h"
#include "resource.h"
#include <shlobj.h>
#include <ntquery.h>
#include <span>

using namespace std;

const GUID FMTID_POSIXAttributes = { 0x34379fdd, 0x93be, 0x4490, { 0x98, 0x58, 0x04, 0x5d, 0x2c, 0x12, 0x85, 0xac } };

#define PID_POSIX_USER      2
#define PID_POSIX_GROUP     3
#define PID_POSIX_MODE      4

static const header_info headers[] = {
    { IDS_NAME, LVCFMT_LEFT, 15, &FMTID_Storage, PID_STG_NAME },
    { IDS_SIZE, LVCFMT_RIGHT, 10, &FMTID_Storage, PID_STG_SIZE },
    { IDS_TYPE, LVCFMT_LEFT, 10, &FMTID_Storage, PID_STG_STORAGETYPE },
    { IDS_MODIFIED, LVCFMT_LEFT, 12, &FMTID_Storage, PID_STG_WRITETIME },
    { IDS_USER, LVCFMT_LEFT, 8, &FMTID_POSIXAttributes, PID_POSIX_USER },
    { IDS_GROUP, LVCFMT_LEFT, 8, &FMTID_POSIXAttributes, PID_POSIX_GROUP },
    { IDS_MODE, LVCFMT_LEFT, 8, &FMTID_POSIXAttributes, PID_POSIX_MODE },
};

HRESULT shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2)
        *ppv = static_cast<IShellFolder2*>(this);
    else if (iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3)
        *ppv = static_cast<IPersistFolder3*>(this);
    else if (iid == IID_IObjectWithFolderEnumMode)
        *ppv = static_cast<IObjectWithFolderEnumMode*>(this);
    else if (iid == IID_IShellFolderViewCB)
        *ppv = static_cast<IShellFolderViewCB*>(this);
    else if (iid == IID_IExplorerPaneVisibility)
        *ppv = static_cast<IExplorerPaneVisibility*>(this);
    else {
        if (iid == IID_IPersistIDList)
            debug("shell_folder::QueryInterface: unsupported interface IID_IPersistIDList");
        else if (iid == IID_IFolderFilter)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderFilter");
        else if (iid == IID_IShellItem)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellItem");
        else if (iid == IID_IParentAndItem)
            debug("shell_folder::QueryInterface: unsupported interface IID_IParentAndItem");
        else if (iid == IID_IFolderView)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderView");
        else if (iid == IID_IFolderViewSettings)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderViewSettings");
        else if (iid == IID_IObjectWithSite)
            debug("shell_folder::QueryInterface: unsupported interface IID_IObjectWithSite");
        else if (iid == IID_IInternetSecurityManager)
            debug("shell_folder::QueryInterface: unsupported interface IID_IInternetSecurityManager");
        else if (iid == IID_IShellIcon)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellIcon");
        else if (iid == IID_IShellIconOverlay)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellIconOverlay");
        else
            debug("shell_folder::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_folder::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_folder::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_folder::ParseDisplayName(HWND hwnd, IBindCtx* pbc, LPWSTR pszDisplayName, ULONG* pchEaten,
                                       PIDLIST_RELATIVE* ppidl, ULONG* pdwAttributes) {
    if (!pszDisplayName || !ppidl)
        return E_INVALIDARG;

    debug("shell_folder::ParseDisplayName({}, {}, {}, {}, {}, {})", (void*)hwnd, (void*)pbc, utf16_to_utf8((char16_t*)pszDisplayName),
                                                                    (void*)pchEaten, (void*)ppidl, (void*)pdwAttributes);

    // split by backslash

    vector<u16string_view> parts;

    {
        u16string_view left = (char16_t*)pszDisplayName;

        do {
            bool found = false;

            for (unsigned int i = 0; i < left.size(); i++) {
                if (left[i] == '\\') {
                    if (i != 0)
                        parts.emplace_back(left.data(), i);

                    left = left.substr(i + 1);
                    found = true;
                    break;
                }
            }

            if (!found) {
                parts.emplace_back(left);
                break;
            }
        } while (true);
    }

    // remove trailing backslashes

    while (!parts.empty() && parts.back().empty()) {
        parts.pop_back();
    }

    // loop and compare case-insensitively

    tar_item* r = root;
    vector<tar_item*> found_list;

    if (pchEaten)
        *pchEaten = 0;

    for (const auto& p : parts) {
        tar_item* c;

        r->find_child(p, &c);

        if (!c)
            break;

        found_list.emplace_back(c);
        r = c;

        if (pchEaten) {
            *pchEaten = (p.data() + p.length()) - (const char16_t*)pszDisplayName;

            while (pszDisplayName[*pchEaten] == '\\') {
                (*pchEaten)++;
            }
        }
    }

    if (found_list.empty())
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);

    // create PIDL

    size_t pidl_length = offsetof(ITEMIDLIST, mkid.abID);

    for (auto fl : found_list) {
        pidl_length += offsetof(ITEMIDLIST, mkid.abID) + fl->name.length();
    }

    auto pidl = (ITEMIDLIST*)CoTaskMemAlloc(pidl_length);

    if (!pidl)
        return E_OUTOFMEMORY;

    *ppidl = pidl;

    for (auto fl : found_list) {
        pidl->mkid.cb = offsetof(ITEMIDLIST, mkid.abID) + fl->name.length();
        memcpy(pidl->mkid.abID, fl->name.data(), fl->name.length());
        pidl = (ITEMIDLIST*)((uint8_t*)pidl + pidl->mkid.cb);
    }

    pidl->mkid.cb = 0;

    if (pdwAttributes && found_list.size() == parts.size())
        *pdwAttributes &= found_list.back()->get_atts();

    return S_OK;
}

HRESULT shell_folder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    debug("shell_folder::EnumObjects({}, {}, {})", (void*)hwnd, grfFlags, (void*)ppenumIDList);

    shell_enum* se = new shell_enum(tar, root, grfFlags);

    return se->QueryInterface(IID_IEnumIDList, (void**)ppenumIDList);
}

HRESULT shell_folder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv) {
    debug("shell_folder::BindToObject({}, {}, {}, {})", (void*)pidl, (void*)pbc, riid, (void*)ppv);

    if (riid == IID_IShellFolder) {
        tar_item* item = root;

        if (pidl) {
            const SHITEMID* sh = &pidl->mkid;

            while (sh->cb != 0) {
                string_view name{(char*)sh->abID, sh->cb - offsetof(SHITEMID, abID)};
                bool found = false;

                for (auto& it : item->children) {
                    if (it.name == name) {
                        found = true;
                        item = &it;
                        break;
                    }
                }

                if (!found)
                    return E_NOINTERFACE;

                sh = (SHITEMID*)((uint8_t*)sh + sh->cb);
            }
        }

        if (!item->dir)
            return E_NOINTERFACE;

        auto new_pidl = ILCombine(root_pidl, pidl);

        auto sf = new shell_folder(tar, item, new_pidl);

        ILFree(new_pidl);

        return sf->QueryInterface(riid, ppv);
    }

    if (riid == IID_IPropertyStoreFactory)
        debug("shell_folder::BindToObject: unsupported interface IID_IPropertyStoreFactory");
    else if (riid == IID_IPropertyStore)
        debug("shell_folder::BindToObject: unsupported interface IID_IPropertyStore");
    else
        debug("shell_folder::BindToObject: unsupported interface {}", riid);

    *ppv = nullptr;
    return E_NOINTERFACE;
}

HRESULT shell_folder::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
    debug("shell_folder::CompareIDs({}, {}, {})", lParam, (void*)pidl1, (void*)pidl2);

    if (!pidl1 || !pidl2)
        return E_INVALIDARG;

    try {
        uint16_t col = lParam & SHCIDS_COLUMNMASK;
        int res;

        tar_item& item1 = get_item_from_pidl_child(pidl1);
        tar_item& item2 = get_item_from_pidl_child(pidl2);

        switch (col) {
            case 0: { // name
                auto val = item1.name.compare(item2.name);

                if (val < 0)
                    res = -1;
                else if (val > 0)
                    res = 1;
                else
                    res = 0;

                break;
            }

            case 1: { // size
                int64_t size1 = item1.dir ? 0 : item1.size;
                int64_t size2 = item2.dir ? 0 : item2.size;

                if (size1 < size2)
                    res = -1;
                else if (size2 < size1)
                    res = 1;
                else
                    res = 0;

                break;
            }

            case 2: { // type
                SHFILEINFOW sfi;
                u16string type1, type2;

                if (!SHGetFileInfoW((LPCWSTR)utf8_to_utf16(item1.name).c_str(),
                                    item1.dir ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL,
                                    &sfi, sizeof(sfi),
                                    SHGFI_USEFILEATTRIBUTES | SHGFI_TYPENAME)) {
                    throw last_error("SHGetFileInfo", GetLastError());
                }

                type1 = (char16_t*)sfi.szTypeName;

                if (!SHGetFileInfoW((LPCWSTR)utf8_to_utf16(item2.name).c_str(),
                                    item2.dir ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL,
                                    &sfi, sizeof(sfi),
                                    SHGFI_USEFILEATTRIBUTES | SHGFI_TYPENAME)) {
                    throw last_error("SHGetFileInfo", GetLastError());
                }

                type2 = (char16_t*)sfi.szTypeName;

                auto val = type1.compare(type2);

                if (val < 0)
                    res = -1;
                else if (val > 0)
                    res = 1;
                else
                    res = 0;

                break;
            }

            case 3: { // date
                if (!item1.mtime.has_value() && !item2.mtime.has_value())
                    res = 0;
                else if (!item1.mtime.has_value())
                    res = -1;
                else if (!item2.mtime.has_value())
                    res = 1;
                else if (item1.mtime.value() < item2.mtime.value())
                    res = -1;
                else if (item2.mtime.value() < item1.mtime.value())
                    res = 1;
                else
                    res = 0;

                break;
            }

            default:
                return E_INVALIDARG;
        }

        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, res == -1 ? 0xffff : res);
    } catch (const invalid_argument&) {
        return E_INVALIDARG;
    } catch (...) {
        return E_FAIL;
    }
}

HRESULT shell_folder::CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv) {
    debug("shell_folder::CreateViewObject({}, {}, {})", (void*)hwndOwner, riid, (void*)ppv);

    if (riid == IID_IShellView) {
        SFV_CREATE sfvc;

        sfvc.cbSize = sizeof(sfvc);
        sfvc.pshf = static_cast<IShellFolder*>(this);
        sfvc.psvOuter = nullptr;

        QueryInterface(IID_IShellFolderViewCB, (void**)&sfvc.psfvcb);

        return SHCreateShellFolderView(&sfvc, (IShellView**)ppv);
    }

    *ppv = nullptr;
    return E_NOINTERFACE;
}

tar_item& shell_folder::get_item_from_pidl_child(const ITEMID_CHILD* pidl) {
    if (pidl->mkid.cb < offsetof(ITEMIDLIST, mkid.abID))
        throw invalid_argument("");

    string_view sv{(char*)pidl->mkid.abID, pidl->mkid.cb - offsetof(ITEMIDLIST, mkid.abID)};

    for (auto& it : root->children) {
        if (it.name == sv)
            return it;
    }

    throw invalid_argument("");
}

HRESULT shell_folder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF* rgfInOut) {
    debug("shell_folder::GetAttributesOf({}, {}, {})", cidl, (void*)apidl, (void*)rgfInOut);

    try {
        SFGAOF common_atts = *rgfInOut;

        while (cidl > 0) {
            SFGAOF atts;

            const auto& item = get_item_from_pidl_child(apidl[0]);

            atts = item.get_atts();

            common_atts &= atts;

            cidl--;
            apidl++;
        }

        *rgfInOut = common_atts;

        return S_OK;
    } catch (const invalid_argument&) {
        return E_INVALIDARG;
    }
}

HRESULT shell_folder::GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                    UINT* rgfReserved, void** ppv) {
    debug("shell_folder::GetUIObjectOf({}, {}, {}, {}, {}, {})", (void*)hwndOwner, cidl, (void*)apidl, riid,
          (void*)rgfReserved, (void*)ppv);

    if (riid == IID_IExtractIconW || riid == IID_IExtractIconA) {
        try {
            if (cidl != 1)
                return E_INVALIDARG;

            const auto& item = get_item_from_pidl_child(apidl[0]);

            return SHCreateFileExtractIconW((LPCWSTR)utf8_to_utf16(item.name).c_str(),
                                            item.dir ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL,
                                            riid, ppv);
        } catch (const invalid_argument&) {
            return E_INVALIDARG;
        }
    } else if (riid == IID_IContextMenu || riid == IID_IDataObject) {
        try {
            if (cidl == 0)
                return E_INVALIDARG;

            vector<tar_item*> itemlist;

            itemlist.reserve(cidl);

            for (unsigned int i = 0; i < cidl; i++) {
                itemlist.emplace_back(&get_item_from_pidl_child(apidl[i]));
            }

            auto scm = new shell_item(root_pidl, tar, itemlist);

            return scm->QueryInterface(riid, ppv);
        } catch (const invalid_argument&) {
            return E_INVALIDARG;
        }
    }

    debug("shell_folder::GetUIObjectOf: unsupported interface {}", riid);

    *ppv = nullptr;
    return E_NOINTERFACE;
}

HRESULT shell_folder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName) {
    debug("shell_folder::GetDisplayNameOf({}, {}, {})", (void*)pidl, uFlags, (void*)pName);

    try {
        const auto& item = get_item_from_pidl_child(pidl);

        auto u16name = utf8_to_utf16(item.name);

        pName->uType = STRRET_WSTR;
        pName->pOleStr = (WCHAR*)CoTaskMemAlloc((u16name.length() + 1) * sizeof(char16_t));
        memcpy(pName->pOleStr, u16name.c_str(), (u16name.length() + 1) * sizeof(char16_t));

        return S_OK;
    } catch (const invalid_argument&) {
        return E_INVALIDARG;
    }
}

HRESULT shell_folder::SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags,
                                PITEMID_CHILD *ppidlOut) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDefaultSearchGUID(GUID *pguid) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::EnumSearches(IEnumExtraSearch **ppenum) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDefaultColumn(DWORD dwRes, ULONG *pSort, ULONG *pDisplay) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF *pcsFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID* pscid, VARIANT* pv) {
    debug("shell_folder::GetDetailsEx({}, {} (fmtid = {}, pid = {}), {}", (void*)pidl,
          (void*)pscid, pscid->fmtid, pscid->pid, (void*)pv);

    if (pscid->fmtid != FMTID_Storage && pscid->fmtid != FMTID_POSIXAttributes)
        return E_NOTIMPL;

    try {
        tar_item& item = get_item_from_pidl_child(pidl);

        if (pscid->fmtid == FMTID_Storage) {
            switch (pscid->pid) {
                case PID_STG_NAME:
                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString((WCHAR*)utf8_to_utf16(item.name).c_str());

                    return S_OK;

                case PID_STG_SIZE:
                    if (item.dir)
                        return S_FALSE;

                    pv->vt = VT_I8;
                    pv->llVal = item.size;

                    return S_OK;

                case PID_STG_STORAGETYPE: {
                    SHFILEINFOW sfi;

                    if (!SHGetFileInfoW((LPCWSTR)utf8_to_utf16(item.name).c_str(),
                                        item.dir ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL,
                                        &sfi, sizeof(sfi),
                                        SHGFI_USEFILEATTRIBUTES | SHGFI_TYPENAME)) {
                        throw last_error("SHGetFileInfo", GetLastError());
                    }

                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString(sfi.szTypeName);

                    return S_OK;
                }

                case PID_STG_WRITETIME: {
                    if (!item.mtime.has_value())
                        return S_FALSE;

                    pv->vt = VT_DATE;
                    pv->date = ((double)item.mtime.value() / 86400.0) + 25569.0;

                    return S_OK;
                }

                default:
                    return E_NOTIMPL;
            }
        } else if (pscid->fmtid == FMTID_POSIXAttributes) {
            switch (pscid->pid) {
                case PID_POSIX_USER:
                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString((WCHAR*)utf8_to_utf16(item.user).c_str());

                    return S_OK;

                case PID_POSIX_GROUP:
                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString((WCHAR*)utf8_to_utf16(item.group).c_str());

                    return S_OK;

                // FIXME - PID_POSIX_MODE

                default:
                    return E_NOTIMPL;
            }
        }
    } catch (const invalid_argument&) {
        return E_INVALIDARG;
    } catch (...) {
        return E_FAIL;
    }

    return E_FAIL;
}

HRESULT shell_folder::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd) {
    HRESULT hr;
    SHCOLUMNID col;
    VARIANT v;

    debug("shell_folder::GetDetailsOf({}, {}, {})", (void*)pidl, iColumn, (void*)psd);

    span sp = headers;

    if (iColumn >= sp.size())
        return E_INVALIDARG;

    if (!pidl) {
        WCHAR buf[256];

        if (LoadStringW(instance, sp[iColumn].res_num, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
            return E_FAIL;

        psd->fmt = sp[iColumn].fmt;
        psd->cxChar = sp[iColumn].cxChar;
        psd->str.uType = STRRET_WSTR;
        psd->str.pOleStr = (WCHAR*)CoTaskMemAlloc((wcslen(buf) + 1) * sizeof(WCHAR));
        memcpy(psd->str.pOleStr, buf, (wcslen(buf) + 1) * sizeof(WCHAR));

        return S_OK;
    }

    col.fmtid = *sp[iColumn].fmtid;
    col.pid = sp[iColumn].pid;

    VariantInit(&v);

    hr = GetDetailsEx(pidl, &col, &v);

    if (FAILED(hr)) {
        VariantClear(&v);
        return hr;
    }

    hr = VariantChangeType(&v, &v, 0, VT_BSTR);

    if (FAILED(hr)) {
        VariantClear(&v);
        return hr;
    }

    psd->fmt = sp[iColumn].fmt;
    psd->cxChar = sp[iColumn].cxChar;
    psd->str.uType = STRRET_WSTR;
    psd->str.pOleStr = (WCHAR*)CoTaskMemAlloc((wcslen(v.bstrVal) + 1) * sizeof(WCHAR));
    memcpy(psd->str.pOleStr, v.bstrVal, (wcslen(v.bstrVal) + 1) * sizeof(WCHAR));

    VariantClear(&v);

    return S_OK;
}

HRESULT shell_folder::MapColumnToSCID(UINT iColumn, SHCOLUMNID* pscid) {
    debug("shell_folder::MapColumnToSCID({}, {})", iColumn, (void*)pscid);

    span sp = headers;

    if (iColumn >= sp.size())
        return E_INVALIDARG;

    pscid->fmtid = *headers[iColumn].fmtid;
    pscid->pid = headers[iColumn].pid;

    return S_OK;
}

HRESULT shell_folder::GetClassID(CLSID* pClassID) {
    debug("shell_folder::GetClassID({})", (void*)pClassID);

    if (!pClassID)
        return E_POINTER;

    *pClassID = CLSID_TarFolder;

    return S_OK;
}

HRESULT shell_folder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    return InitializeEx(nullptr, pidl, nullptr);
}

HRESULT shell_folder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    debug("shell_folder::GetCurFolder({})", (void*)ppidl);

    *ppidl = ILCloneFull(root_pidl);

    return S_OK;
}

HRESULT shell_folder::InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti) {
    HRESULT hr;
    WCHAR path[MAX_PATH];

    debug("shell_folder::InitializeEx({}, {},{})", (void*)pbc, (void*)pidlRoot, (void*)ppfti);

    if (ppfti) {
        debug("shell_folder::InitializeEx: ppfti (pidlTargetFolder = {}, szTargetParsingName = {}, szNetworkProvider = {}, dwAttributes = {}, csidl = {})",
              (void*)ppfti->pidlTargetFolder, utf16_to_utf8((char16_t*)ppfti->szTargetParsingName),
              utf16_to_utf8((char16_t*)ppfti->szNetworkProvider), ppfti->dwAttributes, ppfti->csidl);
    }

    if (!SHGetPathFromIDListW(pidlRoot, path)) {
        debug("SHGetPathFromIDListW failed");
        return E_FAIL;
    }

    try {
        tar.reset(new tar_info(path));
    } catch (const exception& e) {
        debug("shell_folder::InitializeEx: {}", e.what());
        return E_FAIL;
    }

    root = &tar->root;

    root_pidl = ILCloneFull(pidlRoot);

    return S_OK;
}

HRESULT shell_folder::GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::SetMode(FOLDER_ENUM_MODE feMode) {
    debug("shell_folder::SetMode({})", feMode);

    folder_enum_mode = feMode;

    return S_OK;
}

HRESULT shell_folder::GetMode(FOLDER_ENUM_MODE* pfeMode) {
    debug("shell_folder::GetMode({})", (void*)pfeMode);

    if (!pfeMode)
        return E_POINTER;

    *pfeMode = folder_enum_mode;

    return S_OK;
}

HRESULT shell_folder::MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam) {
    debug("shell_folder::MessageSFVCB({}, {}, {})", uMsg, wParam, lParam);

    UNIMPLEMENTED;
}

HRESULT shell_folder::GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE* peps) {
    debug("shell_folder::GetPaneState({}, {})", ep, (void*)peps);

    if (ep == EP_Ribbon)
        *peps = EPS_DEFAULT_ON;
    else
        *peps = EPS_DONTCARE;

    return S_OK;
}

shell_folder::shell_folder(const std::shared_ptr<tar_info>& tar, tar_item* root, PCIDLIST_ABSOLUTE pidl) : tar(tar), root(root) {
    root_pidl = ILCloneFull(pidl);
}

shell_folder::~shell_folder() {
    if (root_pidl)
        CoTaskMemFree(root_pidl);
}
