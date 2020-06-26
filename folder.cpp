#include "tarfldr.h"
#include <shlobj.h>
#include <ntquery.h>
#include <span>

using namespace std;

static const header_info headers[] = { // FIXME - move strings to resource file
    { u"Name", LVCFMT_LEFT, 15, &FMTID_Storage, PID_STG_NAME },
    { u"Size", LVCFMT_RIGHT, 10, &FMTID_Storage, PID_STG_SIZE },
    { u"Type", LVCFMT_LEFT, 10, &FMTID_Storage, PID_STG_STORAGETYPE },
    { u"Modified", LVCFMT_LEFT, 12, &FMTID_Storage, PID_STG_WRITETIME },
};

HRESULT shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2)
        *ppv = static_cast<IShellFolder2*>(this);
    else if (iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3)
        *ppv = static_cast<IPersistFolder3*>(this);
    else if (iid == IID_IObjectWithFolderEnumMode)
        *ppv = static_cast<IObjectWithFolderEnumMode*>(this);
    else {
        if (iid == IID_IExplorerPaneVisibility)
            debug("shell_folder::QueryInterface: unsupported interface IID_IExplorerPaneVisibility");
        else if (iid == IID_IPersistIDList)
            debug("shell_folder::QueryInterface: unsupported interface IID_IPersistIDList");
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
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv) {
    debug("shell_folder::CreateViewObject({}, {}, {})", (void*)hwndOwner, riid, (void*)ppv);

    if (riid == IID_IShellView) {
        SFV_CREATE sfvc;

        sfvc.cbSize = sizeof(sfvc);
        sfvc.pshf = static_cast<IShellFolder*>(this);
        sfvc.psvOuter = nullptr;
        sfvc.psfvcb = nullptr;

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
    } else if (riid == IID_IContextMenu) {
        try {
            if (cidl != 1)
                return E_INVALIDARG;

            auto pidl = ILCombine(root_pidl, apidl[0]);

            const auto& item = get_item_from_pidl_child(apidl[0]);

            auto scm = new shell_context_menu(pidl, item.dir, item.full_path, tar);

            ILFree(pidl);

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

HRESULT shell_folder::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID *pscid, VARIANT *pv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd) {
    debug("shell_folder::GetDetailsOf({}, {}, {})", (void*)pidl, iColumn, (void*)psd);

    if (!pidl) {
        span sp = headers;

        if (iColumn >= sp.size())
            return E_INVALIDARG;

        psd->fmt = sp[iColumn].fmt;
        psd->cxChar = sp[iColumn].cxChar;
        psd->str.uType = STRRET_WSTR;
        psd->str.pOleStr = (WCHAR*)CoTaskMemAlloc((sp[iColumn].name.length() + 1) * sizeof(char16_t));
        memcpy(psd->str.pOleStr, sp[iColumn].name.data(), (sp[iColumn].name.length() + 1) * sizeof(char16_t));

        return S_OK;
    }

    UNIMPLEMENTED; // FIXME
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

shell_folder::shell_folder(const std::shared_ptr<tar_info>& tar, tar_item* root, PCIDLIST_ABSOLUTE pidl) : tar(tar), root(root) {
    root_pidl = ILCloneFull(pidl);
}

shell_folder::~shell_folder() {
    if (root_pidl)
        CoTaskMemFree(root_pidl);
}
