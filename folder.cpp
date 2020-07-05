/* Copyright (c) Mark Harmstone 2020
 *
 * This file is part of tarfldr.
 *
 * tarfldr is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public Licence as published by
 * the Free Software Foundation, either version 3 of the Licence, or
 * (at your option) any later version.
 *
 * tarfldr is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public Licence for more details.
 *
 * You should have received a copy of the GNU Lesser General Public Licence
 * along with tarfldr.  If not, see <http://www.gnu.org/licenses/>. */

#include "tarfldr.h"
#include "resource.h"
#include <shlobj.h>
#include <ntquery.h>
#include <span>

using namespace std;

const GUID FMTID_POSIXAttributes = { 0x34379fdd, 0x93be, 0x4490, { 0x98, 0x58, 0x04, 0x5d, 0x2c, 0x12, 0x85, 0xac } };

static int name_compare(const tar_item& item1, const tar_item& item2);
static int size_compare(const tar_item& item1, const tar_item& item2);
static int type_compare(const tar_item& item1, const tar_item& item2);
static int date_compare(const tar_item& item1, const tar_item& item2);
static int user_compare(const tar_item& item1, const tar_item& item2);
static int group_compare(const tar_item& item1, const tar_item& item2);
static int mode_compare(const tar_item& item1, const tar_item& item2);

static const header_info headers[] = {
    { IDS_NAME, LVCFMT_LEFT, 15, &FMTID_Storage, PID_STG_NAME, name_compare, SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT, false },
    { IDS_SIZE, LVCFMT_RIGHT, 10, &FMTID_Storage, PID_STG_SIZE, size_compare, SHCOLSTATE_TYPE_INT | SHCOLSTATE_ONBYDEFAULT, false },
    { IDS_TYPE, LVCFMT_LEFT, 10, &FMTID_Storage, PID_STG_STORAGETYPE, type_compare, SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT, false },
    { IDS_MODIFIED, LVCFMT_LEFT, 14, &FMTID_Storage, PID_STG_WRITETIME, date_compare, SHCOLSTATE_TYPE_DATE | SHCOLSTATE_ONBYDEFAULT, false },
    { IDS_USER, LVCFMT_LEFT, 8, &FMTID_POSIXAttributes, PID_POSIX_USER, user_compare, SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT, true },
    { IDS_GROUP, LVCFMT_LEFT, 8, &FMTID_POSIXAttributes, PID_POSIX_GROUP, group_compare, SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT, true },
    { IDS_MODE, LVCFMT_LEFT, 10, &FMTID_POSIXAttributes, PID_POSIX_MODE, mode_compare, SHCOLSTATE_TYPE_STR | SHCOLSTATE_ONBYDEFAULT, true },
};

#define __S_IFMT        0170000
#define __S_IFDIR       0040000
#define __S_ISTYPE(mode, mask)  (((mode) & __S_IFMT) == (mask))

HRESULT shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2)
        *ppv = static_cast<IShellFolder2*>(this);
    else if (iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3)
        *ppv = static_cast<IPersistFolder3*>(this);
    else if (iid == IID_IObjectWithFolderEnumMode)
        *ppv = static_cast<IObjectWithFolderEnumMode*>(this);
    else if (iid == IID_IExplorerPaneVisibility)
        *ppv = static_cast<IExplorerPaneVisibility*>(this);
    else {
        if (iid == IID_IPersistIDList)
            debug("shell_folder::QueryInterface: unsupported interface IID_IPersistIDList\n");
        else if (iid == IID_IFolderFilter)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderFilter\n");
        else if (iid == IID_IShellItem)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellItem\n");
        else if (iid == IID_IParentAndItem)
            debug("shell_folder::QueryInterface: unsupported interface IID_IParentAndItem\n");
        else if (iid == IID_IFolderView)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderView\n");
        else if (iid == IID_IFolderViewSettings)
            debug("shell_folder::QueryInterface: unsupported interface IID_IFolderViewSettings\n");
        else if (iid == IID_IObjectWithSite)
            debug("shell_folder::QueryInterface: unsupported interface IID_IObjectWithSite\n");
        else if (iid == IID_IInternetSecurityManager)
            debug("shell_folder::QueryInterface: unsupported interface IID_IInternetSecurityManager\n");
        else if (iid == IID_IShellIcon)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellIcon\n");
        else if (iid == IID_IShellIconOverlay)
            debug("shell_folder::QueryInterface: unsupported interface IID_IShellIconOverlay\n");
        else
            debug("shell_folder::QueryInterface: unsupported interface {}\n", iid);

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

    debug("shell_folder::ParseDisplayName({}, {}, {}, {}, {}, {})\n", (void*)hwnd, (void*)pbc, utf16_to_utf8((char16_t*)pszDisplayName),
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

    // load file, if not done already

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
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
    debug("shell_folder::EnumObjects({}, {}, {})\n", (void*)hwnd, grfFlags, (void*)ppenumIDList);

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
    }

    shell_enum* se = new shell_enum(tar, root, grfFlags);

    return se->QueryInterface(IID_IEnumIDList, (void**)ppenumIDList);
}

HRESULT shell_folder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv) {
    debug("shell_folder::BindToObject({}, {}, {}, {})\n", (void*)pidl, (void*)pbc, riid, (void*)ppv);

    if (riid == IID_IShellFolder) {
        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

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
    } else if (riid == IID_IStream) {
        if (!pidl)
            return E_NOINTERFACE;

        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

        tar_item* item = root;
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

        try {
            auto tis = new tar_item_stream(tar, *item);

            return tis->QueryInterface(riid, ppv);
        } catch (...) {
            return E_FAIL;
        }
    }

    if (riid == IID_IPropertyStoreFactory)
        debug("shell_folder::BindToObject: unsupported interface IID_IPropertyStoreFactory\n");
    else if (riid == IID_IPropertyStore)
        debug("shell_folder::BindToObject: unsupported interface IID_IPropertyStore\n");
    else
        debug("shell_folder::BindToObject: unsupported interface {}\n", riid);

    *ppv = nullptr;
    return E_NOINTERFACE;
}

HRESULT shell_folder::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

static int name_compare(const tar_item& item1, const tar_item& item2) {
    auto val = item1.name.compare(item2.name);

    if (val < 0)
        return -1;
    else if (val > 0)
        return 1;
    else
        return 0;
}

static int size_compare(const tar_item& item1, const tar_item& item2) {
    int64_t size1 = item1.dir ? 0 : item1.size;
    int64_t size2 = item2.dir ? 0 : item2.size;

    if (size1 < size2)
        return -1;
    else if (size2 < size1)
        return 1;
    else
        return 0;
}

static int type_compare(const tar_item& item1, const tar_item& item2) {
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
        return -1;
    else if (val > 0)
        return 1;
    else
        return 0;
}

static int date_compare(const tar_item& item1, const tar_item& item2) {
    if (!item1.mtime.has_value() && !item2.mtime.has_value())
        return 0;
    else if (!item1.mtime.has_value())
        return -1;
    else if (!item2.mtime.has_value())
        return 1;
    else if (item1.mtime.value() < item2.mtime.value())
        return -1;
    else if (item2.mtime.value() < item1.mtime.value())
        return 1;
    else
        return 0;
}

static int user_compare(const tar_item& item1, const tar_item& item2) {
    auto val = item1.user.compare(item2.user);

    if (val < 0)
        return -1;
    else if (val > 0)
        return 1;
    else
        return 0;
}

static int group_compare(const tar_item& item1, const tar_item& item2) {
    auto val = item1.group.compare(item2.group);

    if (val < 0)
        return -1;
    else if (val > 0)
        return 1;
    else
        return 0;
}

static int mode_compare(const tar_item& item1, const tar_item& item2) {
    if (item1.mode < item2.mode)
        return -1;
    else if (item2.mode < item1.mode)
        return 1;
    else
        return 0;
}

tar_item& shell_folder::get_item_from_relative_pidl(PCUIDLIST_RELATIVE pidl) {
    tar_item* r = root;

    while (pidl->mkid.cb != 0) {
        bool found = false;

        if (pidl->mkid.cb < offsetof(ITEMIDLIST, mkid.abID))
            throw invalid_argument("");

        string_view sv{(char*)pidl->mkid.abID, pidl->mkid.cb - offsetof(ITEMIDLIST, mkid.abID)};

        for (auto& it : r->children) {
            if (it.name == sv) {
                r = &it;
                found = true;
                pidl = (ITEMIDLIST*)((uint8_t*)pidl + pidl->mkid.cb);
                break;
            }
        }

        if (!found)
            throw invalid_argument("");
    }

    return *r;
}

HRESULT shell_folder::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
    debug("shell_folder::CompareIDs({}, {}, {})\n", lParam, (void*)pidl1, (void*)pidl2);

    if (!pidl1 || !pidl2)
        return E_INVALIDARG;

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
    }

    try {
        uint16_t col = lParam & SHCIDS_COLUMNMASK;
        int res;

        span h = headers;

        if (col >= h.size())
            return E_INVALIDARG;

        if (!h[col].compare_func)
            res = 0;
        else {
            tar_item& item1 = get_item_from_relative_pidl(pidl1);
            tar_item& item2 = get_item_from_relative_pidl(pidl2);

            res = h[col].compare_func(item1, item2);
        }

        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, res == -1 ? 0xffff : res);
    } catch (const invalid_argument&) {
        return E_INVALIDARG;
    } catch (...) {
        return E_FAIL;
    }
}

HRESULT shell_folder::CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv) {
    debug("shell_folder::CreateViewObject({}, {}, {})\n", (void*)hwndOwner, riid, (void*)ppv);

    if (riid == IID_IShellView) {
        SFV_CREATE sfvc;

        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

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
    debug("shell_folder::GetAttributesOf({}, {}, {})\n", cidl, (void*)apidl, (void*)rgfInOut);

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
    }

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
    debug("shell_folder::GetUIObjectOf({}, {}, {}, {}, {}, {})\n", (void*)hwndOwner, cidl, (void*)apidl, riid,
          (void*)rgfReserved, (void*)ppv);

    if (riid == IID_IExtractIconW || riid == IID_IExtractIconA) {
        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

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
        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

        try {
            if (cidl == 0)
                return E_INVALIDARG;

            vector<tar_item*> itemlist;

            itemlist.reserve(cidl);

            for (unsigned int i = 0; i < cidl; i++) {
                itemlist.emplace_back(&get_item_from_pidl_child(apidl[i]));
            }

            auto scm = new shell_item_list(root_pidl, tar, itemlist, root, true, this);

            return scm->QueryInterface(riid, ppv);
        } catch (const invalid_argument&) {
            return E_INVALIDARG;
        }
    }

    debug("shell_folder::GetUIObjectOf: unsupported interface {}\n", riid);

    *ppv = nullptr;
    return E_NOINTERFACE;
}

HRESULT shell_folder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName) {
    debug("shell_folder::GetDisplayNameOf({}, {}, {})\n", (void*)pidl, uFlags, (void*)pName);

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
    }

    try {
        const auto& item = get_item_from_pidl_child(pidl);
        u16string ret;

        // FIXME - are we supposed to hide extensions here?

        if (uFlags & SHGDN_FORPARSING && !(uFlags & SHGDN_INFOLDER)) {
            tar_item* r;

            ret = utf8_to_utf16(item.name);

            r = item.parent;

            while (r) {
                ret = utf8_to_utf16(r->name) + u"\\" + ret;
                r = r->parent;
            }

            ret = tar->archive_fn.u16string() + ret;
        } else
            ret = utf8_to_utf16(item.name);

        pName->uType = STRRET_WSTR;
        pName->pOleStr = (WCHAR*)CoTaskMemAlloc((ret.length() + 1) * sizeof(char16_t));
        memcpy(pName->pOleStr, ret.c_str(), (ret.length() + 1) * sizeof(char16_t));

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

HRESULT shell_folder::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF* pcsFlags) {
    span h = headers;

    if (iColumn >= h.size())
        return E_INVALIDARG;

    *pcsFlags = h[iColumn].state;

    if (h[iColumn].tarball_only) {
        if (!tar) {
            try {
                tar.reset(new tar_info(path));

                root = &tar->root;
            } catch (const exception& e) {
                return E_FAIL;
            }
        }

        if (!(tar->type & archive_type::tarball)) {
            *pcsFlags &= ~SHCOLSTATE_ONBYDEFAULT;
            *pcsFlags |= SHCOLSTATE_HIDDEN;
        }
    }

    return S_OK;
}

HRESULT shell_folder::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID* pscid, VARIANT* pv) {
    debug("shell_folder::GetDetailsEx({}, {} (fmtid = {}, pid = {}), {}\n", (void*)pidl,
          (void*)pscid, pscid->fmtid, pscid->pid, (void*)pv);

    if (pscid->fmtid != FMTID_Storage && pscid->fmtid != FMTID_POSIXAttributes)
        return E_NOTIMPL;

    if (!tar) {
        try {
            tar.reset(new tar_info(path));

            root = &tar->root;
        } catch (const exception& e) {
            return E_FAIL;
        }
    }

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
                    if (!(tar->type & archive_type::tarball))
                        return E_FAIL;

                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString((WCHAR*)utf8_to_utf16(item.user).c_str());

                    return S_OK;

                case PID_POSIX_GROUP:
                    if (!(tar->type & archive_type::tarball))
                        return E_FAIL;

                    pv->vt = VT_BSTR;
                    pv->bstrVal = SysAllocString((WCHAR*)utf8_to_utf16(item.group).c_str());

                    return S_OK;

                case PID_POSIX_MODE:
                    if (!(tar->type & archive_type::tarball))
                        return E_FAIL;

                    pv->vt = VT_I4;
                    pv->lVal = item.mode;

                    return S_OK;

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

u16string mode_to_u16string(mode_t m) {
    char16_t mode[11];

    mode[0] = __S_ISTYPE(m, __S_IFDIR) ? 'd' : '-';
    mode[1] = m & 0400 ? 'r' : '-';
    mode[2] = m & 0200 ? 'w' : '-';
    mode[3] = m & 0100 ? 'x' : '-';
    mode[4] = m & 0040 ? 'r' : '-';
    mode[5] = m & 0020 ? 'w' : '-';
    mode[6] = m & 0010 ? 'x' : '-';
    mode[7] = m & 0004 ? 'r' : '-';
    mode[8] = m & 0002 ? 'w' : '-';
    mode[9] = m & 0001 ? 'x' : '-';
    mode[10] = 0;

    // FIXME - sticky bits, char / block devices, links, sockets, etc.

    return mode;
}

HRESULT shell_folder::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd) {
    HRESULT hr;
    SHCOLUMNID col;
    VARIANT v;

    debug("shell_folder::GetDetailsOf({}, {}, {})\n", (void*)pidl, iColumn, (void*)psd);

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

    psd->fmt = sp[iColumn].fmt;
    psd->cxChar = sp[iColumn].cxChar;
    psd->str.uType = STRRET_WSTR;

    if (col.fmtid == FMTID_POSIXAttributes && col.pid == PID_POSIX_MODE && v.vt == VT_I4) {
        auto s = mode_to_u16string((mode_t)v.lVal);

        psd->str.pOleStr = (WCHAR*)CoTaskMemAlloc((s.length() + 1) * sizeof(char16_t));
        memcpy(psd->str.pOleStr, s.c_str(), (s.length() + 1) * sizeof(char16_t));
    } else {
        hr = VariantChangeType(&v, &v, 0, VT_BSTR);

        if (FAILED(hr)) {
            VariantClear(&v);
            return hr;
        }

        psd->str.pOleStr = (WCHAR*)CoTaskMemAlloc((wcslen(v.bstrVal) + 1) * sizeof(WCHAR));
        memcpy(psd->str.pOleStr, v.bstrVal, (wcslen(v.bstrVal) + 1) * sizeof(WCHAR));
    }

    VariantClear(&v);

    return S_OK;
}

HRESULT shell_folder::MapColumnToSCID(UINT iColumn, SHCOLUMNID* pscid) {
    debug("shell_folder::MapColumnToSCID({}, {})\n", iColumn, (void*)pscid);

    span sp = headers;

    if (iColumn >= sp.size())
        return E_INVALIDARG;

    pscid->fmtid = *headers[iColumn].fmtid;
    pscid->pid = headers[iColumn].pid;

    return S_OK;
}

HRESULT shell_folder::GetClassID(CLSID* pClassID) {
    debug("shell_folder::GetClassID({})\n", (void*)pClassID);

    if (!pClassID)
        return E_POINTER;

    *pClassID = CLSID_TarFolder;

    return S_OK;
}

HRESULT shell_folder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    return InitializeEx(nullptr, pidl, nullptr);
}

HRESULT shell_folder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    debug("shell_folder::GetCurFolder({})\n", (void*)ppidl);

    *ppidl = ILCloneFull(root_pidl);

    return S_OK;
}

HRESULT shell_folder::InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti) {
    HRESULT hr;
    WCHAR path[MAX_PATH];

    debug("shell_folder::InitializeEx({}, {},{})\n", (void*)pbc, (void*)pidlRoot, (void*)ppfti);

    if (ppfti) {
        debug("shell_folder::InitializeEx: ppfti (pidlTargetFolder = {}, szTargetParsingName = {}, szNetworkProvider = {}, dwAttributes = {}, csidl = {})\n",
              (void*)ppfti->pidlTargetFolder, utf16_to_utf8((char16_t*)ppfti->szTargetParsingName),
              utf16_to_utf8((char16_t*)ppfti->szNetworkProvider), ppfti->dwAttributes, ppfti->csidl);
    }

    if (!SHGetPathFromIDListW(pidlRoot, path)) {
        debug("SHGetPathFromIDListW failed\n");
        return E_FAIL;
    }

    this->path = path;

    tar.reset();

    if (root_pidl)
        ILFree(root_pidl);

    root_pidl = ILCloneFull(pidlRoot);

    return S_OK;
}

HRESULT shell_folder::GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::SetMode(FOLDER_ENUM_MODE feMode) {
    debug("shell_folder::SetMode({})\n", feMode);

    folder_enum_mode = feMode;

    return S_OK;
}

HRESULT shell_folder::GetMode(FOLDER_ENUM_MODE* pfeMode) {
    debug("shell_folder::GetMode({})\n", (void*)pfeMode);

    if (!pfeMode)
        return E_POINTER;

    *pfeMode = folder_enum_mode;

    return S_OK;
}

HRESULT shell_folder::GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE* peps) {
    debug("shell_folder::GetPaneState({}, {})\n", ep, (void*)peps);

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
