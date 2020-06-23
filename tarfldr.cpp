#include "tarfldr.h"

#include <initguid.h>
#include <commoncontrols.h>
#include <ntquery.h>
#include <shlobj.h>
#include <strsafe.h>
#include <span>

using namespace std;

static const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

static const header_info headers[] = { // FIXME - move strings to resource file
    { u"Name", LVCFMT_LEFT, 15, &FMTID_Storage, PID_STG_NAME },
    { u"Size", LVCFMT_RIGHT, 10, &FMTID_Storage, PID_STG_SIZE },
    { u"Type", LVCFMT_LEFT, 10, &FMTID_Storage, PID_STG_STORAGETYPE },
    { u"Modified", LVCFMT_LEFT, 12, &FMTID_Storage, PID_STG_WRITETIME },
};

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

#define OPEN_VERBA "Open"
#define OPEN_VERBW u"Open"

#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + " stub"s).c_str()); return E_NOTIMPL;

// FIXME - installer

template<typename... Args>
static void debug(const string_view& s, Args&&... args) { // FIXME - only if compiled in Debug mode
    string msg;

    msg = fmt::format(s, forward<Args>(args)...);

    OutputDebugStringA(msg.c_str());
}

tar_info::tar_info(const std::filesystem::path& fn) : root("", true) {
    // FIXME

    root.children.emplace_back("hello.txt", false);
    root.children.emplace_back("world.png", false);
    root.children.emplace_back("dir", true);
    root.children.back().children.emplace_back("a.txt", false);
    root.children.back().children.emplace_back("subdir", true);
    root.children.back().children.back().children.emplace_back("b.txt", false);
}

extern "C" STDAPI DllCanUnloadNow(void) {
    return objs_loaded == 0 ? S_OK : S_FALSE;
}

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

    // loop and compare case-insensitively

    tar_item* r = &tar->root;
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

    size_t pidl_length = 0;

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

    if (pdwAttributes && found_list.size() == parts.size())
        *pdwAttributes &= found_list.back()->get_atts();

    return S_OK;
}

HRESULT shell_folder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    shell_enum* se = new shell_enum(tar, grfFlags);

    return se->QueryInterface(IID_IEnumIDList, (void**)ppenumIDList);
}

HRESULT shell_folder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv) {
    if (riid == IID_IShellFolder) {
        // FIXME - parse pidl and check valid

        auto sf = new shell_folder(tar); // FIXME - parse pidl

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

    for (auto& it : tar->root.children) {
        if (it.name == sv)
            return it;
    }

    throw invalid_argument("");
}

HRESULT shell_folder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF* rgfInOut) {
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

            const auto& item = get_item_from_pidl_child(apidl[0]);

            auto scm = new shell_context_menu; // FIXME - constructor

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
    span sp = headers;

    if (iColumn >= sp.size())
        return E_INVALIDARG;

    pscid->fmtid = *headers[iColumn].fmtid;
    pscid->pid = headers[iColumn].pid;

    return S_OK;
}

HRESULT shell_folder::GetClassID(CLSID* pClassID) {
    if (!pClassID)
        return E_POINTER;

    *pClassID = CLSID_TarFolder;

    return S_OK;
}

HRESULT shell_folder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    return InitializeEx(nullptr, pidl, nullptr);
}

HRESULT shell_folder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    *ppidl = ILCloneFull(root_pidl);

    return S_OK;
}

HRESULT shell_folder::InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti) {
    root_pidl = ILCloneFull(pidlRoot);

    return S_OK;
}

HRESULT shell_folder::GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::SetMode(FOLDER_ENUM_MODE feMode) {
    folder_enum_mode = feMode;

    return S_OK;
}

HRESULT shell_folder::GetMode(FOLDER_ENUM_MODE* pfeMode) {
    if (!pfeMode)
        return E_POINTER;

    *pfeMode = folder_enum_mode;

    return S_OK;
}

HRESULT shell_enum::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IEnumIDList)
        *ppv = static_cast<IEnumIDList*>(this);
    else {
        debug("shell_enum::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_enum::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_enum::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

ITEMID_CHILD* tar_item::make_pidl_child() const {
    auto item = (ITEMIDLIST*)CoTaskMemAlloc(offsetof(ITEMIDLIST, mkid.abID) + name.length());

    if (!item)
        throw bad_alloc();

    item->mkid.cb = offsetof(ITEMIDLIST, mkid.abID) + name.length();
    memcpy(item->mkid.abID, name.data(), name.length());

    return item;
}

void tar_item::find_child(const std::u16string_view& name, tar_item** ret) {
    u16string n{name};

    for (auto& c : children) {
        auto cn = utf8_to_utf16(c.name);

        if (!_wcsicmp((wchar_t*)n.c_str(), (wchar_t*)cn.c_str())) {
            *ret = &c;
            return;
        }
    }

    *ret = nullptr;
}

SFGAOF tar_item::get_atts() const {
    SFGAOF atts;

    if (dir) {
        atts = SFGAO_FOLDER | SFGAO_BROWSABLE;
        atts |= SFGAO_HASSUBFOLDER; // FIXME - check for this
    } else
        atts = SFGAO_STREAM;

    // FIXME - SFGAO_CANRENAME, SFGAO_CANDELETE, SFGAO_HIDDEN, etc.

    return atts;
}

HRESULT shell_enum::Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) {
    try {
        if (pceltFetched)
            *pceltFetched = 0;

        // FIXME - only show folders or non-folders as requested

        while (celt > 0 && index < tar->root.children.size()) {
            *rgelt = tar->root.children[index].make_pidl_child();

            celt--;
            index++;

            if (pceltFetched)
                *pceltFetched++;
        }

        return celt == 0 ? S_OK : S_FALSE;
    } catch (const bad_alloc&) {
        return E_OUTOFMEMORY;
    }
}

HRESULT shell_enum::Skip(ULONG celt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_enum::Reset() {
    index = 0;

    return S_OK;
}

HRESULT shell_enum::Clone(IEnumIDList** ppenum) {
    UNIMPLEMENTED; // FIXME
}

shell_folder::~shell_folder() {
    if (root_pidl)
        CoTaskMemFree(root_pidl);
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
        MessageBoxW(pici->hwnd, L"FIXME - open file", L"Error", MB_ICONERROR); // FIXME
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

class factory : public IClassFactory {
public:
    factory() {
        InterlockedIncrement(&objs_loaded);
    }

    virtual ~factory() {
        InterlockedDecrement(&objs_loaded);
    }

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv) {
        if (iid == IID_IUnknown || iid == IID_IClassFactory)
            *ppv = static_cast<IClassFactory*>(this);
        else {
            *ppv = nullptr;
            return E_NOINTERFACE;
        }

        reinterpret_cast<IUnknown*>(*ppv)->AddRef();

        return S_OK;
    }

    ULONG __stdcall AddRef() {
        return InterlockedIncrement(&refcount);
    }

    ULONG __stdcall Release() {
        LONG rc = InterlockedDecrement(&refcount);

        if (rc == 0)
            delete this;

        return rc;
    }

    // IClassFactory

    HRESULT __stdcall CreateInstance(IUnknown* pUnknownOuter, const IID& iid, void** ppv) {
        if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2) {
            shell_folder* sf = new shell_folder("C:\\test.tar"); // FIXME
            if (!sf)
                return E_OUTOFMEMORY;

            return sf->QueryInterface(iid, ppv);
        }

        debug("factor::CreateInstance: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    HRESULT __stdcall LockServer(BOOL bLock) {
        return E_NOTIMPL;
    }

private:
    LONG refcount = 0;
};

extern "C" STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    if (rclsid == CLSID_TarFolder) {
        factory* fact = new factory;
        if (!fact)
            return E_OUTOFMEMORY;
        else
            return fact->QueryInterface(riid, ppv);
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

extern "C" BOOL APIENTRY DllMain(HANDLE hModule, DWORD dwReason, void* lpReserved) {
    if (dwReason == DLL_PROCESS_ATTACH)
        instance = (HINSTANCE)hModule;

    return true;
}
