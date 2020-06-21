#include "tarfldr.h"

#include <initguid.h>
#include <commoncontrols.h>
#include <shlobj.h>
#include <span>

using namespace std;

static const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

static const header_info headers[] = { // FIXME - move strings to resource file
    { u"Name", LVCFMT_LEFT, 15 },
    { u"Size", LVCFMT_RIGHT, 10 },
    { u"Type", LVCFMT_LEFT, 10 },
    { u"Modified", LVCFMT_LEFT, 12 },
};

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + " stub"s).c_str()); return E_NOTIMPL;

// FIXME - installer

template<typename... Args>
static void debug(const string_view& s, Args&&... args) {
    string msg;

    msg = fmt::format(s, forward<Args>(args)...);

    OutputDebugStringA(msg.c_str());
}

tar_info::tar_info(const std::filesystem::path& fn) {
    // FIXME

    items.emplace_back("hello.txt", 0, false);
    items.emplace_back("world.png", 1, false);
    items.emplace_back("dir", 2, true);
    items.back().children.emplace_back("a.txt", 0, false);
    items.back().children.emplace_back("subdir", 1, true);
    items.back().children.back().children.emplace_back("b.txt", 0, false);
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

HRESULT shell_folder::ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten,
                                       PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes) {
    debug("shell_folder::ParseDisplayName({}, {}, {}, {}, {}, {})", (void*)hwnd, (void*)pbc, pszDisplayName ? utf16_to_utf8((char16_t*)pszDisplayName) : "",
                                                                    (void*)pchEaten, (void*)ppidl, (void*)pdwAttributes);

    UNIMPLEMENTED; // FIXME
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

    for (auto& it : tar->items) {
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

            if (item.dir) {
                atts = SFGAO_FOLDER | SFGAO_BROWSABLE;
                atts |= SFGAO_HASSUBFOLDER; // FIXME - check for this
            } else
                atts = SFGAO_STREAM;

            // FIXME - SFGAO_CANRENAME, SFGAO_CANDELETE, SFGAO_HIDDEN, etc.

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

HRESULT shell_folder::MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid) {
    UNIMPLEMENTED; // FIXME
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

HRESULT shell_enum::Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) {
    try {
        if (pceltFetched)
            *pceltFetched = 0;

        // FIXME - only show folders or non-folders as requested

        while (celt > 0 && index < tar->items.size()) {
            *rgelt = tar->items[index].make_pidl_child();

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
