#include "tarfldr.h"

#include <initguid.h>
#include <commoncontrols.h>
#include <shlobj.h>
#include <span>

using namespace std;

static const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

static const header_info headers[] = { // FIXME - move strings to resource file
    { u"Name", LVCFMT_RIGHT, 15 },
    { u"Size", LVCFMT_RIGHT, 10 },
    { u"Type", LVCFMT_RIGHT, 10 },
    { u"Modified", LVCFMT_RIGHT, 12 },
};

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + " stub"s).c_str()); return E_NOTIMPL;

// FIXME - installer

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
        string msg = fmt::format("shell_folder::QueryInterface: unsupported interface {}", iid);

        OutputDebugStringA(msg.c_str());

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
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    shell_enum* se = new shell_enum(items, grfFlags);

    return se->QueryInterface(IID_IEnumIDList, (void**)ppenumIDList);
}

HRESULT shell_folder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
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

HRESULT shell_folder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                    UINT *rgfReserved, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName) {
    UNIMPLEMENTED; // FIXME
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

    __asm("int $3");
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
        string msg = fmt::format("shell_enum::QueryInterface: unsupported interface {}", iid);

        OutputDebugStringA(msg.c_str());

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

HRESULT shell_enum::Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) {
    if (pceltFetched)
        *pceltFetched = 0;

    // FIXME - only show folders or non-folders as requested

    while (celt > 0 && index < items.size()) {
        auto item = (ITEMIDLIST*)CoTaskMemAlloc(offsetof(ITEMIDLIST, mkid.abID) + items[index].name.length());

        if (!item)
            return E_OUTOFMEMORY;

        *rgelt = item;

        item->mkid.cb = items[index].name.length();
        memcpy(item->mkid.abID, items[index].name.data(), items[index].name.length());

        celt--;
        index++;

        if (pceltFetched)
            *pceltFetched++;
    }

    return celt == 0 ? S_OK : S_FALSE;
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

shell_folder::shell_folder() {
    // FIXME

    items.emplace_back("hello.txt");
    items.emplace_back("world.png");
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
            shell_folder* sf = new shell_folder;
            if (!sf)
                return E_OUTOFMEMORY;

            return sf->QueryInterface(iid, ppv);
        }

        string msg = fmt::format("factor::CreateInstance: unsupported interface {}", iid);

        OutputDebugStringA(msg.c_str());

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
