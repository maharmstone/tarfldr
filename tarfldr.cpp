#include "tarfldr.h"

using namespace std;

static const GUID CLSID_TarFldr = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

LONG objs_loaded = 0;

// FIXME - installer

extern "C" STDAPI DllCanUnloadNow(void) {
    return objs_loaded == 0 ? S_OK : S_FALSE;
}

HRESULT __stdcall shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder)
        *ppv = static_cast<IShellFolder*>(this);
    else if (iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3)
        *ppv = static_cast<IPersistFolder3*>(this);
    else if (iid == IID_IPersistIDList)
        *ppv = static_cast<IPersistIDList*>(this);
    else {
        __asm("int $3");
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG __stdcall shell_folder::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG __stdcall shell_folder::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT __stdcall shell_folder::ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten,
                                                 PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList **ppenumIDList) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                              UINT *rgfReserved, void **ppv) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags,
                                          PITEMID_CHILD *ppidlOut) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::GetClassID(CLSID* pClassID) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    return InitializeEx(nullptr, pidl, nullptr);
}

HRESULT __stdcall shell_folder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti) {
    item_id_buf.resize(offsetof(SHITEMID, abID) + pidlRoot->mkid.cb);
    memcpy(item_id_buf.data(), &pidlRoot->mkid, item_id_buf.size());

    item_id = (SHITEMID*)item_id_buf.data();

    return S_OK;
}

HRESULT __stdcall shell_folder::GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::SetIDList(PCIDLIST_ABSOLUTE pidl) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
}

HRESULT __stdcall shell_folder::GetIDList(PIDLIST_ABSOLUTE* ppidl) {
    // FIXME
    __asm("int $3");
    return E_NOTIMPL;
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
        if (iid == IID_IUnknown || iid == IID_IShellFolder) {
            shell_folder* sf = new shell_folder;
            if (!sf)
                return E_OUTOFMEMORY;

            return sf->QueryInterface(iid, ppv);
        }

        __asm("int $3");

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
    if (rclsid == CLSID_TarFldr) {
        factory* fact = new factory;
        if (!fact)
            return E_OUTOFMEMORY;
        else
            return fact->QueryInterface(riid, ppv);
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

extern "C" BOOL APIENTRY DllMain(HANDLE hModule, DWORD dwReason, void* lpReserved) {
    return true;
}
