#include <windows.h>
#include <shobjidl.h>

static const GUID CLSID_TarFldr = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

LONG objs_loaded = 0;

// FIXME - installer

extern "C" STDAPI DllCanUnloadNow(void) {
    return objs_loaded == 0 ? S_OK : S_FALSE;
}

class shell_folder : public IShellFolder {
public:
    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv) {
        if (iid == IID_IUnknown || iid == IID_IShellFolder)
            *ppv = static_cast<IShellFolder*>(this);
        else {
            __asm("int $3");
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

    // IShellFolder

    HRESULT __stdcall ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten,
                                       PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList **ppenumIDList) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                    UINT *rgfReserved, void **ppv) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

    HRESULT __stdcall SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags,
                                PITEMID_CHILD *ppidlOut) {
        // FIXME
        __asm("int $3");
        return E_NOTIMPL;
    }

private:
    LONG refcount = 0;
};

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
