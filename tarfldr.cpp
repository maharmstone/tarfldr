#include "tarfldr.h"

using namespace std;

const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

// FIXME - installer

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

extern "C" HRESULT DllRegisterServer() {
    // FIXME

    return E_NOTIMPL;
}

extern "C" HRESULT DllUnregisterServer() {
    // FIXME

    return E_NOTIMPL;
}

ITEMID_CHILD* tar_item::make_pidl_child() const {
    auto item = (ITEMIDLIST*)CoTaskMemAlloc(offsetof(ITEMIDLIST, mkid.abID) + name.length() + offsetof(ITEMIDLIST, mkid.abID));

    if (!item)
        throw bad_alloc();

    item->mkid.cb = offsetof(ITEMIDLIST, mkid.abID) + name.length();
    memcpy(item->mkid.abID, name.data(), name.length());

    auto nextitem = (ITEMIDLIST*)((uint8_t*)item + item->mkid.cb);
    nextitem->mkid.cb = 0;

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
