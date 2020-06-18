#include "tarfldr.h"

using namespace std;

static const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };
static const WCHAR class_name[] = L"tarfldr_shellview";

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + " stub"s).c_str()); return E_NOTIMPL;

// FIXME - installer

extern "C" STDAPI DllCanUnloadNow(void) {
    return objs_loaded == 0 ? S_OK : S_FALSE;
}

shell_view::~shell_view() {
    if (shell_browser)
        shell_browser->Release();
}

HRESULT shell_view::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellView || iid == IID_IShellView2)
        *ppv = static_cast<IShellView2*>(this);
    else {
        string msg = fmt::format("shell_view::QueryInterface: unsupported interface {}", iid);

        OutputDebugStringA(msg.c_str());

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_view::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_view::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_view::GetWindow(HWND *phwnd) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::ContextSensitiveHelp(WINBOOL fEnterMode) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::TranslateAccelerator(MSG *pmsg) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::EnableModeless(WINBOOL fEnable) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::UIActivate(UINT uState) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::Refresh() {
    UNIMPLEMENTED; // FIXME
}

LRESULT shell_view::wndproc(UINT uMessage, WPARAM wParam, LPARAM lParam) {
    // FIXME

    return DefWindowProcW(wnd, uMessage, wParam, lParam);
}

static LRESULT CALLBACK shell_view_wndproc_stub(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam) {
    if (uMessage == WM_NCCREATE) {
        auto lpcs = (LPCREATESTRUCTW)lParam;
        auto sv = (shell_view*)lpcs->lpCreateParams;

        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (ULONG_PTR)sv);

        return DefWindowProcW(hWnd, uMessage, wParam, lParam);
    }

    auto sv = (shell_view*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);

    return sv->wndproc(uMessage, wParam, lParam);
}

HRESULT shell_view::CreateViewWindow(IShellView* psvPrevious, LPCFOLDERSETTINGS pfs, IShellBrowser* psb,
                                     RECT* prcView, HWND* phWnd) {
    try {
        INITCOMMONCONTROLSEX icex;
        WNDCLASSW wndclass;

        if (!psb || wnd)
            return E_UNEXPECTED;

        if (shell_browser)
            shell_browser->Release();

        shell_browser = psb;

        if (shell_browser)
            shell_browser->AddRef();

        icex.dwSize = sizeof(icex);
        icex.dwICC = ICC_LISTVIEW_CLASSES;
        InitCommonControlsEx(&icex);

        *phWnd = nullptr;

        // FIXME - store view mode and flags

        shell_browser->GetWindow(&wnd_parent);

        // FIXME - get IID_ICommDlgBrowser of shell_browser

        if (!GetClassInfoW(instance, class_name, &wndclass)) {
            wndclass.style = CS_HREDRAW | CS_VREDRAW;
            wndclass.lpfnWndProc = shell_view_wndproc_stub;
            wndclass.cbClsExtra = 0;
            wndclass.cbWndExtra = 0;
            wndclass.hInstance = instance;
            wndclass.hIcon = 0;
            wndclass.hCursor = LoadCursorW(0, (LPWSTR)IDC_ARROW);
            wndclass.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
            wndclass.lpszMenuName = nullptr;
            wndclass.lpszClassName = class_name;

            if (!RegisterClassW(&wndclass))
                throw last_error("RegisterClassW", GetLastError());
        }

        wnd = CreateWindowExW(0, class_name, nullptr, WS_CHILD | WS_TABSTOP,
                              prcView->left, prcView->top,
                              prcView->right - prcView->left,
                              prcView->bottom - prcView->top,
                              wnd_parent, 0, instance, this);

        // FIXME - fiddle about with toolbar

        if (!wnd) {
            auto le = GetLastError();

            shell_browser->Release();
            shell_browser = nullptr;

            throw last_error("CreateWindowExW", le);
        }

        SetWindowPos(wnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        UpdateWindow(wnd);

        *phWnd = wnd;
    } catch (const exception& e) {
        auto msg = utf8_to_utf16(e.what());

        MessageBoxW(0, (WCHAR*)msg.c_str(), L"Error", MB_ICONERROR);

        return E_FAIL;
    }

    return S_OK;
}

HRESULT shell_view::DestroyViewWindow() {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetCurrentInfo(LPFOLDERSETTINGS pfs) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::AddPropertySheetPages(DWORD dwReserved, LPFNSVADDPROPSHEETPAGE pfn, LPARAM lparam) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SaveViewState() {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SelectItem(PCUITEMID_CHILD pidlItem, SVSIF uFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetItemObject(UINT uItem, REFIID riid, void** ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetView(SHELLVIEWID *pvid, ULONG uView) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::CreateViewWindow2(LPSV2CVW2_PARAMS lpParams) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::HandleRename(PCUITEMID_CHILD pidlNew) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SelectAndPositionItem(PCUITEMID_CHILD pidlItem, UINT uFlags, POINT *ppt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder)
        *ppv = static_cast<IShellFolder*>(this);
    else if (iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3)
        *ppv = static_cast<IPersistFolder3*>(this);
    else if (iid == IID_IPersistIDList)
        *ppv = static_cast<IPersistIDList*>(this);
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

HRESULT shell_folder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList **ppenumIDList) {
    UNIMPLEMENTED; // FIXME
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
    if (riid == IID_IUnknown || riid == IID_IShellView) {
        shell_view* sv = new shell_view;
        if (!sv)
            return E_OUTOFMEMORY;

        return sv->QueryInterface(riid, ppv);
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
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti) {
    item_id_buf.resize(offsetof(SHITEMID, abID) + pidlRoot->mkid.cb);
    memcpy(item_id_buf.data(), &pidlRoot->mkid, item_id_buf.size());

    item_id = (SHITEMID*)item_id_buf.data();

    return S_OK;
}

HRESULT shell_folder::GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::SetIDList(PCIDLIST_ABSOLUTE pidl) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::GetIDList(PIDLIST_ABSOLUTE* ppidl) {
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
