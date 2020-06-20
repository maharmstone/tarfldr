#include "tarfldr.h"

#include <initguid.h>
#include <commoncontrols.h>
#include <span>

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

HRESULT shell_view::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellView || iid == IID_IShellView2)
        *ppv = static_cast<IShellView2*>(this);
    else if (iid == IID_IFolderView || iid == IID_IFolderView2)
        *ppv = static_cast<IFolderView2*>(this);
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

void shell_view::on_create() {
    try {
        DWORD style, exstyle;
        HRESULT hr;
        com_object<IEnumIDList> enumidlist;
        LPITEMIDLIST pidl;

        style = WS_TABSTOP | WS_VISIBLE | WS_CHILDWINDOW | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
                LVS_SHAREIMAGELISTS | LVS_EDITLABELS | LVS_ALIGNLEFT | LVS_AUTOARRANGE;
        exstyle = WS_EX_CLIENTEDGE;

        if (view_id) {
            const auto& vid = view_id.value();

            if (vid == VID_LargeIcons)
                style |= LVS_ICON;
            else if (vid == VID_SmallIcons)
                style |= LVS_SMALLICON;
            else if (vid == VID_List)
                style |= LVS_LIST;
            else if (vid == VID_Details)
                style |= LVS_REPORT;
        } else {
            switch (view_mode) {
                case FVM_ICON:
                    style |= LVS_ICON;
                break;

                case FVM_DETAILS:
                    style |= LVS_REPORT;
                break;

                case FVM_SMALLICON:
                    style |= LVS_SMALLICON;
                break;

                case FVM_LIST:
                    style |= LVS_LIST;
                break;
            }
        }

        if (flags & FWF_AUTOARRANGE)
            style |= LVS_AUTOARRANGE;

        if (flags & FWF_DESKTOP)
            flags |= FWF_NOCLIENTEDGE | FWF_NOSCROLL;

        if (flags & FWF_SINGLESEL)
            style |= LVS_SINGLESEL;

        if (flags & FWF_NOCLIENTEDGE)
            exstyle &= ~WS_EX_CLIENTEDGE;

        wnd_list = CreateWindowExW(exstyle, WC_LISTVIEWW, nullptr, style, 0, 0, 0, 0,
                                   wnd, nullptr, instance, nullptr);

        if (!wnd_list)
            throw last_error("CreateWindowExW", GetLastError());

        // FIXME - create columns
        // FIXME - populate

        hr = SHGetImageList(SHIL_LARGE, IID_IImageList, (void**)&image_list_large);
        if (FAILED(hr))
            throw runtime_error("SHGetImageList failed (" + to_string(hr) + ")");

        hr = SHGetImageList(SHIL_SMALL, IID_IImageList, (void**)&image_list_small);
        if (FAILED(hr))
            throw runtime_error("SHGetImageList failed (" + to_string(hr) + ")");

        // FIXME - other sizes?

        SendMessageW(wnd_list, LVM_SETIMAGELIST, LVSIL_NORMAL, (LPARAM)image_list_large.get());
        SendMessageW(wnd_list, LVM_SETIMAGELIST, LVSIL_SMALL, (LPARAM)image_list_small.get());

        {
            IEnumIDList* obj;

            hr = folder->EnumObjects(wnd, SHCONTF_NONFOLDERS | SHCONTF_FOLDERS, &obj);
            if (FAILED(hr))
                throw runtime_error("EnumObjects failed (" + to_string(hr) + ")");

            enumidlist.reset(obj);
        }

        SetRedraw(false);

        while (enumidlist->Next(1, &pidl, nullptr) == S_OK) {
            LVITEMW item;

            item.mask = LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;
            item.iItem = 0;
            item.iSubItem = 0;
            item.lParam = (LPARAM)pidl;
            item.pszText = L"test"; // FIXME // (LPWSTR)f.c_str();
            item.iImage = 0; // FIXME

//             SHFILEINFOW sfi;
//             auto ret = SHGetFileInfoW((LPCWSTR)f.c_str(), FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi),
//                                       SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES);
//             if (!ret)
//                 throw last_error("SHGetFileInfo", GetLastError());
//
//             item.iImage = sfi.iIcon;

            SendMessageW(wnd_list, LVM_INSERTITEMW, 0, (LPARAM)&item);
        }

        // FIXME - send LVM_SORTITEMS

        SetRedraw(true);

//         vector<u16string> files;
//
//         files.emplace_back(u"hello.png");
//         files.emplace_back(u"world.txt");
//
//         for (const auto& f : files) {
//             LVITEMW item;
//
//             item.mask = LVIF_TEXT | LVIF_IMAGE;
//             item.iItem = 0;
//             item.iSubItem = 0;
//             item.lParam = 0;
//             item.pszText = (LPWSTR)f.c_str();
//
//             SHFILEINFOW sfi;
//             auto ret = SHGetFileInfoW((LPCWSTR)f.c_str(), FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi),
//                                       SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES);
//             if (!ret)
//                 throw last_error("SHGetFileInfo", GetLastError());
//
//             item.iImage = sfi.iIcon;
//
//             SendMessageW(wnd_list, LVM_INSERTITEMW, 0, (LPARAM)&item);
//         }
    } catch (const exception& e) {
        auto msg = utf8_to_utf16(e.what());

        MessageBoxW(0, (WCHAR*)msg.c_str(), L"Error", MB_ICONERROR);
    }
}

void shell_view::on_size(unsigned int width, unsigned int height) {
    if (!wnd_list)
        return;

    MoveWindow(wnd_list, 0, 0, width, height, true);
}

LRESULT shell_view::wndproc(UINT uMessage, WPARAM wParam, LPARAM lParam) {
    switch (uMessage) {
        case WM_CREATE:
            on_create();
            return 0;

        case WM_SIZE:
            on_size(LOWORD(lParam), HIWORD(lParam));
            return 0;
    }

    return DefWindowProcW(wnd, uMessage, wParam, lParam);
}

static LRESULT CALLBACK shell_view_wndproc_stub(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam) {
    if (uMessage == WM_NCCREATE) {
        auto lpcs = (LPCREATESTRUCTW)lParam;
        auto sv = (shell_view*)lpcs->lpCreateParams;

        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (ULONG_PTR)sv);
        sv->wnd = hWnd;

        return DefWindowProcW(hWnd, uMessage, wParam, lParam);
    }

    auto sv = (shell_view*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);

    return sv->wndproc(uMessage, wParam, lParam);
}

HRESULT shell_view::CreateViewWindow(IShellView* psvPrevious, LPCFOLDERSETTINGS pfs, IShellBrowser* psb,
                                     RECT* prcView, HWND* phWnd) {
    SV2CVW2_PARAMS params;
    HRESULT hr;

    params.cbSize = sizeof(params);
    params.psvPrev = psvPrevious;
    params.pfs = pfs;
    params.psbOwner = psb;
    params.prcView = prcView;
    params.pvid = nullptr;

    hr = CreateViewWindow2(&params);

    if (SUCCEEDED(hr))
        *phWnd = params.hwndView;

    return hr;
}

HRESULT shell_view::DestroyViewWindow() {
    if (!wnd)
        return S_OK;

    UIActivate(SVUIA_DEACTIVATE);

    DestroyWindow(wnd);

    shell_browser.reset();

    wnd = nullptr;

    return S_OK;
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
    if (uView == SV2GV_CURRENTVIEW) {
        *pvid = view_id.value();
        return S_OK;
    } else if (uView == SV2GV_DEFAULTVIEW) {
        *pvid = *supported_view_ids[0];
        return S_OK;
    }

    span sp = supported_view_ids;

    if (uView >= sp.size())
        return E_INVALIDARG;

    *pvid = *sp[uView];

    return S_OK;
}

HRESULT shell_view::CreateViewWindow2(LPSV2CVW2_PARAMS params) {
    try {
        INITCOMMONCONTROLSEX icex;
        WNDCLASSW wndclass;

        if (!params || wnd)
            return E_UNEXPECTED;

        if (params->pvid) {
            span sp = supported_view_ids;
            bool supported = false;

            for (const auto& v : sp) {
                if (*v == *params->pvid) {
                    supported = true;
                    break;
                }
            }

            if (!supported)
                return E_INVALIDARG;

            view_id = *params->pvid;
        } else
            view_id = nullopt;

        shell_browser.reset(params->psbOwner);

        if (shell_browser)
            shell_browser->AddRef();

        icex.dwSize = sizeof(icex);
        icex.dwICC = ICC_LISTVIEW_CLASSES;
        InitCommonControlsEx(&icex);

        params->hwndView = nullptr;

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

        view_mode = params->pfs->ViewMode;
        flags = params->pfs->fFlags;

        wnd = CreateWindowExW(0, class_name, nullptr, WS_CHILD | WS_TABSTOP,
                              params->prcView->left, params->prcView->top,
                              params->prcView->right - params->prcView->left,
                              params->prcView->bottom - params->prcView->top,
                              wnd_parent, 0, instance, this);

        // FIXME - fiddle about with toolbar

        if (!wnd) {
            auto le = GetLastError();

            shell_browser.reset();

            throw last_error("CreateWindowExW", le);
        }

        SetWindowPos(wnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        UpdateWindow(wnd);

        params->hwndView = wnd;
    } catch (const exception& e) {
        auto msg = utf8_to_utf16(e.what());

        MessageBoxW(0, (WCHAR*)msg.c_str(), L"Error", MB_ICONERROR);

        return E_FAIL;
    }

    return S_OK;
}

HRESULT shell_view::HandleRename(PCUITEMID_CHILD pidlNew) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SelectAndPositionItem(PCUITEMID_CHILD pidlItem, UINT uFlags, POINT *ppt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetCurrentViewMode(UINT *pViewMode) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetCurrentViewMode(UINT ViewMode) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetFolder(REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::Item(int iItemIndex, PITEMID_CHILD *ppidl) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::ItemCount(UINT uFlags, int *pcItems) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::Items(UINT uFlags, REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSelectionMarkedItem(int *piItem) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetFocusedItem(int *piItem) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetItemPosition(PCUITEMID_CHILD pidl, POINT *ppt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSpacing(POINT *ppt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetDefaultSpacing(POINT *ppt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetAutoArrange() {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SelectItem(int iItem, DWORD dwFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SelectAndPositionItems(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, POINT *apt, DWORD dwFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetGroupBy(REFPROPERTYKEY key, WINBOOL fAscending) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetGroupBy(PROPERTYKEY *pkey, WINBOOL *pfAscending) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetViewProperty(PCUITEMID_CHILD pidl, REFPROPERTYKEY propkey, REFPROPVARIANT propvar) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetViewProperty(PCUITEMID_CHILD pidl, REFPROPERTYKEY propkey, PROPVARIANT *ppropvar) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetTileViewProperties(PCUITEMID_CHILD pidl, LPCWSTR pszPropList) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetExtendedTileViewProperties(PCUITEMID_CHILD pidl, LPCWSTR pszPropList) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetText(FVTEXTTYPE iType, LPCWSTR pwszText) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetCurrentFolderFlags(DWORD dwMask, DWORD dwFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetCurrentFolderFlags(DWORD *pdwFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSortColumnCount(int *pcColumns) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetSortColumns(const SORTCOLUMN *rgSortColumns, int cColumns) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSortColumns(SORTCOLUMN *rgSortColumns, int cColumns) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetItem(int iItem, REFIID riid, void **ppv) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetVisibleItem(int iStart, WINBOOL fPrevious, int *piItem) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSelectedItem(int iStart, int *piItem) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSelection(WINBOOL fNoneImpliesFolder, IShellItemArray **ppsia) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetSelectionState(PCUITEMID_CHILD pidl, DWORD *pdwFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::InvokeVerbOnSelection(LPCSTR pszVerb) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetViewModeAndIconSize(FOLDERVIEWMODE uViewMode, int iImageSize) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetViewModeAndIconSize(FOLDERVIEWMODE *puViewMode, int *piImageSize) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetGroupSubsetCount(UINT cVisibleRows) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::GetGroupSubsetCount(UINT *pcVisibleRows) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::SetRedraw(WINBOOL fRedrawOn) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::IsMoveInSameFolder() {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_view::DoRename() {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_folder::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2)
        *ppv = static_cast<IShellFolder2*>(this);
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
    if (riid == IID_IUnknown || riid == IID_IShellView || riid == IID_IShellView2 || riid == IID_IFolderView || riid == IID_IFolderView2) {
        shell_view* sv = new shell_view(static_cast<IShellFolder*>(this));
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
        auto item = (ITEMIDLIST*)CoTaskMemAlloc(offsetof(ITEMIDLIST, mkid.abID) + (items[index].name.length() * sizeof(char16_t)));

        if (!item)
            return E_OUTOFMEMORY;

        *rgelt = item;

        item->mkid.cb = items[index].name.length() * sizeof(char16_t);
        memcpy(item->mkid.abID, items[index].name.data(), items[index].name.length() * sizeof(char16_t));

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

    items.emplace_back(u"hello.txt");
    items.emplace_back(u"world.png");
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
