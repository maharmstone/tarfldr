#define _WIN32_WINNT 0x0601 // Windows 7

#include <windows.h>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <stdexcept>
#include <stdint.h>
#include <shlguid.h>
#include <fmt/format.h>

template<typename T>
class com_object_closer {
public:
    typedef T* pointer;

    void operator()(T* t) {
        if (t)
            t->Release();
    }
};

template<typename T>
using com_object = std::unique_ptr<T*, com_object_closer<T>>;

static const SHELLVIEWID* supported_view_ids[] = { &VID_LargeIcons, &VID_SmallIcons, &VID_List, &VID_Details }; // FIXME - VID_Tile?

class shell_folder;

class shell_view : public IShellView2, public IFolderView2 {
public:
    shell_view(IShellFolder* parent) : folder{parent} {
        folder->AddRef();
    }

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IShellView2

    HRESULT __stdcall GetWindow(HWND *phwnd);
    HRESULT __stdcall ContextSensitiveHelp(WINBOOL fEnterMode);
    HRESULT __stdcall TranslateAccelerator(MSG *pmsg);
    HRESULT __stdcall EnableModeless(WINBOOL fEnable);
    HRESULT __stdcall UIActivate(UINT uState);
    HRESULT __stdcall Refresh();
    HRESULT __stdcall CreateViewWindow(IShellView* psvPrevious, LPCFOLDERSETTINGS pfs, IShellBrowser* psb,
                                       RECT* prcView, HWND* phWnd);
    HRESULT __stdcall DestroyViewWindow();
    HRESULT __stdcall GetCurrentInfo(LPFOLDERSETTINGS pfs);
    HRESULT __stdcall AddPropertySheetPages(DWORD dwReserved, LPFNSVADDPROPSHEETPAGE pfn, LPARAM lparam);
    HRESULT __stdcall SaveViewState();
    HRESULT __stdcall SelectItem(PCUITEMID_CHILD pidlItem, SVSIF uFlags);
    HRESULT __stdcall GetItemObject(UINT uItem, REFIID riid, void** ppv);
    HRESULT __stdcall GetView(SHELLVIEWID *pvid, ULONG uView);
    HRESULT __stdcall CreateViewWindow2(LPSV2CVW2_PARAMS lpParams);
    HRESULT __stdcall HandleRename(PCUITEMID_CHILD pidlNew);
    HRESULT __stdcall SelectAndPositionItem(PCUITEMID_CHILD pidlItem, UINT uFlags, POINT *ppt);

    // IFolderView2

    HRESULT __stdcall GetCurrentViewMode(UINT *pViewMode);
    HRESULT __stdcall SetCurrentViewMode(UINT ViewMode);
    HRESULT __stdcall GetFolder(REFIID riid, void **ppv);
    HRESULT __stdcall Item(int iItemIndex, PITEMID_CHILD *ppidl);
    HRESULT __stdcall ItemCount(UINT uFlags, int *pcItems);
    HRESULT __stdcall Items(UINT uFlags, REFIID riid, void **ppv);
    HRESULT __stdcall GetSelectionMarkedItem(int *piItem);
    HRESULT __stdcall GetFocusedItem(int *piItem);
    HRESULT __stdcall GetItemPosition(PCUITEMID_CHILD pidl, POINT *ppt);
    HRESULT __stdcall GetSpacing(POINT *ppt);
    HRESULT __stdcall GetDefaultSpacing(POINT *ppt);
    HRESULT __stdcall GetAutoArrange();
    HRESULT __stdcall SelectItem(int iItem, DWORD dwFlags);
    HRESULT __stdcall SelectAndPositionItems(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, POINT *apt, DWORD dwFlags);
    HRESULT __stdcall SetGroupBy(REFPROPERTYKEY key, WINBOOL fAscending);
    HRESULT __stdcall GetGroupBy(PROPERTYKEY *pkey, WINBOOL *pfAscending);
    HRESULT __stdcall SetViewProperty(PCUITEMID_CHILD pidl, REFPROPERTYKEY propkey, REFPROPVARIANT propvar);
    HRESULT __stdcall GetViewProperty(PCUITEMID_CHILD pidl, REFPROPERTYKEY propkey, PROPVARIANT *ppropvar);
    HRESULT __stdcall SetTileViewProperties(PCUITEMID_CHILD pidl, LPCWSTR pszPropList);
    HRESULT __stdcall SetExtendedTileViewProperties(PCUITEMID_CHILD pidl, LPCWSTR pszPropList);
    HRESULT __stdcall SetText(FVTEXTTYPE iType, LPCWSTR pwszText);
    HRESULT __stdcall SetCurrentFolderFlags(DWORD dwMask, DWORD dwFlags);
    HRESULT __stdcall GetCurrentFolderFlags(DWORD *pdwFlags);
    HRESULT __stdcall GetSortColumnCount(int *pcColumns);
    HRESULT __stdcall SetSortColumns(const SORTCOLUMN *rgSortColumns, int cColumns);
    HRESULT __stdcall GetSortColumns(SORTCOLUMN *rgSortColumns, int cColumns);
    HRESULT __stdcall GetItem(int iItem, REFIID riid, void **ppv);
    HRESULT __stdcall GetVisibleItem(int iStart, WINBOOL fPrevious, int *piItem);
    HRESULT __stdcall GetSelectedItem(int iStart, int *piItem);
    HRESULT __stdcall GetSelection(WINBOOL fNoneImpliesFolder, IShellItemArray **ppsia);
    HRESULT __stdcall GetSelectionState(PCUITEMID_CHILD pidl, DWORD *pdwFlags);
    HRESULT __stdcall InvokeVerbOnSelection(LPCSTR pszVerb);
    HRESULT __stdcall SetViewModeAndIconSize(FOLDERVIEWMODE uViewMode, int iImageSize);
    HRESULT __stdcall GetViewModeAndIconSize(FOLDERVIEWMODE *puViewMode, int *piImageSize);
    HRESULT __stdcall SetGroupSubsetCount(UINT cVisibleRows);
    HRESULT __stdcall GetGroupSubsetCount(UINT *pcVisibleRows);
    HRESULT __stdcall SetRedraw(WINBOOL fRedrawOn);
    HRESULT __stdcall IsMoveInSameFolder();
    HRESULT __stdcall DoRename();

    LRESULT wndproc(UINT uMessage, WPARAM wParam, LPARAM lParam);

    HWND wnd = nullptr;

private:
    void on_create();
    void on_size(unsigned int width, unsigned int height);

    LONG refcount = 0;
    com_object<IShellBrowser> shell_browser;
    HWND wnd_parent = nullptr, wnd_list = nullptr;
    unsigned int view_mode, flags;
    com_object<IImageList> image_list_large, image_list_small;
    std::optional<SHELLVIEWID> view_id = *supported_view_ids[0];
    com_object<IShellFolder> folder;
};

typedef struct {
    std::u16string name;
} tar_item;

class shell_folder : public IShellFolder2, public IPersistFolder3, public IPersistIDList, public IObjectWithFolderEnumMode {
public:
    shell_folder();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IShellFolder2

    HRESULT __stdcall ParseDisplayName(HWND hwnd, IBindCtx *pbc, LPWSTR pszDisplayName, ULONG *pchEaten,
                                       PIDLIST_RELATIVE *ppidl, ULONG *pdwAttributes);
    HRESULT __stdcall EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList **ppenumIDList);
    HRESULT __stdcall BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv);
    HRESULT __stdcall BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv);
    HRESULT __stdcall CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2);
    HRESULT __stdcall CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv);
    HRESULT __stdcall GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF *rgfInOut);
    HRESULT __stdcall GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                    UINT *rgfReserved, void **ppv);
    HRESULT __stdcall GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET *pName);
    HRESULT __stdcall SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, LPCWSTR pszName, SHGDNF uFlags,
                                PITEMID_CHILD *ppidlOut);
    HRESULT __stdcall GetDefaultSearchGUID(GUID *pguid);
    HRESULT __stdcall EnumSearches(IEnumExtraSearch **ppenum);
    HRESULT __stdcall GetDefaultColumn(DWORD dwRes, ULONG *pSort, ULONG *pDisplay);
    HRESULT __stdcall GetDefaultColumnState(UINT iColumn, SHCOLSTATEF *pcsFlags);
    HRESULT __stdcall GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID *pscid, VARIANT *pv);
    HRESULT __stdcall GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS *psd);
    HRESULT __stdcall MapColumnToSCID(UINT iColumn, SHCOLUMNID *pscid);

    // IPersistFolder3

    HRESULT __stdcall GetClassID(CLSID* pClassID);
    HRESULT __stdcall Initialize(PCIDLIST_ABSOLUTE pidl);
    HRESULT __stdcall GetCurFolder(PIDLIST_ABSOLUTE* ppidl);
    HRESULT __stdcall InitializeEx(IBindCtx* pbc, PCIDLIST_ABSOLUTE pidlRoot, const PERSIST_FOLDER_TARGET_INFO* ppfti);
    HRESULT __stdcall GetFolderTargetInfo(PERSIST_FOLDER_TARGET_INFO* ppfti);

    // IPersistIDList

    HRESULT __stdcall SetIDList(PCIDLIST_ABSOLUTE pidl);
    HRESULT __stdcall GetIDList(PIDLIST_ABSOLUTE* ppidl);

    // IObjectWithFolderEnumMode

    HRESULT __stdcall SetMode(FOLDER_ENUM_MODE feMode);
    HRESULT __stdcall GetMode(FOLDER_ENUM_MODE *pfeMode);

private:
    LONG refcount = 0;
    std::vector<uint8_t> item_id_buf;
    SHITEMID* item_id = nullptr;
    FOLDER_ENUM_MODE folder_enum_mode = FEM_VIEWRESULT;
    std::vector<tar_item> items;
};

class shell_enum : public IEnumIDList {
public:
    shell_enum(const std::vector<tar_item>& items, SHCONTF flags): items(items), flags(flags) {
    }

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IEnumIDList

    HRESULT __stdcall Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched);
    HRESULT __stdcall Skip(ULONG celt);
    HRESULT __stdcall Reset();
    HRESULT __stdcall Clone(IEnumIDList** ppenum);

private:
    SHCONTF flags;
    std::vector<tar_item> items;
    LONG refcount = 0;
    size_t index = 0;
};

__inline std::u16string utf8_to_utf16(const std::string_view& s) {
    std::u16string ret;

    if (s.empty())
        return u"";

    auto len = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.length(), nullptr, 0);

    if (len == 0)
        throw std::runtime_error("MultiByteToWideChar 1 failed.");

    ret.resize(len);

    len = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.length(), (wchar_t*)ret.data(), len);

    if (len == 0)
        throw std::runtime_error("MultiByteToWideChar 2 failed.");

    return ret;
}

__inline std::string utf16_to_utf8(const std::u16string_view& s) {
    std::string ret;

    if (s.empty())
        return "";

    auto len = WideCharToMultiByte(CP_UTF8, 0, (const wchar_t*)s.data(), (int)s.length(), nullptr, 0,
                                   nullptr, nullptr);

    if (len == 0)
        throw std::runtime_error("WideCharToMultiByte 1 failed.");

    ret.resize(len);

    len = WideCharToMultiByte(CP_UTF8, 0, (const wchar_t*)s.data(), (int)s.length(), ret.data(), len,
                              nullptr, nullptr);

    if (len == 0)
        throw std::runtime_error("WideCharToMultiByte 2 failed.");

    return ret;
}

class last_error : public std::exception {
public:
    last_error(const std::string_view& function, int le) {
        std::string nice_msg;
        char16_t* fm;

        if (FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr,
                            le, 0, reinterpret_cast<LPWSTR>(&fm), 0, nullptr)) {
            try {
                std::u16string_view s = fm;

                while (!s.empty() && (s[s.length() - 1] == u'\r' || s[s.length() - 1] == u'\n')) {
                    s.remove_suffix(1);
                }

                nice_msg = utf16_to_utf8(s);
            } catch (...) {
                LocalFree(fm);
                throw;
            }

            LocalFree(fm);
        }

        msg = std::string(function) + " failed (error " + std::to_string(le) + (!nice_msg.empty() ? (", " + nice_msg) : "") + ").";
    }

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};

template<>
struct fmt::formatter<GUID> {
    constexpr auto parse(format_parse_context& ctx) {
            auto it = ctx.begin(), end = ctx.end();

            if (*it != '}')
                    throw format_error("invalid format");

            it++;

            return it;
    }

    template<typename format_context>
    auto format(const GUID& g, format_context& ctx) {
            return format_to(ctx.out(), "{{{:08x}-{:04x}-{:04x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}}}",
                             g.Data1, g.Data2, g.Data3, g.Data4[0], g.Data4[1], g.Data4[2], g.Data4[3], g.Data4[4],
                             g.Data4[5], g.Data4[6], g.Data4[7]);
    }
};
