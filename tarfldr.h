#define _WIN32_WINNT 0x0601 // Windows 7

#include <windows.h>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <stdexcept>
#include <stdint.h>

class shell_view : public IShellView {
public:
    virtual ~shell_view();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IShellView

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

    LRESULT wndproc(UINT uMessage, WPARAM wParam, LPARAM lParam);

private:
    LONG refcount = 0;
    IShellBrowser* shell_browser = nullptr;
    HWND wnd = nullptr, wnd_parent = nullptr;
};

class shell_folder : public IShellFolder, public IPersistFolder3, public IPersistIDList, public IObjectWithFolderEnumMode {
public:
    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IShellFolder

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
