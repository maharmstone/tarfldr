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

typedef struct {
    std::u16string name;
    int fmt;
    int cxChar;
} header_info;

typedef struct {
    std::string name;
} tar_item;

class shell_folder : public IShellFolder2, public IPersistFolder3, public IObjectWithFolderEnumMode {
public:
    shell_folder();
    virtual ~shell_folder();

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

    // IObjectWithFolderEnumMode

    HRESULT __stdcall SetMode(FOLDER_ENUM_MODE feMode);
    HRESULT __stdcall GetMode(FOLDER_ENUM_MODE *pfeMode);

private:
    LONG refcount = 0;
    std::vector<uint8_t> item_id_buf;
    SHITEMID* item_id = nullptr;
    FOLDER_ENUM_MODE folder_enum_mode = FEM_VIEWRESULT;
    std::vector<tar_item> items;
    PIDLIST_ABSOLUTE root_pidl = nullptr;
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
