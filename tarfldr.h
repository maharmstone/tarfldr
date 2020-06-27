#define _WIN32_WINNT 0x0601 // Windows 7

#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <stdexcept>
#include <filesystem>
#include <stdint.h>
#include <shlguid.h>
#include <fmt/format.h>
#include <archive.h>
#include <archive_entry.h>

extern const GUID CLSID_TarFolder;

#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + " stub"s).c_str()); return E_NOTIMPL;

template<typename... Args>
static void debug(const std::string_view& s, Args&&... args) { // FIXME - only if compiled in Debug mode
    std::string msg = fmt::format(s, std::forward<Args>(args)...);

    OutputDebugStringA(msg.c_str());
}

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

class handle_closer {
public:
    typedef HANDLE pointer;

    void operator()(HANDLE h) {
        if (h == INVALID_HANDLE_VALUE)
            return;

        CloseHandle(h);
    }
};

typedef std::unique_ptr<HANDLE, handle_closer> unique_handle;

typedef struct {
    std::u16string name;
    int fmt;
    int cxChar;
    const GUID* fmtid;
    DWORD pid;
} header_info;

class tar_item {
public:
    tar_item(const std::string_view& name, bool dir, const std::string_view& full_path) : name(name), dir(dir), full_path(full_path) { }

    ITEMID_CHILD* make_pidl_child() const;
    void find_child(const std::u16string_view& name, tar_item** ret);
    SFGAOF get_atts() const;

    std::string name, full_path;
    bool dir;
    std::vector<tar_item> children;
};

class tar_info {
public:
    tar_info(const std::filesystem::path& fn);

    void extract_file(const std::string& path, const std::filesystem::path& dest);

    tar_item root;

private:
    void add_entry(const std::string_view& fn);

    const std::filesystem::path archive_fn;
};

class shell_folder : public IShellFolder2, public IPersistFolder3, public IObjectWithFolderEnumMode, public IShellFolderViewCB {
public:
    shell_folder() { }
    shell_folder(const std::shared_ptr<tar_info>& tar, tar_item* root, PCIDLIST_ABSOLUTE pidl);
    virtual ~shell_folder();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IShellFolder2

    HRESULT __stdcall ParseDisplayName(HWND hwnd, IBindCtx* pbc, LPWSTR pszDisplayName, ULONG* pchEaten,
                                       PIDLIST_RELATIVE* ppidl, ULONG* pdwAttributes);
    HRESULT __stdcall EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList **ppenumIDList);
    HRESULT __stdcall BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv);
    HRESULT __stdcall BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx *pbc, REFIID riid, void **ppv);
    HRESULT __stdcall CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2);
    HRESULT __stdcall CreateViewObject(HWND hwndOwner, REFIID riid, void **ppv);
    HRESULT __stdcall GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, SFGAOF* rgfInOut);
    HRESULT __stdcall GetUIObjectOf(HWND hwndOwner, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                    UINT* rgfReserved, void** ppv);
    HRESULT __stdcall GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName);
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

    // IShellFolderViewCB

    HRESULT __stdcall MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    tar_item& get_item_from_pidl_child(const ITEMID_CHILD* pidl);

    LONG refcount = 0;
    FOLDER_ENUM_MODE folder_enum_mode = FEM_VIEWRESULT;
    std::shared_ptr<tar_info> tar;
    tar_item* root;
    PIDLIST_ABSOLUTE root_pidl = nullptr;
};

class shell_enum : public IEnumIDList {
public:
    shell_enum(const std::shared_ptr<tar_info>& tar, tar_item* root, SHCONTF flags): tar(tar), root(root), flags(flags) { }

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
    std::shared_ptr<tar_info> tar;
    tar_item* root;
    LONG refcount = 0;
    size_t index = 0;
};

class shell_item : public IContextMenu, public IDataObject {
public:
    shell_item(PIDLIST_ABSOLUTE root_pidl, const std::shared_ptr<tar_info>& tar,
               const std::vector<tar_item*>& itemlist);
    virtual ~shell_item();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IContextMenu

    HRESULT __stdcall QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst,
                                       UINT idCmdLast, UINT uFlags);
    HRESULT __stdcall InvokeCommand(CMINVOKECOMMANDINFO* pici);
    HRESULT __stdcall GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved,
                                       CHAR* pszName, UINT cchMax);

    // IDataObject

    HRESULT __stdcall GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium);
    HRESULT __stdcall GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium);
    HRESULT __stdcall QueryGetData(FORMATETC* pformatetc);
    HRESULT __stdcall GetCanonicalFormatEtc(FORMATETC* pformatectIn,
                                            FORMATETC* pformatetcOut);
    HRESULT __stdcall SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, WINBOOL fRelease);
    HRESULT __stdcall EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc);
    HRESULT __stdcall DAdvise(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink,
                              DWORD* pdwConnection);
    HRESULT __stdcall DUnadvise(DWORD dwConnection);
    HRESULT __stdcall EnumDAdvise(IEnumSTATDATA* *ppenumAdvise);

private:
    HGLOBAL make_shell_id_list();

    LONG refcount = 0;
    PIDLIST_ABSOLUTE root_pidl;
    std::shared_ptr<tar_info> tar;
    std::vector<tar_item*> itemlist;
    CLIPFORMAT cf_shell_id_list, cf_file_contents;
};

struct data_format {
    data_format(CLIPFORMAT format, DWORD tymed) : format(format), tymed(tymed) { }

    CLIPFORMAT format;
    DWORD tymed;
};

class shell_item_enum_format : public IEnumFORMATETC {
public:
    shell_item_enum_format(CLIPFORMAT cf_shell_id_list, CLIPFORMAT cf_file_contents);

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IEnumFORMATETC

    HRESULT __stdcall Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched);
    HRESULT __stdcall Skip(ULONG celt);
    HRESULT __stdcall Reset();
    HRESULT __stdcall Clone(IEnumFORMATETC** ppenum);

private:
    LONG refcount = 0;
    std::vector<data_format> formats;
    unsigned int index = 0;
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
            return format_to(ctx.out(), "{{{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}}}",
                             g.Data1, g.Data2, g.Data3, g.Data4[0], g.Data4[1], g.Data4[2], g.Data4[3], g.Data4[4],
                             g.Data4[5], g.Data4[6], g.Data4[7]);
    }
};

class formatted_error : public std::exception {
public:
    template<typename... Args>
    formatted_error(const std::string_view& s, Args&&... args) {
        msg = fmt::format(s, std::forward<Args>(args)...);
    }

    const char* what() const noexcept {
        return msg.c_str();
    }

private:
    std::string msg;
};
