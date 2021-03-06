/* Copyright (c) Mark Harmstone 2020
 *
 * This file is part of tarfldr.
 *
 * tarfldr is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public Licence as published by
 * the Free Software Foundation, either version 3 of the Licence, or
 * (at your option) any later version.
 *
 * tarfldr is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public Licence for more details.
 *
 * You should have received a copy of the GNU Lesser General Public Licence
 * along with tarfldr.  If not, see <http://www.gnu.org/licenses/>. */

#define _WIN32_WINNT 0x0601 // Windows 7

#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <list>
#include <stdexcept>
#include <filesystem>
#include <optional>
#include <functional>
#include <stdint.h>
#include <shlguid.h>
#include <fmt/format.h>
#include <archive.h>
#include <archive_entry.h>
#include <zlib.h>
#include <bzlib.h>
#include <lzma.h>

extern const GUID CLSID_TarFolder;
extern const GUID FMTID_POSIXAttributes;

#define PID_POSIX_USER      2
#define PID_POSIX_GROUP     3
#define PID_POSIX_MODE      4

extern HINSTANCE instance;

#ifdef _DEBUG
#define UNIMPLEMENTED OutputDebugStringA((__PRETTY_FUNCTION__ + std::string(" stub\n")).c_str()); return E_NOTIMPL;
#else
#define UNIMPLEMENTED return E_NOTIMPL;
#endif

template<typename... Args>
static __inline void debug(const std::string_view& s, Args&&... args) {
#ifdef _DEBUG
    std::string msg = fmt::format(s, std::forward<Args>(args)...);

    OutputDebugStringA(msg.c_str());
#endif
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

class tar_item {
public:
    tar_item(const std::string_view& name, int64_t size, bool dir,
             const std::string_view& full_path, const std::optional<time_t>& mtime,
             const std::string_view& user, const std::string_view& group,
             mode_t mode, tar_item* parent) :
        name(name), size(size), dir(dir), full_path(full_path), mtime(mtime), user(user), group(group), mode(mode), parent(parent) { }

    ITEMID_CHILD* make_pidl_child() const;
    ITEMID_CHILD* make_relative_pidl(tar_item* root) const;
    void find_child(const std::u16string_view& name, tar_item** ret);
    SFGAOF get_atts() const;

    std::string name, full_path, user, group;
    tar_item* parent;
    int64_t size;
    bool dir;
    std::list<tar_item> children;
    std::optional<time_t> mtime;
    mode_t mode;
};

enum class archive_type {
    unknown = 0,
    tarball = 1,
    gzip = 2,
    bz2 = 4,
    xz = 8
};

static bool operator&(const archive_type& a, const archive_type& b) {
    return (int)a & (int)b;
}

static void operator|=(archive_type& a, const archive_type& b) {
    a = (archive_type)((int)a | (int)b);
}

static archive_type operator|(archive_type a, archive_type b) {
    return (archive_type)((int)a | (int)b);
}

class tar_info {
public:
    tar_info(const std::filesystem::path& fn);

    tar_item root;
    const std::filesystem::path archive_fn;
    enum archive_type type;

private:
    void add_entry(const std::string_view& fn, int64_t size, const std::optional<time_t>& mtime, bool is_dir,
                   const char* user, const char* group, mode_t mode);
};

class factory : public IClassFactory {
public:
    factory(const CLSID& clsid);
    virtual ~factory();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IClassFactory

    HRESULT __stdcall CreateInstance(IUnknown* pUnknownOuter, const IID& iid, void** ppv);
    HRESULT __stdcall LockServer(BOOL bLock);

private:
    CLSID clsid;
    LONG refcount = 0;
};

class shell_folder : public IShellFolder2, public IPersistFolder3, public IObjectWithFolderEnumMode, public IExplorerPaneVisibility {
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
    HRESULT __stdcall GetDefaultColumn(DWORD dwRes, ULONG* pSort, ULONG* pDisplay);
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

    // IExplorerPaneVisibility

    HRESULT __stdcall GetPaneState(REFEXPLORERPANE ep, EXPLORERPANESTATE* peps);

    int name_compare(const tar_item& item1, const tar_item& item2) const;
    int size_compare(const tar_item& item1, const tar_item& item2) const;

private:
    tar_item& get_item_from_pidl_child(const ITEMID_CHILD* pidl);
    tar_item& get_item_from_relative_pidl(PCUIDLIST_RELATIVE pidl);

    LONG refcount = 0;
    FOLDER_ENUM_MODE folder_enum_mode = FEM_VIEWRESULT;
    std::shared_ptr<tar_info> tar;
    tar_item* root;
    PIDLIST_ABSOLUTE root_pidl = nullptr;
    std::filesystem::path path;
};

typedef struct {
    unsigned int res_num;
    int fmt;
    int cxChar;
    const GUID* fmtid;
    DWORD pid;
    std::function<int(const tar_item&, const tar_item&)> compare_func;
    SHCOLSTATEF state;
    bool tarball_only;
} header_info;

class shell_enum : public IEnumIDList {
public:
    shell_enum(const std::shared_ptr<tar_info>& tar, tar_item* root, SHCONTF flags) :
        tar(tar), root(root), flags(flags), it(root->children.begin()) { }

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
    decltype(root->children.begin()) it;
};

class shell_item_details {
public:
    shell_item_details(tar_item* item, const std::u16string_view& relative_path) : item(item), relative_path(relative_path) { }

    tar_item* item;
    std::u16string relative_path;
};

class shell_item_list : public IContextMenu, public IDataObject {
public:
    shell_item_list(PIDLIST_ABSOLUTE root_pidl, const std::shared_ptr<tar_info>& tar,
                    const std::vector<tar_item*>& itemlist, tar_item* root, bool recursive, shell_folder* folder);
    virtual ~shell_item_list();

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
    HRESULT __stdcall EnumDAdvise(IEnumSTATDATA** ppenumAdvise);

    HRESULT open_cmd(CMINVOKECOMMANDINFO* pici);
    HRESULT copy_cmd(CMINVOKECOMMANDINFO* pici);
    HRESULT properties(CMINVOKECOMMANDINFO* pici);
    INT_PTR PropSheetDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
    std::u16string get_item_prop(tar_item& item, const GUID& guid, DWORD pid);
    uint64_t calc_size();

private:
    HGLOBAL make_shell_id_list();
    HGLOBAL make_file_descriptor();
    void populate_full_itemlist();
    void populate_full_itemlist2(tar_item* item, const std::u16string& prefix);

    LONG refcount = 0;
    PIDLIST_ABSOLUTE root_pidl;
    std::shared_ptr<tar_info> tar;
    std::vector<tar_item*> itemlist;
    std::vector<shell_item_details> full_itemlist;
    CLIPFORMAT cf_shell_id_list, cf_file_contents, cf_file_descriptor;
    tar_item* root;
    bool recursive;
    shell_folder* folder;
};

struct data_format {
    data_format(CLIPFORMAT format, DWORD tymed) : format(format), tymed(tymed) { }

    CLIPFORMAT format;
    DWORD tymed;
};

class shell_item_enum_format : public IEnumFORMATETC {
public:
    shell_item_enum_format(CLIPFORMAT cf_shell_id_list, CLIPFORMAT cf_file_contents,
                           CLIPFORMAT cf_file_descriptor);

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

class tar_item_stream : public IStream {
public:
    tar_item_stream(const std::shared_ptr<tar_info>& tar, tar_item& item);
    ~tar_item_stream();

    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // ISequentialStream

    HRESULT __stdcall Read(void* pv, ULONG cb, ULONG* pcbRead);
    HRESULT __stdcall Write(const void* pv, ULONG cb, ULONG* pcbWritten);

    // IStream

    HRESULT __stdcall Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition);
    HRESULT __stdcall SetSize(ULARGE_INTEGER libNewSize);
    HRESULT __stdcall CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten);
    HRESULT __stdcall Commit(DWORD grfCommitFlags);
    HRESULT __stdcall Revert();
    HRESULT __stdcall LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    HRESULT __stdcall UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    HRESULT __stdcall Stat(STATSTG* pstatstg, DWORD grfStatFlag);
    HRESULT __stdcall Clone(IStream** ppstm);

    void extract_file(const std::filesystem::path& fn);

private:
    LONG refcount = 0;
    struct archive* a = nullptr;
    tar_item& item;
    std::string buf;
    gzFile gzf = nullptr;
    unique_handle h;
    lzma_stream xz_strm = LZMA_STREAM_INIT;
    bz_stream bz2_strm;
    std::string inbuf, outbuf;
    enum archive_type type;
    bool eof = false;
    int lzma_ret = LZMA_OK;
    int bz2_ret = BZ_OK;
    uint64_t position = 0;
};

class shell_context_menu;

class shell_context_menu_item {
public:
    shell_context_menu_item(unsigned int res_num, const std::string_view& verba, const std::u16string_view& verbw,
                            const std::function<void(shell_context_menu*, CMINVOKECOMMANDINFO*)>& cmd, bool sub_item) :
                            res_num(res_num), verba(verba), verbw(verbw), cmd(cmd), sub_item(sub_item) {
    }

    unsigned int res_num;
    std::string verba;
    std::u16string verbw;
    std::function<void(shell_context_menu*, CMINVOKECOMMANDINFO*)> cmd;
    bool sub_item;
};

class shell_context_menu : public IContextMenu, public IShellExtInit {
public:
    // IUnknown

    HRESULT __stdcall QueryInterface(REFIID iid, void** ppv);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IContextMenu

    HRESULT __stdcall QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    HRESULT __stdcall InvokeCommand(CMINVOKECOMMANDINFO* pici);
    HRESULT __stdcall GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax);

    // IShellExtInit

    HRESULT __stdcall Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID);

    void extract_all(CMINVOKECOMMANDINFO* pici);
    void decompress(CMINVOKECOMMANDINFO* pici);
    void compress(CMINVOKECOMMANDINFO* pici, archive_type type);

private:
    LONG refcount = 0;
    std::vector<std::tuple<std::string, enum archive_type>> files;
    std::vector<shell_context_menu_item> items;
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

// folder.cpp
std::u16string mode_to_u16string(mode_t m);

// tarfldr.cpp
enum archive_type identify_file_type(const std::u16string_view& fn2);
