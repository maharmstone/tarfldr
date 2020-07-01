#include "tarfldr.h"

using namespace std;

const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };
const GUID CLSID_TarContextMenu = { 0xa23f73ab, 0x6c42, 0x4689, {0xa6, 0xab, 0x30, 0x13, 0x0c, 0xe7, 0x2a, 0x90 } };

static const array file_extensions = { u".tar", u".gz", u".bz2", u".xz", u".tgz", u".tbz2", u".txz" };
static const array prog_ids = { make_tuple(u"TarFolder", 0), make_tuple(u"TarFolderCompressed", 1) }; // name, icon number

#define BLOCK_SIZE 20480

LONG objs_loaded = 0;
HINSTANCE instance = nullptr;

void tar_info::add_entry(const string_view& fn, int64_t size, const optional<time_t>& mtime, bool is_dir,
                         const char* user, const char* group, mode_t mode) {
    vector<string_view> parts;
    string_view file_part;
    tar_item* r;

    // split by slashes

    {
        string_view left = fn;

        do {
            bool found = false;

            for (unsigned int i = 0; i < left.size(); i++) {
                if (left[i] == '/') {
                    if (i != 0 && (i != 1 || left[0] != '.'))
                        parts.emplace_back(left.data(), i);

                    left = left.substr(i + 1);
                    found = true;
                    break;
                }
            }

            if (!found) {
                parts.emplace_back(left);
                break;
            }
        } while (true);
    }

    while (!parts.empty() && parts.back().empty()) {
        parts.pop_back();
    }

    if (parts.empty())
        return;

    file_part = parts.back();
    parts.pop_back();

    r = &root;

    // add dirs

    for (const auto& p : parts) {
        bool found = false;

        for (auto& c : r->children) {
            if (c.name == p) {
                found = true;
                r = &c;
                break;
            }
        }

        if (!found) {
            r->children.emplace_back(p, size, true, "", nullopt, "", "", 0, r);
            r = &r->children.back();
        }
    }

    // add child

    if (!file_part.empty()) {
        r->children.emplace_back(file_part, size, is_dir, fn, mtime, user ? user : "",
                                 group ? group : "", mode, r);
    }
}

enum archive_type identify_file_type(const u16string_view& fn2) {
    enum archive_type type = archive_type::unknown;
    auto st = fn2.rfind(u".");
    u16string last_ext;

    if (st == string::npos)
        return archive_type::unknown;

    last_ext = fn2.substr(st + 1);

    for (auto& c : last_ext) {
        c = tolower(c);
    }

    if (last_ext == u"gz")
        type = archive_type::gzip;
    else if (last_ext == u"bz2")
        type = archive_type::bz2;
    else if (last_ext == u"xz")
        type = archive_type::xz;
    else if (last_ext == u"tar")
        type = archive_type::tarball;

    if (st != 0 && (type == archive_type::gzip || type == archive_type::bz2 || type == archive_type::xz)) {
        auto st2 = fn2.rfind(u".", st - 1);

        if (st2 != string::npos) {
            u16string penult_ext{fn2.substr(st2 + 1, st - st2 - 1)};

            for (auto& c : penult_ext) {
                c = tolower(c);
            }

            if (penult_ext == u"tar")
                type |= archive_type::tarball;
        }
    }

    return type;
}

tar_info::tar_info(const filesystem::path& fn) : archive_fn(fn), root("", 0, true, "", nullopt, "", "", 0, nullptr) {
    auto fn2 = fn.filename().u16string();
    bool is_tarball = true;

    type = identify_file_type(fn2);

    if (type & archive_type::tarball) {
        struct archive_entry* entry;
        struct archive* a = archive_read_new();

        try {
            archive_read_support_filter_all(a);
            archive_read_support_format_all(a);

            auto r = archive_read_open_filename(a, (char*)fn.u8string().c_str(), BLOCK_SIZE);

            if (r != ARCHIVE_OK)
                throw runtime_error(archive_error_string(a));

            while (archive_read_next_header(a, &entry) == ARCHIVE_OK) {
                if (archive_entry_pathname_utf8(entry)) {
                    add_entry(archive_entry_pathname_utf8(entry), archive_entry_size(entry),
                            archive_entry_mtime_is_set(entry) ? optional<time_t>{archive_entry_mtime(entry)} : optional<time_t>{nullopt},
                            archive_entry_filetype(entry) == AE_IFDIR, archive_entry_uname_utf8(entry),
                            archive_entry_gname_utf8(entry), archive_entry_mode(entry));
                }
            }
        } catch (...) {
            archive_read_free(a);
            throw;
        }

        archive_read_free(a);
    } else if (type & archive_type::gzip || type & archive_type::bz2 || type & archive_type::xz) {
        auto st = fn2.rfind(u".");
        auto orig_fn = fn2.substr(0, st);
        size_t size = 0;
        optional<time_t> mtime = nullopt; // FIXME

        if (type & archive_type::gzip) {
            LARGE_INTEGER li;
            uint32_t size2;
            DWORD read;

            // Read uncompressed length from end of file

            // FIXME - this isn't reliable: see https://stackoverflow.com/questions/9209138/uncompressed-file-size-using-zlibs-gzip-file-access-function

            unique_handle h{CreateFileW((LPCWSTR)fn.u16string().c_str(), GENERIC_READ, 0, nullptr,
                                        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)};

            if (h.get() == INVALID_HANDLE_VALUE)
                throw last_error("CreateFile", GetLastError());

            li.QuadPart = -((int64_t)sizeof(uint32_t));

            if (!SetFilePointerEx(h.get(), li, nullptr, FILE_END))
                throw last_error("SetFilePointerEx", GetLastError());

            if (!ReadFile(h.get(), &size2, sizeof(uint32_t), &read, nullptr))
                throw last_error("ReadFile", GetLastError());

            size = size2;
        } else if (type & archive_type::bz2) {
            auto bzf = BZ2_bzopen((char*)fn.u8string().c_str(), "r");

            if (!bzf)
                throw formatted_error("Could not open bzip2 file {}.", fn.string());

            while (true) {
                char buf[4096];

                auto ret = BZ2_bzread(bzf, buf, sizeof(buf));

                if (ret <= 0)
                    break;

                size += ret;
            }

            BZ2_bzclose(bzf);
        } else if (type & archive_type::xz) {
            uint8_t inbuf[4096], outbuf[4096];
            DWORD read;

            lzma_stream strm = LZMA_STREAM_INIT;

            unique_handle h{CreateFileW((LPCWSTR)fn.u16string().c_str(), GENERIC_READ, 0, nullptr,
                                        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)};

            if (h.get() == INVALID_HANDLE_VALUE)
                throw last_error("CreateFile", GetLastError());

            auto ret = lzma_stream_decoder(&strm, UINT64_MAX, 0);

            if (ret != LZMA_OK)
                throw formatted_error("lzma_stream_decoder returned {}.", ret);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = outbuf;
            strm.avail_out = sizeof(outbuf);

            while (true) {
                if (strm.avail_in == 0) {
                    strm.next_in = inbuf;

                    if (!ReadFile(h.get(), inbuf, sizeof(inbuf), &read, nullptr))
                        throw last_error("ReadFile", GetLastError());

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        break;
                }

                auto ret = lzma_code(&strm, LZMA_RUN);

                if (strm.avail_out == 0 || ret == LZMA_STREAM_END)
                    size += sizeof(outbuf) - strm.avail_out;

                if (strm.avail_out == 0) {
                    strm.next_out = outbuf;
                    strm.avail_out = sizeof(outbuf);
                }

                if (ret == LZMA_STREAM_END)
                    break;

                if (ret != LZMA_OK)
                    throw formatted_error("lzma_code returned {}.", ret);
            }
        }

        add_entry(utf16_to_utf8(orig_fn).c_str(), size, mtime, false, nullptr, nullptr, 0);
    }
}

extern "C" STDAPI DllCanUnloadNow(void) {
    return objs_loaded == 0 ? S_OK : S_FALSE;
}

static void create_reg_key(HKEY hkey, const u16string& subkey, const u16string& value = u"", bool expand_sz = false) {
    LSTATUS ret;
    HKEY hk;
    DWORD dispos;

    ret = RegCreateKeyExW(hkey, (WCHAR*)subkey.c_str(), 0, nullptr, 0, KEY_ALL_ACCESS, nullptr, &hk, &dispos);
    if (ret != ERROR_SUCCESS)
        throw formatted_error("RegCreateKeyEx failed for {} (error {}).", utf16_to_utf8(subkey), ret);

    if (!value.empty()) {
        ret = RegSetValueExW(hk, nullptr, 0, expand_sz ? REG_EXPAND_SZ : REG_SZ, (BYTE*)value.c_str(), (value.length() + 1) * sizeof(char16_t));
        if (ret != ERROR_SUCCESS) {
            RegCloseKey(hk);
            throw formatted_error("RegCreateKeyEx failed for root of {} (error {}).", utf16_to_utf8(subkey), ret);
        }
    }

    RegCloseKey(hk);
}

static void set_reg_value(HKEY hkey, const u16string& path, const u16string& name, const u16string& value) {
    LSTATUS ret;
    HKEY key;

    ret = RegOpenKeyW(hkey, (WCHAR*)path.c_str(), &key);
    if (ret != ERROR_SUCCESS)
        throw formatted_error("RegOpenKey failed for {} (error {}).", utf16_to_utf8(path), ret);

    ret = RegSetValueExW(key, (WCHAR*)name.c_str(), 0, REG_SZ, (BYTE*)value.c_str(), (value.length() + 1) * sizeof(char16_t));
    if (ret != ERROR_SUCCESS) {
        RegCloseKey(key);
        throw formatted_error("RegCreateKeyEx failed for {}\\{} (error {}).", utf16_to_utf8(path), utf16_to_utf8(name), ret);
    }

    RegCloseKey(key);
}

static void set_reg_value(HKEY hkey, const u16string& path, const u16string& name, uint32_t value) {
    LSTATUS ret;
    HKEY key;

    ret = RegOpenKeyW(hkey, (WCHAR*)path.c_str(), &key);
    if (ret != ERROR_SUCCESS)
        throw formatted_error("RegOpenKey failed for {} (error {}).", utf16_to_utf8(path), ret);

    ret = RegSetValueExW(key, (WCHAR*)name.c_str(), 0, REG_DWORD, (BYTE*)&value, sizeof(value));
    if (ret != ERROR_SUCCESS) {
        RegCloseKey(key);
        throw formatted_error("RegCreateKeyEx failed for {}\\{} (error {}).", utf16_to_utf8(path), utf16_to_utf8(name), ret);
    }

    RegCloseKey(key);
}

extern "C" HRESULT DllRegisterServer() {
    try {
        auto clsid = utf8_to_utf16(fmt::format("{}", CLSID_TarFolder));
        char16_t file[MAX_PATH]; // FIXME - size dynamically?
        DWORD ret;

        ret = GetModuleFileNameW(instance, (WCHAR*)file, sizeof(file) / sizeof(char16_t));
        if (ret == 0)
            throw last_error("GetModuleFileName", ret);

        auto browsable_shellext = utf8_to_utf16(fmt::format("{}", CATID_BrowsableShellExt));
        auto exec_folder = utf8_to_utf16(fmt::format("{}", CLSID_ExecuteFolder));

        for (const auto& ext : file_extensions) {
            const auto& prog_id = get<0>(ext == u".tar"s ? prog_ids[0] : prog_ids[1]);

            create_reg_key(HKEY_CLASSES_ROOT, ext, prog_id);
            create_reg_key(HKEY_CLASSES_ROOT, ext + u"\\"s + prog_id);

            set_reg_value(HKEY_CLASSES_ROOT, ext, u"PerceivedType", u"compressed");

            create_reg_key(HKEY_CLASSES_ROOT, ext + u"\\OpenWithProgids"s);
            set_reg_value(HKEY_CLASSES_ROOT, ext + u"\\OpenWithProgids"s, u"TarFolder", u"");
        }

        auto clsid_menu = utf8_to_utf16(fmt::format("{}", CLSID_TarContextMenu));

        for (const auto& p : prog_ids) {
            const auto& prog_id = get<0>(p);

            create_reg_key(HKEY_CLASSES_ROOT, prog_id);
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\CLSID"s, clsid);
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\DefaultIcon"s, file + u","s + utf8_to_utf16(to_string(get<1>(p))));

            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\shell"s);
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\shell\\Open"s);
            set_reg_value(HKEY_CLASSES_ROOT, prog_id + u"\\shell\\Open"s, u"MultiSelectModel", u"Document");
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\shell\\Open\\Command"s, u"%SystemRoot%\\Explorer.exe /idlist,%I,%L", true);
            set_reg_value(HKEY_CLASSES_ROOT, prog_id + u"\\shell\\Open\\Command"s, u"DelegateExecute", exec_folder);

            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\ShellEx"s);
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\ShellEx\\StorageHandler"s, clsid);

            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\ShellEx\\ContextMenuHandlers"s);
            create_reg_key(HKEY_CLASSES_ROOT, prog_id + u"\\ShellEx\\ContextMenuHandlers\\tarfldr"s, clsid_menu);
        }

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid, get<0>(prog_ids[0]));
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\DefaultIcon", file + u",0"s);

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\Implemented Categories");
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\Implemented Categories\\" + browsable_shellext);
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\InProcServer32", file);
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\InProcServer32", u"ThreadingModel", u"Apartment");

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ProgID", get<0>(prog_ids[0]));

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ShellFolder");
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ShellFolder", u"Attributes", SFGAO_FOLDER);

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu, get<0>(prog_ids[0]));
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu + u"\\InProcServer32", file);
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu + u"\\InProcServer32", u"ThreadingModel", u"Apartment");

        return S_OK;
    } catch (const exception& e) {
        MessageBoxW(nullptr, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
        return E_FAIL;
    }
}

static void delete_reg_tree(HKEY hkey, const u16string& subkey) {
    LSTATUS ret;

    ret = RegDeleteTreeW(hkey, (WCHAR*)subkey.c_str());
    if (ret != ERROR_SUCCESS)
        throw formatted_error("RegDeleteTree failed for {} (error {}).", utf16_to_utf8(subkey), ret);
}

extern "C" HRESULT DllUnregisterServer() {
    try {
        for (const auto& ext : file_extensions) {
            delete_reg_tree(HKEY_CLASSES_ROOT, ext);
        }

        delete_reg_tree(HKEY_CLASSES_ROOT, u"CLSID\\" + utf8_to_utf16(fmt::format("{}", CLSID_TarFolder)));
        delete_reg_tree(HKEY_CLASSES_ROOT, u"CLSID\\" + utf8_to_utf16(fmt::format("{}", CLSID_TarContextMenu)));

        for (const auto& prog_id : prog_ids) {
            delete_reg_tree(HKEY_CLASSES_ROOT, get<0>(prog_id));
        }

        return S_OK;
    } catch (const exception& e) {
        MessageBoxW(nullptr, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
        return E_FAIL;
    }
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

ITEMID_CHILD* tar_item::make_relative_pidl(tar_item* root) const {
    size_t size = offsetof(ITEMIDLIST, mkid.abID);

    const tar_item* p = this;

    while (p && p != root) {
        size += offsetof(ITEMIDLIST, mkid.abID) + p->name.length();
        p = p->parent;
    }

    auto item = (ITEMIDLIST*)CoTaskMemAlloc(size);

    if (!item)
        throw bad_alloc();

    auto ptr = (ITEMIDLIST*)((uint8_t*)item + size - offsetof(ITEMIDLIST, mkid.abID));

    ptr->mkid.cb = 0;

    p = this;

    while (p && p != root) {
        ptr = (ITEMIDLIST*)((uint8_t*)ptr - offsetof(ITEMIDLIST, mkid.abID) - p->name.length());

        ptr->mkid.cb = offsetof(ITEMIDLIST, mkid.abID) + p->name.length();
        memcpy(ptr->mkid.abID, p->name.data(), p->name.length());

        p = p->parent;
    }

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
    SFGAOF atts = SFGAO_CANCOPY;

    if (dir) {
        atts |= SFGAO_FOLDER | SFGAO_BROWSABLE;
        atts |= SFGAO_HASSUBFOLDER; // FIXME - check for this
    } else
        atts |= SFGAO_STREAM;

    // FIXME - SFGAO_CANRENAME, SFGAO_CANDELETE, SFGAO_HIDDEN, etc.

    return atts;
}

factory::factory(const CLSID& clsid) : clsid(clsid) {
    InterlockedIncrement(&objs_loaded);
}

factory::~factory() {
    InterlockedDecrement(&objs_loaded);
}

HRESULT factory::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IClassFactory)
        *ppv = static_cast<IClassFactory*>(this);
    else {
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG factory::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG factory::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT factory::CreateInstance(IUnknown* pUnknownOuter, const IID& iid, void** ppv) {
    if (clsid == CLSID_TarFolder) {
        if (iid == IID_IUnknown || iid == IID_IShellFolder || iid == IID_IShellFolder2 ||
            iid == IID_IPersist || iid == IID_IPersistFolder || iid == IID_IPersistFolder2 || iid == IID_IPersistFolder3 ||
            iid == IID_IObjectWithFolderEnumMode || iid == IID_IShellFolderViewCB) {
            auto sf = new shell_folder;
            if (!sf)
                return E_OUTOFMEMORY;

            return sf->QueryInterface(iid, ppv);
        }
    } else if (clsid == CLSID_TarContextMenu) {
        if (iid == IID_IUnknown || iid == IID_IContextMenu || iid == IID_IShellExtInit) {
            auto scm = new shell_context_menu;
            if (!scm)
                return E_OUTOFMEMORY;

            return scm->QueryInterface(iid, ppv);
        }
    }

    debug("factory::CreateInstance: unsupported interface {} on {}", iid, clsid);

    *ppv = nullptr;
    return E_NOINTERFACE;
}

HRESULT factory::LockServer(BOOL bLock) {
    return E_NOTIMPL;
}

extern "C" STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    if (rclsid == CLSID_TarFolder || rclsid == CLSID_TarContextMenu) {
        factory* fact = new factory(rclsid);
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
