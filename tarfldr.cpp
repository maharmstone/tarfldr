#include "tarfldr.h"

using namespace std;

const GUID CLSID_TarFolder = { 0x95b57a60, 0xcb8e, 0x49fc, { 0x8d, 0x4c, 0xef, 0x12, 0x25, 0x20, 0x0d, 0x7d } };
const GUID CLSID_TarContextMenu = { 0xa23f73ab, 0x6c42, 0x4689, {0xa6, 0xab, 0x30, 0x13, 0x0c, 0xe7, 0x2a, 0x90 } };

static const array file_extensions = { u".tar", u".gz", u".bz2", u".xz", u".tgz", u".tbz2", u".txz" };

#define PROGID u"TarFolder"

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

tar_info::tar_info(const filesystem::path& fn) : archive_fn(fn), root("", 0, true, "", nullopt, "", "", 0, nullptr) {
    auto fn2 = fn.filename().u8string();
    bool is_tarball = true;
    decltype(fn2) last_ext;

    // Determine whether the file is a tarball or a single compressed file

    auto st = fn2.rfind(u8".");

    if (st == string::npos)
        throw runtime_error("Unsupported file type.");

    last_ext = fn2.substr(st + 1);

    for (auto& c : last_ext) {
        c = tolower(c);
    }

    if (st != 0 && (last_ext == u8"gz" || last_ext == u8"bz2" || last_ext == u8"xz")) {
        auto st2 = fn2.rfind(u8".", st - 1);

        if (st2 != string::npos) {
            decltype(fn2) penult_ext = fn2.substr(st2 + 1, st - st2 - 1);

            for (auto& c : penult_ext) {
                c = tolower(c);
            }

            is_tarball = penult_ext == u8"tar";
        } else
            is_tarball = false;
    }

    if (is_tarball) {
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

        single_file = false;
    } else {
        auto orig_fn = fn2.substr(0, st);
        size_t size = 0; // FIXME
        optional<time_t> mtime = nullopt; // FIXME

        add_entry((char*)orig_fn.c_str(), size, mtime, false, nullptr, nullptr, 0);

        single_file = true;
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
            create_reg_key(HKEY_CLASSES_ROOT, ext, PROGID);
            create_reg_key(HKEY_CLASSES_ROOT, ext + u"\\"s + PROGID);

            set_reg_value(HKEY_CLASSES_ROOT, ext, u"PerceivedType", u"compressed");

            create_reg_key(HKEY_CLASSES_ROOT, ext + u"\\OpenWithProgids"s);
            set_reg_value(HKEY_CLASSES_ROOT, ext + u"\\OpenWithProgids"s, u"TarFolder", u"");
        }

        create_reg_key(HKEY_CLASSES_ROOT, PROGID);
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\CLSID", clsid);
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\DefaultIcon", file + u",0"s);

        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\shell");
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\shell\\Open");
        set_reg_value(HKEY_CLASSES_ROOT, PROGID u"\\shell\\Open", u"MultiSelectModel", u"Document");
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\shell\\Open\\Command", u"%SystemRoot%\\Explorer.exe /idlist,%I,%L", true);
        set_reg_value(HKEY_CLASSES_ROOT, PROGID u"\\shell\\Open\\Command", u"DelegateExecute", exec_folder);

        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\ShellEx");
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\ShellEx\\StorageHandler", clsid);

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid, PROGID);
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\DefaultIcon", file + u",0"s);

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\Implemented Categories");
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\Implemented Categories\\" + browsable_shellext);
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\InProcServer32", file);
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\InProcServer32", u"ThreadingModel", u"Apartment");

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ProgID", PROGID);

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ShellFolder");
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid + u"\\ShellFolder", u"Attributes", SFGAO_FOLDER);

        auto clsid_menu = utf8_to_utf16(fmt::format("{}", CLSID_TarContextMenu));

        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu, PROGID);
        create_reg_key(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu + u"\\InProcServer32", file);
        set_reg_value(HKEY_CLASSES_ROOT, u"CLSID\\" + clsid_menu + u"\\InProcServer32", u"ThreadingModel", u"Apartment");

        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\ShellEx\\ContextMenuHandlers");
        create_reg_key(HKEY_CLASSES_ROOT, PROGID u"\\ShellEx\\ContextMenuHandlers\\tarfldr", clsid_menu);

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
        delete_reg_tree(HKEY_CLASSES_ROOT, PROGID);

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

tar_item_stream::tar_item_stream(const std::shared_ptr<tar_info>& tar, tar_item& item) : item(item) {
    struct archive_entry* entry;

    a = archive_read_new();

    try {
        int r;
        bool found = false;

        archive_read_support_filter_all(a);
        archive_read_support_format_all(a);

        r = archive_read_open_filename(a, (char*)tar->archive_fn.u8string().c_str(), BLOCK_SIZE);

        if (r != ARCHIVE_OK)
            throw runtime_error(archive_error_string(a));

        while (archive_read_next_header(a, &entry) == ARCHIVE_OK) {
            string_view name = archive_entry_pathname(entry);

            if (name == item.full_path) {
                found = true;
                break;
            }
        }

        if (!found)
            throw formatted_error("Could not find {} in archive.", item.full_path);
    } catch (...) {
        archive_read_free(a);
        throw;
    }
}

tar_item_stream::~tar_item_stream() {
    if (a)
        archive_read_free(a);
}

HRESULT tar_item_stream::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_ISequentialStream || iid == IID_IStream)
        *ppv = static_cast<IStream*>(this);
    else {
        debug("tar_item_stream::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG tar_item_stream::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG tar_item_stream::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT tar_item_stream::Read(void* pv, ULONG cb, ULONG* pcbRead) {
    int r;
    size_t size, copy_size;
    int64_t offset;
    const void* readbuf;

    if (item.dir)
        return E_NOTIMPL;

    *pcbRead = 0;

    if (!buf.empty()) {
        size_t copy_size = min(buf.size(), (size_t)cb);

        if (copy_size > 0) {
            memcpy(pv, buf.data(), copy_size);
            buf = buf.substr(copy_size);

            cb -= copy_size;
            *pcbRead += copy_size;
            pv = (uint8_t*)pv + copy_size;
        }
    }

    if (cb == 0)
        return S_OK;

    r = archive_read_data_block(a, &readbuf, &size, &offset);

    if (r != ARCHIVE_OK && r != ARCHIVE_EOF)
        throw runtime_error(archive_error_string(a));

    if (size == 0)
        return S_OK;

    copy_size = min(size, (size_t)cb);

    memcpy(pv, readbuf, copy_size);
    *pcbRead += copy_size;

    if (size > cb)
        buf.append(string_view((char*)readbuf + cb, size - cb));

    return S_OK;
}

void tar_item_stream::extract_file(const filesystem::path& fn) {
    HRESULT hr;
    char buf[BLOCK_SIZE];
    ULONG read, total_read = 0;

    unique_handle h{CreateFileW((LPCWSTR)fn.u16string().c_str(), GENERIC_WRITE, 0, nullptr,
                                CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr)};

    if (h.get() == INVALID_HANDLE_VALUE)
        throw last_error("CreateFile", GetLastError());

    while (total_read < item.size) {
        DWORD written;

        hr = Read(buf, BLOCK_SIZE, &read);

        if (FAILED(hr))
            throw formatted_error("tar_item_stream::Read returned {:08x}.", hr);

        if (read == 0)
            throw formatted_error("tar_item_stream::Read returned 0 bytes.");

        if (!WriteFile(h.get(), buf, read, &written, nullptr))
            throw last_error("WriteFile", GetLastError());

        total_read += read;
    }
}

HRESULT tar_item_stream::Write(const void* pv, ULONG cb, ULONG* pcbWritten) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::SetSize(ULARGE_INTEGER libNewSize) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::Commit(DWORD grfCommitFlags) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::Revert() {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) {
    UNIMPLEMENTED; // FIXME
}

HRESULT tar_item_stream::Stat(STATSTG* pstatstg, DWORD grfStatFlag) {
    debug("tar_item_stream::Stat({}, {:#x})", (void*)pstatstg, grfStatFlag);

    if (!(grfStatFlag & STATFLAG_NONAME)) {
        auto name = utf8_to_utf16(item.name);

        tar_item* p = item.parent;
        while (p) {
            if (!p->name.empty())
                name = utf8_to_utf16(p->name) + u"\\" + name;

            p = p->parent;
        }

        pstatstg->pwcsName = (WCHAR*)CoTaskMemAlloc((name.length() + 1) * sizeof(char16_t));
        memcpy(pstatstg->pwcsName, name.c_str(), (name.length() + 1) * sizeof(char16_t));
    } else
        pstatstg->pwcsName = nullptr;

    pstatstg->type = STGTY_STREAM;
    pstatstg->cbSize.QuadPart = item.size;
    pstatstg->mtime.dwLowDateTime = 0; // FIXME
    pstatstg->mtime.dwHighDateTime = 0; // FIXME
    pstatstg->ctime.dwLowDateTime = 0;
    pstatstg->ctime.dwHighDateTime = 0;
    pstatstg->atime.dwLowDateTime = 0;
    pstatstg->atime.dwHighDateTime = 0;
    pstatstg->grfMode = STGM_READ;
    pstatstg->grfLocksSupported = LOCK_WRITE;
    pstatstg->clsid = CLSID_NULL;
    pstatstg->grfStateBits = 0;

    return S_OK;
}

HRESULT tar_item_stream::Clone(IStream** ppstm) {
    UNIMPLEMENTED; // FIXME
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
