#include "tarfldr.h"
#include "resource.h"
#include <strsafe.h>

using namespace std;

#define HIDA_GetPIDLFolder(pida) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[0])
#define HIDA_GetPIDLItem(pida, i) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[i+1])

HRESULT shell_context_menu::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
    else if (iid == IID_IShellExtInit)
        *ppv = static_cast<IShellExtInit*>(this);
    else {
        debug("shell_context_menu::QueryInterface: unsupported interface {}", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_context_menu::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_context_menu::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_context_menu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
    UINT cmd = idCmdFirst;
    MENUITEMINFOW mii;

    debug("shell_context_menu::QueryContextMenu({}, {}, {}, {}, {:#x})", (void*)hmenu, indexMenu, idCmdFirst,
          idCmdLast, uFlags);

    if (!(uFlags & CMF_DEFAULTONLY)) {
        for (unsigned int i = 0; i < items.size(); i++) {
            WCHAR buf[256];

            mii.cbSize = sizeof(mii);
            mii.wID = cmd;

            if (LoadStringW(instance, items[i].res_num, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
                return E_FAIL;

            mii.fMask = MIIM_FTYPE | MIIM_ID | MIIM_STRING;
            mii.fType = MFT_STRING;
            mii.dwTypeData = buf;

            InsertMenuItemW(hmenu, indexMenu + i, true, &mii);
            cmd++;
        }
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

void shell_context_menu::extract_all(CMINVOKECOMMANDINFO* pici) {
    try {
        HRESULT hr;
        BROWSEINFOW bi;
        WCHAR msg[256], buf[MAX_PATH];
        com_object<IShellItem> dest;
        com_object<IFileOperation> ifo;

        // FIXME - can we preserve LXSS metadata when extracting?

        if (LoadStringW(instance, IDS_EXTRACT_TEXT, msg, sizeof(msg) / sizeof(WCHAR)) <= 0)
            throw last_error("LoadString", GetLastError());

        bi.hwndOwner = pici->hwnd;
        bi.pidlRoot = nullptr;
        bi.pszDisplayName = buf;
        bi.lpszTitle = msg;
        bi.ulFlags = 0;
        bi.lpfn = nullptr;

        auto dest_pidl = SHBrowseForFolderW(&bi);

        if (!dest_pidl)
            return;

        {
            IShellItem* si;

            hr = SHCreateItemFromIDList(dest_pidl, IID_IShellItem, (void**)&si);
            if (FAILED(hr)) {
                ILFree(dest_pidl);
                throw formatted_error("SHCreateItemFromIDList returned {:08x}.", (uint32_t)hr);
            }

            dest.reset(si);
        }

        ILFree(dest_pidl);

        {
            IFileOperation* fo;

            hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_ALL, IID_IFileOperation, (void**)&fo);
            if (FAILED(hr))
                throw formatted_error("CoCreateInstance returned {:08x} for CLSID_FileOperation.", (uint32_t)hr);

            ifo.reset(fo);
        }

        vector<shell_item> shell_items;

        for (const auto& file : files) {
            if (get<1>(file) & archive_type::tarball) {
                vector<tar_item*> itemlist;
                WCHAR path[MAX_PATH];

                if (!SHGetPathFromIDListW((ITEMIDLIST*)get<0>(file).data(), path))
                    throw runtime_error("SHGetPathFromIDList failed");

                shared_ptr<tar_info> ti{new tar_info(path)};

                for (auto& item : ti->root.children) {
                    itemlist.push_back(&item);
                }

                shell_items.emplace_back((ITEMIDLIST*)get<0>(file).data(), ti, itemlist, &ti->root, false);
            }
        }

        for (auto& si : shell_items) {
            IUnknown* unk;

            hr = si.QueryInterface(IID_IUnknown, (void**)&unk);
            if (FAILED(hr))
                throw formatted_error("shell_item::QueryInterface returned {:08x}.", (uint32_t)hr);

            hr = ifo->CopyItems(unk, dest.get());
            if (FAILED(hr))
                throw formatted_error("IFileOperation::CopyItems returned {:08x}.", (uint32_t)hr);
        }

        hr = ifo->PerformOperations();
        if (FAILED(hr))
            throw formatted_error("IFileOperation::PerformOperations returned {:08x}.", (uint32_t)hr);
    } catch (const exception& e) {
        MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
    }
}

static void decompress_file(ITEMIDLIST* pidl, archive_type type) {
    HRESULT hr;
    com_object<IShellItem> isi;
    com_object<IStream> stream;
    u16string orig_fn, new_fn;
    FILETIME creation_time, access_time, write_time;

    {
        WCHAR buf[MAX_PATH];

        if (!SHGetPathFromIDListW(pidl, (WCHAR*)buf))
            throw runtime_error("SHGetPathFromIDList failed");

        orig_fn = new_fn = (char16_t*)buf;
    }

    {
        unique_handle h{CreateFileW((LPCWSTR)orig_fn.c_str(), READ_ATTRIBUTES, 0, nullptr,
                                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)};

        if (h.get() == INVALID_HANDLE_VALUE)
            throw last_error("CreateFile", GetLastError());

        if (!GetFileTime(h.get(), &creation_time, &access_time, &write_time))
            throw last_error("GetFileTime", GetLastError());
    }

    {
        IShellItem* tmp;

        hr = SHCreateItemFromIDList(pidl, IID_IShellItem, (void**)&tmp);
        if (FAILED(hr))
            throw formatted_error("SHCreateItemFromIDList returned {:08x}.", (uint32_t)hr);

        isi.reset(tmp);
    }

    {
        IStream* tmp;

        hr = isi->BindToHandler(nullptr, BHID_Stream, IID_IStream, (void**)&tmp);
        if (FAILED(hr))
            throw formatted_error("IShellItem::BindToHandler returned {:08x}.", (uint32_t)hr);

        stream.reset(tmp);
    }

    auto st = new_fn.rfind(u".");

    if (st == string::npos)
        throw runtime_error("Could not find file extension.");

    new_fn = new_fn.substr(0, st);

    unique_handle h{CreateFileW((LPCWSTR)new_fn.c_str(), GENERIC_WRITE, 0, nullptr,
                    CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr)};

    if (h.get() == INVALID_HANDLE_VALUE)
        throw last_error("CreateFile", GetLastError());

    try {
        if (type & archive_type::gzip) {
            int ret;
            z_stream strm;
            uint8_t inbuf[4096], outbuf[4096];

            // FIXME - can we do this via IStream rather than CreateFile etc.?

            strm.zalloc = Z_NULL;
            strm.zfree = Z_NULL;
            strm.opaque = Z_NULL;

            ret = inflateInit2(&strm, 16 + MAX_WBITS);
            if (ret != Z_OK)
                throw formatted_error("inflateInit2 returned {}.", ret);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = outbuf;
            strm.avail_out = sizeof(outbuf);

            while (true) {
                if (strm.avail_in == 0) {
                    ULONG read;

                    strm.next_in = inbuf;

                    hr = stream->Read(inbuf, sizeof(inbuf), &read);
                    if (FAILED(hr))
                        throw formatted_error("IStream::Read returned {:08x}.", (uint32_t)hr);

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        break;
                }

                ret = inflate(&strm, Z_NO_FLUSH);
                if (ret != Z_OK && ret != Z_STREAM_END)
                    throw formatted_error("inflate returned {}.", ret);

                if (strm.avail_out == 0 || ret == Z_STREAM_END) {
                    DWORD written;

                    if (!WriteFile(h.get(), outbuf, sizeof(outbuf) - strm.avail_out, &written, nullptr))
                        throw last_error("WriteFile", GetLastError());
                }

                if (strm.avail_out == 0) {
                    strm.next_out = outbuf;
                    strm.avail_out = sizeof(outbuf);
                }

                if (ret == Z_STREAM_END)
                    break;
            }
        } else if (type & archive_type::bz2) {
            int ret;
            bz_stream strm;
            char inbuf[4096], outbuf[4096];

            strm.bzalloc = nullptr;
            strm.bzfree = nullptr;
            strm.opaque = nullptr;

            ret = BZ2_bzDecompressInit(&strm, 0, 0);
            if (ret != BZ_OK)
                throw formatted_error("BZ2_bzDecompressInit returned {}.", ret);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = outbuf;
            strm.avail_out = sizeof(outbuf);

            while (true) {
                if (strm.avail_in == 0) {
                    ULONG read;

                    strm.next_in = inbuf;

                    hr = stream->Read(inbuf, sizeof(inbuf), &read);
                    if (FAILED(hr))
                        throw formatted_error("IStream::Read returned {:08x}.", (uint32_t)hr);

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        break;
                }

                ret = BZ2_bzDecompress(&strm);
                if (ret != BZ_OK && ret != BZ_STREAM_END)
                    throw formatted_error("BZ2_bzDecompress returned {}.", ret);

                if (strm.avail_out == 0 || ret == BZ_STREAM_END) {
                    DWORD written;

                    if (!WriteFile(h.get(), outbuf, sizeof(outbuf) - strm.avail_out, &written, nullptr))
                        throw last_error("WriteFile", GetLastError());
                }

                if (strm.avail_out == 0) {
                    strm.next_out = outbuf;
                    strm.avail_out = sizeof(outbuf);
                }

                if (ret == BZ_STREAM_END)
                    break;
            }
        } else if (type & archive_type::xz) {
            int ret;
            lzma_stream strm = LZMA_STREAM_INIT;
            uint8_t inbuf[4096], outbuf[4096];

            ret = lzma_stream_decoder(&strm, UINT64_MAX, 0);
            if (ret != LZMA_OK)
                throw formatted_error("lzma_stream_decoder returned {}.", ret);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = outbuf;
            strm.avail_out = sizeof(outbuf);

            while (true) {
                if (strm.avail_in == 0) {
                    ULONG read;

                    strm.next_in = inbuf;

                    hr = stream->Read(inbuf, sizeof(inbuf), &read);
                    if (FAILED(hr))
                        throw formatted_error("IStream::Read returned {:08x}.", (uint32_t)hr);

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        break;
                }

                ret = lzma_code(&strm, LZMA_RUN);
                if (ret != LZMA_OK && ret != LZMA_STREAM_END)
                    throw formatted_error("lzma_code returned {}.", ret);

                if (strm.avail_out == 0 || ret == LZMA_STREAM_END) {
                    DWORD written;

                    if (!WriteFile(h.get(), outbuf, sizeof(outbuf) - strm.avail_out, &written, nullptr))
                        throw last_error("WriteFile", GetLastError());
                }

                if (strm.avail_out == 0) {
                    strm.next_out = outbuf;
                    strm.avail_out = sizeof(outbuf);
                }

                if (ret == LZMA_STREAM_END)
                    break;
            }
        }

        stream.reset(); // close IStream

        // change times to those of original file

        if (!SetFileTime(h.get(), &creation_time, &access_time, &write_time))
            throw last_error("SetFileTime", GetLastError());

        // FIXME - copy ADSes, extended attributes, and SD?
    } catch (...) {
        h.reset();
        DeleteFileW((WCHAR*)new_fn.c_str());

        throw;
    }

    DeleteFileW((WCHAR*)orig_fn.c_str());
}

void shell_context_menu::decompress(CMINVOKECOMMANDINFO* pici) {
    try {
        for (const auto& file : files) {
            if (get<1>(file) & archive_type::gzip || get<1>(file) & archive_type::bz2 || get<1>(file) & archive_type::xz)
                decompress_file((ITEMIDLIST*)get<0>(file).data(), get<1>(file));
        }
    } catch (const exception& e) {
        MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
    }
}

static void compress_file(ITEMIDLIST* pidl, archive_type type) {
    HRESULT hr;
    com_object<IShellItem> isi;
    com_object<IStream> stream;
    u16string orig_fn, new_fn;
    FILETIME creation_time, access_time, write_time;

    {
        WCHAR buf[MAX_PATH];

        if (!SHGetPathFromIDListW(pidl, (WCHAR*)buf))
            throw runtime_error("SHGetPathFromIDList failed");

        orig_fn = new_fn = (char16_t*)buf;
    }

    {
        unique_handle h{CreateFileW((LPCWSTR)orig_fn.c_str(), READ_ATTRIBUTES, 0, nullptr,
                                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)};

        if (h.get() == INVALID_HANDLE_VALUE)
            throw last_error("CreateFile", GetLastError());

        if (!GetFileTime(h.get(), &creation_time, &access_time, &write_time))
            throw last_error("GetFileTime", GetLastError());
    }

    {
        IShellItem* tmp;

        hr = SHCreateItemFromIDList(pidl, IID_IShellItem, (void**)&tmp);
        if (FAILED(hr))
            throw formatted_error("SHCreateItemFromIDList returned {:08x}.", (uint32_t)hr);

        isi.reset(tmp);
    }

    {
        IStream* tmp;

        hr = isi->BindToHandler(nullptr, BHID_Stream, IID_IStream, (void**)&tmp);
        if (FAILED(hr))
            throw formatted_error("IShellItem::BindToHandler returned {:08x}.", (uint32_t)hr);

        stream.reset(tmp);
    }

    new_fn += u".bz2";

    unique_handle h{CreateFileW((LPCWSTR)new_fn.c_str(), GENERIC_WRITE, 0, nullptr,
                    CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr)};

    if (h.get() == INVALID_HANDLE_VALUE)
        throw last_error("CreateFile", GetLastError());

    try {
        if (type & archive_type::bz2) {
            int ret;
            bz_stream strm;
            char inbuf[4096], outbuf[4096];
            bool eof = false;

            // FIXME - can we do this via IStream rather than CreateFile etc.?

            strm.bzalloc = nullptr;
            strm.bzfree = nullptr;
            strm.opaque = nullptr;

            ret = BZ2_bzCompressInit(&strm, 9, 0, 30);
            if (ret != BZ_OK)
                throw formatted_error("BZ2_bzCompressInit returned {}.", ret);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = outbuf;
            strm.avail_out = sizeof(outbuf);

            do {
                if (strm.avail_in == 0 && !eof) {
                    ULONG read;

                    strm.next_in = inbuf;

                    hr = stream->Read(inbuf, sizeof(inbuf), &read);
                    if (FAILED(hr))
                        throw formatted_error("IStream::Read returned {:08x}.", (uint32_t)hr);

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        eof = true;
                }

                ret = BZ2_bzCompress(&strm, eof ? BZ_FINISH : BZ_RUN);
                if (ret != BZ_RUN_OK && ret != BZ_FINISH_OK && ret != BZ_STREAM_END)
                    throw formatted_error("BZ2_bzCompress returned {}.", ret);

                if (strm.avail_out == 0 || eof) {
                    DWORD written;

                    if (!WriteFile(h.get(), outbuf, sizeof(outbuf) - strm.avail_out, &written, nullptr))
                        throw last_error("WriteFile", GetLastError());
                }

                if (strm.avail_out == 0) {
                    strm.next_out = outbuf;
                    strm.avail_out = sizeof(outbuf);
                }
            } while (ret != BZ_STREAM_END);
        }

        // FIXME - gzip
        // FIXME - xz

        stream.reset(); // close IStream

        // change times to those of original file

        if (!SetFileTime(h.get(), &creation_time, &access_time, &write_time))
            throw last_error("SetFileTime", GetLastError());

        // FIXME - copy ADSes, extended attributes, and SD?
    } catch (...) {
        h.reset();
        DeleteFileW((WCHAR*)new_fn.c_str());

        throw;
    }

    DeleteFileW((WCHAR*)orig_fn.c_str());
}

void shell_context_menu::compress(CMINVOKECOMMANDINFO* pici) {
    try {
        for (const auto& file : files) {
            if (!(get<1>(file) & archive_type::gzip || get<1>(file) & archive_type::bz2 || get<1>(file) & archive_type::xz))
                compress_file((ITEMIDLIST*)get<0>(file).data(), archive_type::bz2); // FIXME - type selection
        }
    } catch (const exception& e) {
        MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
    }
}

HRESULT shell_context_menu::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_context_menu::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    if (IS_INTRESOURCE(pici->lpVerb)) {
        if ((uintptr_t)pici->lpVerb >= items.size())
            return E_INVALIDARG;

        items[(uintptr_t)pici->lpVerb].cmd(this, pici);

        return S_OK;
    }

    for (const auto& item : items) {
        if (item.verba == pici->lpVerb) {
            item.cmd(this, pici);

            return S_OK;
        }
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax) {
    debug("shell_context_menu::GetCommandString({}, {}, {}, {}, {})", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    if (!IS_INTRESOURCE(idCmd)) {
        bool found = false;

        for (unsigned int i = 0; i < items.size(); i++) {
            if (items[i].verba == (char*)idCmd) {
                idCmd = i;
                found = true;
                break;
            }
        }

        if (!found)
            return E_INVALIDARG;
    } else {
        if (idCmd >= items.size())
            return E_INVALIDARG;
    }

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            return StringCchCopyA(pszName, cchMax, items[idCmd].verba.c_str());

        case GCS_VERBW:
            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)items[idCmd].verbw.c_str());
    }

    return E_INVALIDARG;
}

HRESULT shell_context_menu::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject* pdtobj, HKEY hkeyProgID) {
    CLIPFORMAT cf = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
    FORMATETC format = { cf, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    HRESULT hr;
    UINT num_files;
    CIDA* cida;
    STGMEDIUM stgm;
    WCHAR path[MAX_PATH];
    bool show_extract_all = false, show_decompress = false, show_compress = false;

    if (pidlFolder || !pdtobj)
        return E_INVALIDARG;

    stgm.tymed = TYMED_HGLOBAL;

    hr = pdtobj->GetData(&format, &stgm);
    if (FAILED(hr))
        return hr;

    cida = (CIDA*)GlobalLock(stgm.hGlobal);

    if (!cida) {
        ReleaseStgMedium(&stgm);
        return E_INVALIDARG;
    }

    for (unsigned int i = 0; i < cida->cidl; i++) {
        auto pidl = ILCombine(HIDA_GetPIDLFolder(cida), HIDA_GetPIDLItem(cida, i));

        files.emplace_back(string_view((char*)pidl, ILGetSize(pidl)), archive_type::unknown);

        ILFree(pidl);
    }

    GlobalUnlock(stgm.hGlobal);
    ReleaseStgMedium(&stgm);

    // get names of files

    vector<const ITEMIDLIST*> pidls;
    pidls.reserve(files.size());

    for (const auto& f : files) {
        pidls.push_back((ITEMIDLIST*)get<0>(f).data());
    }

    com_object<IShellItemArray> isia;

    {
        IShellItemArray* tmp;

        hr = SHCreateShellItemArrayFromIDLists(pidls.size(), pidls.data(), &tmp);
        if (FAILED(hr))
            return hr;

        isia.reset(tmp);
    }

    for (unsigned int i = 0; i < files.size(); i++) {
        com_object<IShellItem> isi;
        WCHAR* buf;

        {
            IShellItem* tmp;
            hr = isia->GetItemAt(i, &tmp);
            if (FAILED(hr))
                return hr;

            isi.reset(tmp);
        }

        hr = isi->GetDisplayName(SIGDN_NORMALDISPLAY, &buf);
        if (FAILED(hr))
            return hr;

        auto type = get<1>(files[i]) = identify_file_type((char16_t*)buf);

        if (type & archive_type::tarball)
            show_extract_all = true;

        if (type & archive_type::gzip || type & archive_type::bz2 || type & archive_type::xz)
            show_decompress = true;
        else
            show_compress = true;

        CoTaskMemFree(buf);
    }

    if (show_extract_all)
        items.emplace_back(IDS_EXTRACT_ALL, "extract", u"extract", shell_context_menu::extract_all);

    if (show_decompress)
        items.emplace_back(IDS_DECOMPRESS, "decompress", u"decompress", shell_context_menu::decompress);

    if (show_compress)
        items.emplace_back(IDS_COMPRESS, "compress", u"compress", shell_context_menu::compress);

    return S_OK;
}
