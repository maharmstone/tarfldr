#include "tarfldr.h"

using namespace std;

#define BLOCK_SIZE 20480

tar_item_stream::~tar_item_stream() {
    if (a)
        archive_read_free(a);

    if (gzf)
        gzclose(gzf);

    if (bzf)
        BZ2_bzclose(bzf);
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

    if (type & archive_type::tarball) {
        auto r = archive_read_data_block(a, &readbuf, &size, &offset);

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

    switch (type) {
        case archive_type::gzip: {
            auto ret = gzread(gzf, pv, cb);

            if (ret < 0)
                throw formatted_error("gzread returned {}.", ret); // FIXME - use gzerror to get actual error

            *pcbRead = ret;

            break;
        }

        case archive_type::bz2: {
            auto ret = BZ2_bzread(bzf, pv, cb);

            if (ret < 0)
                throw formatted_error("BZ2_bzread returned {}.", ret); // FIXME - use BZ2_bzerror to get actual error

            *pcbRead = ret;

            break;
        }

        case archive_type::xz: {
            while (true) {
                if (strm.avail_in == 0) {
                    DWORD read;

                    strm.next_in = (uint8_t*)lzma_inbuf.data();

                    if (!ReadFile(h.get(), lzma_inbuf.data(), lzma_inbuf.length(), &read, nullptr))
                        throw last_error("ReadFile", GetLastError());

                    strm.avail_in = read;

                    if (read == 0) // end of file
                        break;
                }

                auto ret = lzma_code(&strm, LZMA_RUN);

                if (ret != LZMA_OK && ret != LZMA_STREAM_END)
                    throw formatted_error("lzma_code returned {}.", ret);

                if (strm.avail_out == 0 || ret == LZMA_STREAM_END) {
                    size_t read_size = lzma_outbuf.length() - strm.avail_out;

                    copy_size = min(read_size, (size_t)cb);

                    memcpy(pv, lzma_outbuf.data(), copy_size);

                    pv = (uint8_t*)pv + copy_size;
                    cb -= copy_size;
                    *pcbRead += copy_size;

                    if (read_size > cb)
                        buf.append(string_view(lzma_outbuf).substr(cb, read_size - cb));
                }

                if (strm.avail_out == 0) {
                    strm.next_out = (uint8_t*)lzma_outbuf.data();
                    strm.avail_out = lzma_outbuf.length();
                }

                if (ret == LZMA_STREAM_END)
                    break;
            }

            break;
        }
    }

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

tar_item_stream::tar_item_stream(const std::shared_ptr<tar_info>& tar, tar_item& item) : item(item), type(tar->type) {
    if (tar->type & archive_type::tarball) {
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

        return;
    }

    switch (tar->type) {
        case archive_type::gzip: {
            gzf = gzopen((char*)tar->archive_fn.u8string().c_str(), "r");

            if (!gzf)
                throw formatted_error("Could not open gzip file {}.", tar->archive_fn.string());

            break;
        }

        case archive_type::bz2: {
            bzf = BZ2_bzopen((char*)tar->archive_fn.u8string().c_str(), "r");

            if (!bzf)
                throw formatted_error("Could not open bzip2 file {}.", tar->archive_fn.string());

            break;
        }

        case archive_type::xz: {
            h.reset(CreateFileW((LPCWSTR)tar->archive_fn.u16string().c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));

            if (h.get() == INVALID_HANDLE_VALUE)
                throw last_error("CreateFile", GetLastError());

            auto ret = lzma_stream_decoder(&strm, UINT64_MAX, 0);

            if (ret != LZMA_OK)
                throw formatted_error("lzma_stream_decoder returned {}.", ret);

            lzma_inbuf.resize(4096);
            lzma_outbuf.resize(4096);

            strm.next_in = nullptr;
            strm.avail_in = 0;
            strm.next_out = (uint8_t*)lzma_outbuf.data();
            strm.avail_out = lzma_outbuf.length();

            break;
        }

        default:
            throw runtime_error("FIXME - unsupported archive type"); // FIXME
    }
}
