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

#include "tarfldr.h"
#include "resource.h"
#include <shlwapi.h>
#include <strsafe.h>
#include <shlobj.h>
#include <ntquery.h>
#include <span>
#include <functional>

using namespace std;

static const struct {
    unsigned int res_num;
    const char* verba;
    const char16_t* verbw;
    function<HRESULT(shell_item_list*, CMINVOKECOMMANDINFO*)> cmd;
} menu_items[] = {
    { IDS_OPEN, "open", u"open", &shell_item_list::open_cmd },
    { 0, nullptr, nullptr, nullptr },
    { IDS_COPY, "copy", u"copy", &shell_item_list::copy_cmd },
    { 0, nullptr, nullptr, nullptr },
    { IDS_PROPERTIES, "properties", u"properties", &shell_item_list::properties }
};

// FIXME - others: Extract, Cut, Paste, Properties

shell_item_list::shell_item_list(PIDLIST_ABSOLUTE root_pidl, const shared_ptr<tar_info>& tar,
                                 const vector<tar_item*>& itemlist, tar_item* root, bool recursive, shell_folder* folder) :
                                 tar(tar), itemlist(itemlist), root(root), recursive(recursive), folder(folder) {
    this->root_pidl = ILCloneFull(root_pidl);
    cf_shell_id_list = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
    cf_file_contents = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    cf_file_descriptor = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);

    if (folder)
        folder->AddRef();
}

shell_item_list::~shell_item_list() {
    if (root_pidl)
        ILFree(root_pidl);

    if (folder)
        folder->Release();
}

HRESULT shell_item_list::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IContextMenu)
        *ppv = static_cast<IContextMenu*>(this);
    else if (iid == IID_IDataObject)
        *ppv = static_cast<IDataObject*>(this);
    else {
        if (iid == IID_IStdMarshalInfo)
            debug("shell_item_list::QueryInterface: unsupported interface IID_IStdMarshalInfo\n");
        else if (iid == IID_INoMarshal)
            debug("shell_item_list::QueryInterface: unsupported interface IID_INoMarshal\n");
        else if (iid == IID_IAgileObject)
            debug("shell_item_list::QueryInterface: unsupported interface IID_IAgileObject\n");
        else if (iid == IID_ICallFactory)
            debug("shell_item_list::QueryInterface: unsupported interface IID_ICallFactory\n");
        else if (iid == IID_IExternalConnection)
            debug("shell_item_list::QueryInterface: unsupported interface IID_IExternalConnection\n");
        else if (iid == IID_IMarshal)
            debug("shell_item_list::QueryInterface: unsupported interface IID_IMarshal\n");
        else if (iid == IID_IDataObjectAsyncCapability)
            debug("shell_item_list::QueryInterface: unsupported interface IID_IDataObjectAsyncCapability\n");
        else
            debug("shell_item_list::QueryInterface: unsupported interface {}\n", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_item_list::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_item_list::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_item_list::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst,
                                          UINT idCmdLast, UINT uFlags) {
    UINT cmd = idCmdFirst;
    MENUITEMINFOW mii;

    debug("shell_item_list::QueryContextMenu({}, {}, {}, {}, {:#x})\n", (void*)hmenu, indexMenu, idCmdFirst,
          idCmdLast, uFlags);

    span mi = menu_items;

    for (unsigned int i = 0; i < mi.size(); i++) {
        mii.cbSize = sizeof(mii);
        mii.wID = cmd;

        if (mi[i].res_num != 0) {
            WCHAR buf[256];

            if (LoadStringW(instance, mi[i].res_num, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
                return E_FAIL;

            mii.fMask = MIIM_FTYPE | MIIM_STATE | MIIM_ID | MIIM_STRING;
            mii.fType = MFT_STRING;
            mii.fState = i == 0 ? MFS_DEFAULT : 0;
            mii.dwTypeData = buf;
        } else {
            mii.fMask = MIIM_FTYPE | MIIM_ID;
            mii.fType = MFT_SEPARATOR;
        }

        InsertMenuItemW(hmenu, indexMenu + i, true, &mii);
        cmd++;

        if (uFlags & CMF_DEFAULTONLY)
            break;
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, cmd - idCmdFirst);
}

static filesystem::path get_temp_file_name(const filesystem::path& dir, const u16string& prefix, unsigned int unique) {
    WCHAR tmpfn[MAX_PATH];

    if (GetTempFileNameW((WCHAR*)dir.u16string().c_str(), (WCHAR*)prefix.c_str(), unique, tmpfn) == 0)
        throw runtime_error("GetTempFileName failed.");

    return tmpfn;
}

HRESULT shell_item_list::open_cmd(CMINVOKECOMMANDINFO* pici) {
    for (auto item : itemlist) {
        if (item->dir) {
            SHELLEXECUTEINFOW sei;

            auto child_pidl = item->make_pidl_child();
            auto pidl = ILCombine(root_pidl, child_pidl);

            ILFree(child_pidl);

            memset(&sei, 0, sizeof(sei));
            sei.cbSize = sizeof(sei);
            sei.fMask = SEE_MASK_IDLIST | SEE_MASK_CLASSNAME;
            sei.lpIDList = pidl;
            sei.lpClass = L"Folder";
            sei.hwnd = pici->hwnd;
            sei.nShow = SW_SHOWNORMAL;
            sei.lpVerb = L"open";
            ShellExecuteExW(&sei);

            ILFree(pidl);
        } else {
            try {
                WCHAR temp_path[MAX_PATH];

                if (GetTempPathW(sizeof(temp_path) / sizeof(WCHAR), temp_path) == 0)
                    throw last_error("GetTempPath", GetLastError());

                filesystem::path fn = get_temp_file_name(temp_path, u"tar", 0);

                // replace extension with original one

                auto st = item->full_path.rfind(".");
                if (st != string::npos) {
                    string_view ext = string_view(item->full_path).substr(st + 1);
                    fn.replace_extension(ext);
                }

                {
                    tar_item_stream tis(tar, *item);

                    tis.extract_file(fn);
                }

                // open using normal handler

                auto ret = (intptr_t)ShellExecuteW(pici->hwnd, L"open", (WCHAR*)fn.u16string().c_str(), nullptr,
                                                   nullptr, SW_SHOW);
                if (ret <= 32)
                    throw formatted_error("ShellExecute returned {}.", ret);
            } catch (const exception& e) {
                MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
                return E_FAIL;
            }
        }
    }

    return S_OK;
}

HRESULT shell_item_list::copy_cmd(CMINVOKECOMMANDINFO* pici) {
    HRESULT hr;
    IDataObject* dataobj;

    hr = QueryInterface(IID_IDataObject, (void**)&dataobj);
    if (FAILED(hr))
        return hr;

    OleSetClipboard(dataobj);

    dataobj->Release();

    return S_OK;
}

u16string shell_item_list::get_item_prop(tar_item& item, const GUID& fmtid, DWORD pid) {
    HRESULT hr;
    u16string val;
    SHCOLUMNID scid;
    VARIANT v;
    auto pidl = item.make_pidl_child();

    VariantInit(&v);

    scid.fmtid = fmtid;
    scid.pid = pid;

    hr = folder->GetDetailsEx(pidl, &scid, &v);

    if (FAILED(hr)) {
        ILFree(pidl);
        return u"";
    }

    if (fmtid == FMTID_POSIXAttributes && pid == PID_POSIX_MODE && v.vt == VT_I4)
        val = mode_to_u16string((mode_t)v.lVal);
    else {
        hr = VariantChangeType(&v, &v, 0, VT_BSTR);

        if (FAILED(hr)) {
            ILFree(pidl);
            VariantClear(&v);
            return u"";
        }

        val = (char16_t*)v.bstrVal;
    }

    VariantClear(&v);

    ILFree(pidl);

    return val;
}

static uint64_t calc_size_dir(const tar_item& item) {
    uint64_t size = 0;

    for (const auto& c : item.children) {
        if (c.dir)
            size += calc_size_dir(c);
        else
            size += c.size;
    }

    return size;
}

uint64_t shell_item_list::calc_size() {
    uint64_t size = 0;

    for (auto item : itemlist) {
        if (item->dir)
            size += calc_size_dir(*item);
        else
            size += item->size;
    }

    return size;
}

INT_PTR shell_item_list::PropSheetDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    try {
        switch (uMsg) {
            case WM_INITDIALOG: {
                HRESULT hr;
                u16string multiple, type, modified, user, group, mode, location;
                uint64_t size = calc_size();

                if (itemlist.size() == 1) {
                    IExtractIconW* ieiw;
                    ITEMIDLIST* apidl[1];

                    apidl[0] = itemlist[0]->make_relative_pidl(root);

                    hr = folder->GetUIObjectOf(hwndDlg, 1, (PCUITEMID_CHILD_ARRAY)apidl, IID_IExtractIconW, nullptr, (void**)&ieiw);
                    if (!FAILED(hr)) {
                        WCHAR buf[MAX_PATH];
                        int index;
                        UINT flags;

                        hr = ieiw->GetIconLocation(0, buf, sizeof(buf) / sizeof(WCHAR), &index, &flags);
                        if (!FAILED(hr)) {
                            HICON icon;

                            hr = ieiw->Extract(buf, index, &icon, nullptr, 32); // FIXME - size?

                            if (!FAILED(hr))
                                SendDlgItemMessageW(hwndDlg, IDC_FILE_ICON, STM_SETICON, (WPARAM)icon, 0);
                        }

                        ieiw->Release();
                    }

                    ILFree(apidl[0]);
                } else { // use icon for multiple files
                    HICON icon = LoadIconW(LoadLibraryW(L"shell32.dll"), MAKEINTRESOURCE(133)); // FIXME - is this constant defined anywhere?

                    if (icon)
                        SendDlgItemMessageW(hwndDlg, IDC_FILE_ICON, STM_SETICON, (WPARAM)icon, 0);
                }

                if (itemlist.size() > 1) {
                    char16_t buf[255];

                    if (LoadStringW(instance, IDS_MULTIPLE, (WCHAR*)buf, sizeof(buf) / sizeof(char16_t)) <= 0)
                        throw runtime_error("LoadString failed.");

                    multiple = buf;
                }

                if (itemlist.size() == 1)
                    SetDlgItemTextW(hwndDlg, IDC_FILE_NAME, (WCHAR*)utf8_to_utf16(itemlist[0]->name).c_str());
                else
                    SetDlgItemTextW(hwndDlg, IDC_FILE_NAME, (WCHAR*)multiple.c_str());

                for (unsigned int i = 0; i < itemlist.size(); i++) {
                    auto val = get_item_prop(*itemlist[i], FMTID_Storage, PID_STG_STORAGETYPE);

                    if (i == 0)
                        type = val;
                    else if (type != val) {
                        type = multiple;
                        break;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_FILE_TYPE, (WCHAR*)type.c_str());

                for (unsigned int i = 0; i < itemlist.size(); i++) {
                    auto val = get_item_prop(*itemlist[i], FMTID_Storage, PID_STG_WRITETIME);

                    if (i == 0)
                        modified = val;
                    else if (modified != val) {
                        modified = multiple;
                        break;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_MODIFIED, (WCHAR*)modified.c_str());

                if (root == &tar->root) {
                    char16_t buf[255];

                    if (LoadStringW(instance, IDS_ARCHIVE_ROOT, (WCHAR*)buf, sizeof(buf) / sizeof(char16_t)) <= 0)
                        throw runtime_error("LoadString failed.");

                    location = buf;
                } else {
                    tar_item* r = root;

                    while (r != &tar->root) {
                        location = utf8_to_utf16(r->name) + (location.empty() ? u"" : u"/") + location;

                        r = r->parent;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_LOCATION, (WCHAR*)location.c_str());

                {
                    char16_t sizestr[50];

                    sizestr[0] = 0;

                    StrFormatByteSizeW(size, (WCHAR*)sizestr, sizeof(sizestr) / sizeof(char16_t));

                    SetDlgItemTextW(hwndDlg, IDC_FILE_SIZE, (WCHAR*)sizestr);
                }

                for (unsigned int i = 0; i < itemlist.size(); i++) {
                    auto val = get_item_prop(*itemlist[i], FMTID_POSIXAttributes, PID_POSIX_USER);

                    if (i == 0)
                        user = val;
                    else if (user != val) {
                        user = multiple;
                        break;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_POSIX_USER, (WCHAR*)user.c_str());

                for (unsigned int i = 0; i < itemlist.size(); i++) {
                    auto val = get_item_prop(*itemlist[i], FMTID_POSIXAttributes, PID_POSIX_GROUP);

                    if (i == 0)
                        group = val;
                    else if (group != val) {
                        group = multiple;
                        break;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_POSIX_GROUP, (WCHAR*)group.c_str());

                for (unsigned int i = 0; i < itemlist.size(); i++) {
                    auto val = get_item_prop(*itemlist[i], FMTID_POSIXAttributes, PID_POSIX_MODE);

                    if (i == 0)
                        mode = val;
                    else if (mode != val) {
                        mode = multiple;
                        break;
                    }
                }

                SetDlgItemTextW(hwndDlg, IDC_POSIX_MODE, (WCHAR*)mode.c_str());

                break;
            }
        }
    } catch (const exception& e) {
        MessageBoxW(hwndDlg, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);
    }

    return false;
}

static INT_PTR __stdcall PropSheetDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    shell_item_list* si;

    if (uMsg == WM_INITDIALOG) {
        auto psp = (PROPSHEETPAGE*)lParam;

        si = (shell_item_list*)psp->lParam;

        SetWindowLongPtr(hwndDlg, GWLP_USERDATA, (LONG_PTR)si);
    } else
        si = (shell_item_list*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);

    return si->PropSheetDlgProc(hwndDlg, uMsg, wParam, lParam);
}

HRESULT shell_item_list::properties(CMINVOKECOMMANDINFO* pici) {
    try {
        PROPSHEETPAGEW psp;
        PROPSHEETHEADERW psh;
        u16string caption;

        {
            WCHAR buf[256];
            string name;

            if (LoadStringW(instance, IDS_PROPSHEET_WINDOW, buf, sizeof(buf) / sizeof(WCHAR)) <= 0)
                return E_FAIL;

            name = itemlist[0]->name;

            if (itemlist.size() > 1)
                name += ", ...";

            caption = utf8_to_utf16(fmt::format(utf16_to_utf8((char16_t*)buf), name));
        }

        psp.dwSize = sizeof(psp);
        psp.dwFlags = PSP_USETITLE;
        psp.hInstance = instance;
        psp.pszTemplate = MAKEINTRESOURCEW(IDD_PROPSHEET);
        psp.hIcon = 0;
        psp.pszTitle = MAKEINTRESOURCEW(IDS_PROPSHEET_TITLE);
        psp.pfnDlgProc = ::PropSheetDlgProc;
        psp.pfnCallback = nullptr;
        psp.lParam = (LPARAM)this;

        memset(&psh, 0, sizeof(PROPSHEETHEADERW));

        psh.dwSize = sizeof(PROPSHEETHEADERW);
        psh.dwFlags = PSH_PROPSHEETPAGE;
        psh.hwndParent = pici->hwnd;
        psh.hInstance = psp.hInstance;
        psh.pszCaption = (WCHAR*)caption.c_str();
        psh.nPages = 1;
        psh.ppsp = &psp;

        PropertySheetW(&psh);

        return S_OK;
    } catch (const exception& e) {
        MessageBoxW(pici->hwnd, (WCHAR*)utf8_to_utf16(e.what()).c_str(), L"Error", MB_ICONERROR);

        return E_FAIL;
    }
}

HRESULT shell_item_list::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici)
        return E_INVALIDARG;

    debug("shell_item_list::InvokeCommand(cbSize = {}, fMask = {:#x}, hwnd = {}, lpVerb = {}, lpParameters = {}, lpDirectory = {}, nShow = {}, dwHotKey = {}, hIcon = {})\n",
          pici->cbSize, pici->fMask, (void*)pici->hwnd, IS_INTRESOURCE(pici->lpVerb) ? to_string((uintptr_t)pici->lpVerb) : pici->lpVerb,
          pici->lpParameters ? pici->lpParameters : "NULL", pici->lpDirectory ? pici->lpDirectory : "NULL", pici->nShow, pici->dwHotKey,
          (void*)pici->hIcon);

    span mi = menu_items;

    if (IS_INTRESOURCE(pici->lpVerb)) {
        if ((uintptr_t)pici->lpVerb >= mi.size())
            return E_INVALIDARG;

        return mi[(uintptr_t)pici->lpVerb].cmd(this, pici);
    }

    for (const auto& mie : mi) {
        if (mie.verba && !strcmp(pici->lpVerb, mie.verba))
            return mie.cmd(this, pici);
    }

    return E_INVALIDARG;
}

HRESULT shell_item_list::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved,
                                          CHAR* pszName, UINT cchMax) {
    debug("shell_item_list::GetCommandString({}, {}, {}, {}, {})\n", idCmd, uType,
          (void*)pReserved, (void*)pszName, cchMax);

    span mi = menu_items;

    if (!IS_INTRESOURCE(idCmd)) {
        bool found = false;

        for (unsigned int i = 0; i < mi.size(); i++) {
            if (mi[i].verba == (char*)idCmd) {
                idCmd = i;
                found = true;
                break;
            }
        }

        if (!found)
            return E_INVALIDARG;
    } else {
        if (idCmd >= mi.size())
            return E_INVALIDARG;
    }

    switch (uType) {
        case GCS_VALIDATEA:
        case GCS_VALIDATEW:
            return S_OK;

        case GCS_VERBA:
            if (!mi[idCmd].verba)
                return E_INVALIDARG;

            return StringCchCopyA(pszName, cchMax, mi[idCmd].verba);

        case GCS_VERBW:
            if (!mi[idCmd].verbw)
                return E_INVALIDARG;

            return StringCchCopyW((STRSAFE_LPWSTR)pszName, cchMax, (WCHAR*)mi[idCmd].verbw);
    }

    return E_INVALIDARG;
}

void shell_item_list::populate_full_itemlist2(tar_item* item, const u16string& prefix) {
    u16string name = (prefix.empty() ? u"" : (prefix + u"\\"s)) + utf8_to_utf16(item->name);

    full_itemlist.emplace_back(item, name);

    for (auto& c : item->children) {
        populate_full_itemlist2(&c, name + u"\\");
    }
}

void shell_item_list::populate_full_itemlist() {
    if (!recursive) {
        for (auto item : itemlist) {
            full_itemlist.emplace_back(item, utf8_to_utf16(item->name));
        }

        return;
    }

    for (auto item : itemlist) {
        populate_full_itemlist2(item, u"");
    }
}

HRESULT shell_item_list::GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium) {
    char16_t format[256];

    if (!pformatetcIn || !pmedium)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetcIn->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item_list::GetData(pformatetcIn = [cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})], pmedium = [tymed = {}, hGlobal = {}])\n",
          pformatetcIn->cfFormat, utf16_to_utf8(format), (void*)pformatetcIn->ptd, pformatetcIn->dwAspect,
          pformatetcIn->lindex, pformatetcIn->tymed, pmedium->tymed, pmedium->hGlobal);

    if (pformatetcIn->cfFormat == cf_shell_id_list && pformatetcIn->tymed & TYMED_HGLOBAL) {
        if (full_itemlist.empty())
            populate_full_itemlist();

        pmedium->tymed = TYMED_HGLOBAL;
        pmedium->hGlobal = make_shell_id_list();
        pmedium->pUnkForRelease = nullptr;

        return S_OK;
    } else if (pformatetcIn->cfFormat == cf_file_descriptor && pformatetcIn->tymed & TYMED_HGLOBAL) {
        if (full_itemlist.empty())
            populate_full_itemlist();

        pmedium->tymed = TYMED_HGLOBAL;
        pmedium->hGlobal = make_file_descriptor();
        pmedium->pUnkForRelease = nullptr;

        return S_OK;
    } else if (pformatetcIn->cfFormat == cf_file_contents && pformatetcIn->tymed & TYMED_ISTREAM) {
        if (full_itemlist.empty())
            populate_full_itemlist();

        if (pformatetcIn->lindex >= full_itemlist.size())
            return E_INVALIDARG;

        try {
            auto tis = new tar_item_stream(tar, *full_itemlist[pformatetcIn->lindex].item);
            HRESULT hr;

            pmedium->tymed = TYMED_ISTREAM;

            hr = tis->QueryInterface(IID_IStream, (void**)&pmedium->pstm);
            if (FAILED(hr)) {
                delete tis;
                return hr;
            }

            hr = tis->QueryInterface(IID_IUnknown, (void**)&pmedium->pUnkForRelease);
            if (FAILED(hr)) {
                delete tis;
                return hr;
            }
        } catch (...) {
            return E_FAIL;
        }

        return S_OK;
    }

    return E_INVALIDARG;
}

HGLOBAL shell_item_list::make_shell_id_list() {
    HGLOBAL hg;
    CIDA* cida;
    size_t size, root_pidl_size;
    uint8_t* ptr;
    UINT* off;

    root_pidl_size = ILGetSize(root_pidl);

    size = offsetof(CIDA, aoffset) + (sizeof(UINT) * (full_itemlist.size() + 1)) + root_pidl_size;

    for (const auto& item : full_itemlist) {
        auto child_pidl = item.item->make_relative_pidl(root);

        size += ILGetSize(child_pidl);

        ILFree(child_pidl);
    }

    hg = GlobalAlloc(GHND | GMEM_SHARE, size);

    if (!hg)
        return nullptr;

    cida = (CIDA*)GlobalLock(hg);
    cida->cidl = full_itemlist.size();

    off = &cida->aoffset[0];
    ptr = (uint8_t*)cida + offsetof(CIDA, aoffset) + ((cida->cidl + 1) * sizeof(UINT));

    *off = ptr - (uint8_t*)cida;
    memcpy(ptr, root_pidl, root_pidl_size);
    ptr += root_pidl_size;
    off++;

    for (const auto& item : full_itemlist) {
        auto child_pidl = item.item->make_relative_pidl(root);
        size_t child_pidl_size = ILGetSize(child_pidl);

        *off = ptr - (uint8_t*)cida;
        memcpy(ptr, child_pidl, child_pidl_size);

        ILFree(child_pidl);

        ptr += child_pidl_size;
        off++;
    }

    GlobalUnlock(hg);

    return hg;
}

HGLOBAL shell_item_list::make_file_descriptor() {
    HGLOBAL hg;
    FILEGROUPDESCRIPTORW* fgd;
    FILEDESCRIPTORW* fd;

    hg = GlobalAlloc(GHND | GMEM_SHARE, offsetof(FILEGROUPDESCRIPTORW, fgd) + (full_itemlist.size() * sizeof(FILEDESCRIPTORW)));

    if (!hg)
        return nullptr;

    fgd = (FILEGROUPDESCRIPTORW*)GlobalLock(hg);
    fgd->cItems = full_itemlist.size();

    fd = &fgd->fgd[0];

    for (const auto& item : full_itemlist) {
        if (item.relative_path.length() >= MAX_PATH) {
            GlobalUnlock(hg);
            return nullptr;
        }

        fd->dwFlags = FD_ATTRIBUTES | FD_FILESIZE | FD_UNICODE; // FIXME

        if (item.item->dir)
            fd->dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        else
            fd->dwFileAttributes = 0;

        // FIXME - other attributes
        // FIXME - times

        fd->nFileSizeHigh = item.item->size >> 32;
        fd->nFileSizeLow = item.item->size & 0xffffffff;

        memcpy(fd->cFileName, item.relative_path.c_str(), (item.relative_path.length() + 1) * sizeof(char16_t));

        fd++;
    }

    GlobalUnlock(hg);

    return hg;
}

HRESULT shell_item_list::GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_list::QueryGetData(FORMATETC* pformatetc) {
    char16_t format[256];

    if (!pformatetc)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetc->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item_list::QueryGetData(cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})\n",
          pformatetc->cfFormat, utf16_to_utf8(format), (void*)pformatetc->ptd, pformatetc->dwAspect,
          pformatetc->lindex, pformatetc->tymed);

    if (pformatetc->cfFormat == cf_shell_id_list && pformatetc->tymed & TYMED_HGLOBAL)
        return S_OK;
    else if (pformatetc->cfFormat == cf_file_contents && pformatetc->tymed & TYMED_ISTREAM)
        return S_OK;
    else if (pformatetc->cfFormat == cf_file_descriptor && pformatetc->tymed & TYMED_HGLOBAL)
        return S_OK;

    return DV_E_TYMED;
}

HRESULT shell_item_list::GetCanonicalFormatEtc(FORMATETC* pformatectIn, FORMATETC* pformatetcOut) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_list::SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, WINBOOL fRelease) {
    char16_t format[256];

    if (!pformatetc || !pmedium)
        return E_INVALIDARG;

    GetClipboardFormatNameW(pformatetc->cfFormat, (WCHAR*)format, sizeof(format) / sizeof(char16_t));

    debug("shell_item_list::SetData(pformatetc = [cfFormat = {} ({}), ptd = {}, dwAspect = {}, lindex = {}, tymed = {})], pmedium = [tymed = {}, hGlobal = {}], fRelease = {})\n",
          pformatetc->cfFormat, utf16_to_utf8(format), (void*)pformatetc->ptd, pformatetc->dwAspect,
          pformatetc->lindex, pformatetc->tymed, pmedium->tymed, pmedium->hGlobal, fRelease);

    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_list::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc) {
    if (dwDirection == DATADIR_GET) {
        auto sief = new shell_item_enum_format(cf_shell_id_list, cf_file_contents, cf_file_descriptor);

        return sief->QueryInterface(IID_IEnumFORMATETC, (void**)ppenumFormatEtc);
    }

    return E_NOTIMPL;
}

HRESULT shell_item_list::DAdvise(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink, DWORD* pdwConnection) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_list::DUnadvise(DWORD dwConnection) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_list::EnumDAdvise(IEnumSTATDATA* *ppenumAdvise) {
    UNIMPLEMENTED; // FIXME
}

shell_item_enum_format::shell_item_enum_format(CLIPFORMAT cf_shell_id_list, CLIPFORMAT cf_file_contents,
                                               CLIPFORMAT cf_file_descriptor) {
    formats.emplace_back(cf_shell_id_list, TYMED_HGLOBAL);
    formats.emplace_back(cf_file_contents, TYMED_ISTREAM);
    formats.emplace_back(cf_file_descriptor, TYMED_HGLOBAL);
}

HRESULT shell_item_enum_format::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IEnumFORMATETC)
        *ppv = static_cast<IEnumFORMATETC*>(this);
    else {
        debug("shell_item_enum_format::QueryInterface: unsupported interface {}\n", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_item_enum_format::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_item_enum_format::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_item_enum_format::Next(ULONG celt, FORMATETC* rgelt, ULONG* pceltFetched) {
    if (pceltFetched)
        *pceltFetched = 0;

    while (celt > 0 && index < formats.size()) {
        rgelt->cfFormat = formats[index].format;
        rgelt->dwAspect = DVASPECT_CONTENT;
        rgelt->ptd = nullptr;
        rgelt->tymed = formats[index].tymed;
        rgelt->lindex = -1;

        rgelt++;
        celt--;
        index++;

        if (pceltFetched)
            (*pceltFetched)++;
    }

    return celt == 0 ? S_OK : S_FALSE;
}

HRESULT shell_item_enum_format::Skip(ULONG celt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_item_enum_format::Reset() {
    index = 0;

    return S_OK;
}

HRESULT shell_item_enum_format::Clone(IEnumFORMATETC** ppenum) {
    UNIMPLEMENTED; // FIXME
}
