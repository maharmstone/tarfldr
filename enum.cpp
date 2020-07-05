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

using namespace std;

HRESULT shell_enum::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IEnumIDList)
        *ppv = static_cast<IEnumIDList*>(this);
    else {
        debug("shell_enum::QueryInterface: unsupported interface {}\n", iid);

        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    reinterpret_cast<IUnknown*>(*ppv)->AddRef();

    return S_OK;
}

ULONG shell_enum::AddRef() {
    return InterlockedIncrement(&refcount);
}

ULONG shell_enum::Release() {
    LONG rc = InterlockedDecrement(&refcount);

    if (rc == 0)
        delete this;

    return rc;
}

HRESULT shell_enum::Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) {
    try {
        if (pceltFetched)
            *pceltFetched = 0;

        // FIXME - SHCONTF_INCLUDEHIDDEN

        while (celt > 0 && it != root->children.end()) {
            if (!(flags & SHCONTF_FOLDERS) && it->dir) {
                it++;
                continue;
            }

            if (!(flags & SHCONTF_NONFOLDERS) && !it->dir) {
                it++;
                continue;
            }

            *rgelt = it->make_pidl_child();

            rgelt++;
            celt--;
            it++;

            if (pceltFetched)
                (*pceltFetched)++;
        }

        return celt == 0 ? S_OK : S_FALSE;
    } catch (const bad_alloc&) {
        return E_OUTOFMEMORY;
    }
}

HRESULT shell_enum::Skip(ULONG celt) {
    UNIMPLEMENTED; // FIXME
}

HRESULT shell_enum::Reset() {
    it = root->children.begin();

    return S_OK;
}

HRESULT shell_enum::Clone(IEnumIDList** ppenum) {
    UNIMPLEMENTED; // FIXME
}
