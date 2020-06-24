#include "tarfldr.h"

using namespace std;

HRESULT shell_enum::QueryInterface(REFIID iid, void** ppv) {
    if (iid == IID_IUnknown || iid == IID_IEnumIDList)
        *ppv = static_cast<IEnumIDList*>(this);
    else {
        debug("shell_enum::QueryInterface: unsupported interface {}", iid);

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

        // FIXME - only show folders or non-folders as requested

        while (celt > 0 && index < root->children.size()) {
            *rgelt = root->children[index].make_pidl_child();

            rgelt++;
            celt--;
            index++;

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
    index = 0;

    return S_OK;
}

HRESULT shell_enum::Clone(IEnumIDList** ppenum) {
    UNIMPLEMENTED; // FIXME
}
