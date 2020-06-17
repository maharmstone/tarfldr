#include <windows.h>

extern "C" STDAPI DllCanUnloadNow(void) {
    // FIXME
    return S_OK;
}

extern "C" STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    // FIXME

    return CLASS_E_CLASSNOTAVAILABLE;
}

extern "C" BOOL APIENTRY DllMain(HANDLE hModule, DWORD dwReason, void* lpReserved) {
    return true;
}
