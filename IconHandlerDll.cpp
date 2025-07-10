#include "pch.h"
#include <initguid.h>
#include "IconHandler.h"
#include <shlwapi.h> // For StrStrIW and registry functions
#include <strsafe.h>
void LogToFile(const WCHAR* pszMessage); // For StringCchPrintfW
#include <objbase.h> // For GUID_STRING_LEN

#define GUID_STRING_LEN 39 // Manually define if not found in objbase.h

#pragma comment(lib, "shlwapi.lib") // Link with shlwapi.lib

// {7F4A8F8C-4C4E-4CEB-B055-35524A51E54B}
DEFINE_GUID(CLSID_IconHandler, 0x7f4a8f8c, 0x4c4e, 0x4ceb, 0xb0, 0x55, 0x35, 0x52, 0x4a, 0x51, 0xe5, 0x4b);

HINSTANCE g_hInst = NULL;
long g_cDllRef = 0;

// Helper function to set a registry key's value
HRESULT SetRegKeyValue(HKEY hKey, PCWSTR pszSubKey, PCWSTR pszValueName, PCWSTR pszData)
{
    HKEY hSubKey;
    HRESULT hr = HRESULT_FROM_WIN32(RegCreateKeyExW(hKey, pszSubKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hSubKey, NULL));
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(RegSetValueExW(hSubKey, pszValueName, 0, REG_SZ, (LPBYTE)pszData, (lstrlenW(pszData) + 1) * sizeof(WCHAR)));
        RegCloseKey(hSubKey);
    }
    return hr;
}

// Helper function to delete a registry key
HRESULT DeleteRegKey(HKEY hKey, PCWSTR pszSubKey)
{
    LSTATUS status = RegDeleteTreeW(hKey, pszSubKey);
    if (status == ERROR_FILE_NOT_FOUND)
    {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(status);
}

class CClassFactory : public IClassFactory
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv);
    IFACEMETHODIMP LockServer(BOOL fLock);

    CClassFactory() : m_cRef(1) {}
    ~CClassFactory() {}

private:
    long m_cRef;
};

// IUnknown
IFACEMETHODIMP CClassFactory::QueryInterface(REFIID riid, void **ppv)
{
    if (riid == IID_IUnknown || riid == IID_IClassFactory)
    {
        *ppv = static_cast<IClassFactory *>(this);
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    reinterpret_cast<IUnknown *>(*ppv)->AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) CClassFactory::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) CClassFactory::Release()
{
    long cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

// IClassFactory
IFACEMETHODIMP CClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
    *ppv = NULL;
    if (pUnkOuter)
    {
        return CLASS_E_NOAGGREGATION;
    }

    CIconHandler *pIconHandler = new CIconHandler();
    if (!pIconHandler)
    {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pIconHandler->QueryInterface(riid, ppv);
    pIconHandler->Release();
    return hr;
}

IFACEMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
    if (fLock)
    {
        InterlockedIncrement(&g_cDllRef);
    }
    else
    {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
    *ppv = NULL;
    if (rclsid == CLSID_IconHandler)
    {
        CClassFactory *pcf = new CClassFactory();
        if (!pcf)
        {
            return E_OUTOFMEMORY;
        }
        HRESULT hr = pcf->QueryInterface(riid, ppv);
        pcf->Release();
        return hr;
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow(void)
{
    return (g_cDllRef == 0) ? S_OK : S_FALSE;
}

STDAPI DllRegisterServer(void)
{
    HRESULT hr = S_OK;
    WCHAR szModulePath[MAX_PATH];
    WCHAR szCLSID[GUID_STRING_LEN];
    WCHAR szLogMessage[MAX_PATH + 200]; // Increased buffer for log messages

    // Get module path
    if (!GetModuleFileNameW(g_hInst, szModulePath, ARRAYSIZE(szModulePath)))
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: GetModuleFileNameW failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: Module Path: %s\n", szModulePath);
    LogToFile(szLogMessage);

    // Convert CLSID to string
    if (StringFromGUID2(CLSID_IconHandler, szCLSID, ARRAYSIZE(szCLSID)) == 0)
    {
        hr = E_FAIL;
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: StringFromGUID2 failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: CLSID: %s\n", szCLSID);
    LogToFile(szLogMessage);

    // Register CLSID
    WCHAR szCLSIDKey[MAX_PATH];
    hr = StringCchPrintfW(szCLSIDKey, ARRAYSIZE(szCLSIDKey), L"CLSID\\%s", szCLSID);
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: StringCchPrintfW for CLSIDKey failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: CLSID Key: %s\n", szCLSIDKey);
    LogToFile(szLogMessage);

    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, szCLSIDKey, NULL, L"MyExt Icon Handler");
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for CLSID failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: CLSID registered successfully.\n");
    LogToFile(szLogMessage);

    // Register InprocServer32
    WCHAR szInprocServer32Key[MAX_PATH];
    hr = StringCchPrintfW(szInprocServer32Key, ARRAYSIZE(szInprocServer32Key), L"%s\\InprocServer32", szCLSIDKey);
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: StringCchPrintfW for InprocServer32Key failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: InprocServer32 Key: %s\n", szInprocServer32Key);
    LogToFile(szLogMessage);

    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, szInprocServer32Key, NULL, szModulePath);
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for InprocServer32 path failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: InprocServer32 path registered successfully.\n");
    LogToFile(szLogMessage);

    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, szInprocServer32Key, L"ThreadingModel", L"Apartment");
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for ThreadingModel failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: ThreadingModel registered successfully.\n");
    LogToFile(szLogMessage);

    // Register file extension and association
    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, L".myext", NULL, L"MyExtFile");
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for .myext failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: .myext registered successfully.\n");
    LogToFile(szLogMessage);

    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, L"MyExtFile", NULL, L"My Custom File Type");
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for MyExtFile failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: MyExtFile registered successfully.\n");
    LogToFile(szLogMessage);

    WCHAR szShellExKey[MAX_PATH];
    hr = StringCchPrintfW(szShellExKey, ARRAYSIZE(szShellExKey), L"MyExtFile\\ShellEx\\IconHandler");
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: StringCchPrintfW for ShellExKey failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: ShellEx Key: %s\n", szShellExKey);
    LogToFile(szLogMessage);

    hr = SetRegKeyValue(HKEY_CLASSES_ROOT, szShellExKey, NULL, szCLSID);
    if (FAILED(hr))
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SetRegKeyValue for ShellEx failed with HRESULT: 0x%08X\n", hr);
        LogToFile(szLogMessage);
        return hr;
    }
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: ShellEx registered successfully.\n");
    LogToFile(szLogMessage);

    // Notify shell of changes
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"DllRegisterServer: SHChangeNotify called.\n");
    LogToFile(szLogMessage);

    return S_OK;
}

STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;
    WCHAR szCLSID[GUID_STRING_LEN];

    // Convert CLSID to string
    if (StringFromGUID2(CLSID_IconHandler, szCLSID, ARRAYSIZE(szCLSID)) == 0)
    {
        return E_FAIL;
    }

    // Unregister file extension and association.
    // We try to delete all keys and ignore errors, especially "not found" errors,
    // to ensure the unregistration process is as complete as possible.
    DeleteRegKey(HKEY_CLASSES_ROOT, L"MyExtFile");
    DeleteRegKey(HKEY_CLASSES_ROOT, L".myext");

    // Unregister CLSID
    WCHAR szCLSIDKey[MAX_PATH];
    if (SUCCEEDED(StringCchPrintfW(szCLSIDKey, ARRAYSIZE(szCLSIDKey), L"CLSID\%s", szCLSID)))
    {
        DeleteRegKey(HKEY_CLASSES_ROOT, szCLSIDKey);
    }

    // Notify shell of changes
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

    return S_OK;
}
