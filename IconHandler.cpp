#include "pch.h"
#include "IconHandler.h"
#include "resource.h"
#include <shlwapi.h>
#include <strsafe.h> // Add this for StringCchPrintfW
#include <fstream> // Add this for file logging

#pragma comment(lib, "shlwapi.lib")

extern HINSTANCE g_hInst;

// Helper function for logging to file
void LogToFile(const WCHAR* pszMessage)
{
    std::wofstream logFile(L"C:\\Users\\admin\\Desktop\\icon_handler_log.txt", std::ios_base::app);
    if (logFile.is_open())
    {
        logFile << pszMessage << std::endl;
        logFile.close();
    }
}

CIconHandler::CIconHandler() : m_cRef(1)
{
    m_szFileName[0] = L'\0';
}

CIconHandler::~CIconHandler()
{
}

// IUnknown
IFACEMETHODIMP CIconHandler::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(CIconHandler, IExtractIconW),
        QITABENT(CIconHandler, IPersistFile),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CIconHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) CIconHandler::Release()
{
    long cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

// IExtractIconW
IFACEMETHODIMP CIconHandler::GetIconLocation(UINT uFlags, LPWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags)
{
    WCHAR szLogMessage[MAX_PATH + 200]; // Increased buffer for log messages
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::GetIconLocation - uFlags: %u, cchMax: %u\n", uFlags, cchMax);
    LogToFile(szLogMessage);

    pszIconFile[0] = L'\0'; 
    *piIndex = -1;
    *pwFlags = GIL_PERINSTANCE | GIL_NOTFILENAME | GIL_DONTCACHE;

	//*pszIconFile = L'\0'; // (동적으로 제공할 경우 여전히 빈 경로 가능)
	//*piIndex = HashFromFileName(m_szFileName); // 간단한 해시 등
	//*pwFlags = GIL_PERINSTANCE | GIL_NOTFILENAME;

    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::GetIconLocation - pszIconFile: %s, piIndex: %d, pwFlags: %u\n", pszIconFile, *piIndex, *pwFlags);
    LogToFile(szLogMessage);

    return S_OK;
}

IFACEMETHODIMP CIconHandler::Extract(LPCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize)
{
    UINT uIconId = IDI_ICON1; // Default icon
    WCHAR szLogMessage[MAX_PATH + 100]; // Buffer for log message

    LPCWSTR pszFileNameToUse = m_szFileName;

    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - pszFile (received): %s \n", pszFile);
    LogToFile(szLogMessage);

    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - pszFileNameToUse (from m_szFileName): %s \n", pszFileNameToUse);
    LogToFile(szLogMessage);

    LPCWSTR pszFileNameOnly = PathFindFileNameW(pszFileNameToUse);

    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - pszFileNameOnly: %s \n", pszFileNameOnly);
    LogToFile(szLogMessage);

    if (PathMatchSpecW(pszFileNameOnly, L"*.xls.myext"))
    {
        uIconId = IDI_ICON_XLS;
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - Matched .xls.myext, uIconId: %d \n", uIconId);
        LogToFile(szLogMessage);
    }
    else if (PathMatchSpecW(pszFileNameOnly, L"*.pdf.myext"))
    {
        uIconId = IDI_ICON_PDF;
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - Matched .pdf.myext, uIconId: %d \n", uIconId);
        LogToFile(szLogMessage);
    }
    else
    {
        StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Extract - No match, uIconId: %d \n", uIconId);
        LogToFile(szLogMessage);
    }

    if (phiconLarge)
    {
        *phiconLarge = (HICON)LoadImageW(g_hInst, MAKEINTRESOURCEW(uIconId), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
    }
    if (phiconSmall)
    {
        *phiconSmall = (HICON)LoadImageW(g_hInst, MAKEINTRESOURCEW(uIconId), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
    }

    return S_OK;
}

// IPersistFile
IFACEMETHODIMP CIconHandler::GetClassID(CLSID *pClassID)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CIconHandler::IsDirty()
{
    return S_FALSE;
}

IFACEMETHODIMP CIconHandler::Load(LPCOLESTR pszFileName, DWORD dwMode)
{
    WCHAR szLogMessage[MAX_PATH + 100]; // Buffer for log message
    StringCchPrintfW(szLogMessage, ARRAYSIZE(szLogMessage), L"IconHandler::Load - pszFileName: %s \n", pszFileName);
    LogToFile(szLogMessage);

    StringCchCopyW(m_szFileName, ARRAYSIZE(m_szFileName), pszFileName);
    return S_OK;
}

IFACEMETHODIMP CIconHandler::Save(LPCOLESTR pszFileName, BOOL fRemember)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CIconHandler::SaveCompleted(LPCOLESTR pszFileName)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CIconHandler::GetCurFile(LPOLESTR *ppszFileName)
{
    return E_NOTIMPL;
}