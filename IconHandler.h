#pragma once
#include <windows.h>
#include <shlobj.h>

class CIconHandler : public IExtractIconW, public IPersistFile
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IExtractIconW
    IFACEMETHODIMP GetIconLocation(UINT uFlags, LPWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags);
	IFACEMETHODIMP Extract(LPCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize);

    // IPersistFile
    IFACEMETHODIMP GetClassID(CLSID *pClassID);
    IFACEMETHODIMP IsDirty();
    IFACEMETHODIMP Load(LPCOLESTR pszFileName, DWORD dwMode);
    IFACEMETHODIMP Save(LPCOLESTR pszFileName, BOOL fRemember);
    IFACEMETHODIMP SaveCompleted(LPCOLESTR pszFileName);
    IFACEMETHODIMP GetCurFile(LPOLESTR *ppszFileName);

    CIconHandler();

protected:
    ~CIconHandler();

private:
    long m_cRef;
    WCHAR m_szFileName[MAX_PATH];
};
