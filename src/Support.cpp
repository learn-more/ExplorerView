/*
 * PROJECT:     ExplorerView
 * LICENSE:     MIT (https://spdx.org/licenses/MIT)
 * PURPOSE:     Debugging support functions
 * COPYRIGHT:   Copyright 2021 Mark Jansen <mark.jansen@reactos.org>
 */


#include "pch.h"

#define MAX_DBG_BUF_SIZE 4096

void DbgReport(_In_opt_z_ char const* format, ...)
{
    va_list arglist;
    va_start(arglist, format);

    char szMessage[MAX_DBG_BUF_SIZE]{ 0 };

    int szlen = vsprintf_s(szMessage, format, arglist);
    if (szlen >= 0)
    {
        OutputDebugStringA(szMessage);
    }

    va_end(arglist);
}

ATL::CStringW DumpIID(_In_ REFIID iid) throw()
{
    ATL::CStringW result;
    DWORD dwType;
    DWORD dw = 100 * sizeof(OLECHAR);

    ATL::CComHeapPtr<OLECHAR> pszGUID;
    if (FAILED(StringFromCLSID(iid, &pszGUID)))
        return result;

    LPWSTR szName = result.GetBuffer(100);

    ATL::CRegKey key;
    if (key.Open(HKEY_CLASSES_ROOT, _T("Interface"), KEY_READ) == ERROR_SUCCESS)
    {
        if (key.Open(key, pszGUID, KEY_READ) == ERROR_SUCCESS)
        {
            *szName = 0;
            if (RegQueryValueExW(key.m_hKey, (LPTSTR)NULL, NULL, &dwType, (LPBYTE)szName, &dw) == ERROR_SUCCESS)
            {
                result.ReleaseBuffer();
                return result;
            }
        }
    }
    // Attempt to find it in the clsid section
    if (key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), KEY_READ) == ERROR_SUCCESS)
    {
        if (key.Open(key, pszGUID, KEY_READ) == ERROR_SUCCESS)
        {
            dw = 100 * sizeof(OLECHAR);
            if (RegQueryValueEx(key.m_hKey, (LPTSTR)NULL, NULL, &dwType, (LPBYTE)szName, &dw) == ERROR_SUCCESS)
            {
                result.ReleaseBuffer();
                return result;
            }
        }
    }

    result.ReleaseBuffer();
    result = pszGUID;

    return result;
}


HRESULT _ShellFolderFromPidl(PCUIDLIST_RELATIVE pidl, ATL::CComPtr<IShellFolder>& folder)
{
    ATL::CComPtr<IShellFolder> spDesktop;
    HRESULT hr = SHGetDesktopFolder(&spDesktop);
    if (FAILED(hr))
        return hr;

    // Not the desktop?
    if (pidl->mkid.cb)
    {
        return spDesktop->BindToObject(pidl, NULL, IID_PPV_ARGS(&folder));
    }
    else
    {
        folder = spDesktop;
        return S_OK;
    }
}

