/*
 * PROJECT:     ExplorerView
 * LICENSE:     MIT (https://spdx.org/licenses/MIT)
 * PURPOSE:     Main window implementation
 * COPYRIGHT:   Copyright 2021 Mark Jansen <mark.jansen@reactos.org>
 */

#include "pch.h"
#include "resource.h"

using namespace ATL;


struct Location
{
    CComHeapPtr<ITEMIDLIST> Pidl;
    CComPtr<IShellFolder> Folder;
};


class CExplorerHost: public CWindowImpl<CExplorerHost, CWindow, CFrameWinTraits>
    , IShellBrowser
    , IServiceProviderImpl<CExplorerHost>
{
public:
    HWND m_hView;
    CComPtr<IShellView> m_spView;

    Location m_Current;

    CExplorerHost()
        :m_hView(NULL)
    {
    }

    BEGIN_MSG_MAP(CExplorerHost)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
    END_MSG_MAP()

    BEGIN_SERVICE_MAP(CExplorerHost)
        SERVICE_ENTRY(SID_STopLevelBrowser)
        SERVICE_ENTRY(SID_SShellBrowser)
        // Show failed QueryService attempts
        DBG_TRACE(__FUNCTION__ "(guidService = %S, riid = %S) = E_NOINTERFACE\n", DumpIID(guidService).GetString(), DumpIID(riid).GetString());
    END_SERVICE_MAP()

    HRESULT STDMETHODCALLTYPE v_MayTranslateAccelerator(MSG* pmsg)
    {
        if (m_spView)
            return m_spView->TranslateAccelerator(pmsg);
        return S_FALSE;
    }

    // +IUnknown
    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject)
    {
        if (riid == IID_IUnknown)
        {
            //DBG_TRACE(__FUNCTION__ "(riid = IID_IUnknown) = S_OK\n");
            *ppvObject = (IUnknown*)(IShellBrowser*)this;
            return S_OK;
        }
        if (riid == IID_IServiceProvider)
        {
            //DBG_TRACE(__FUNCTION__ "(riid = IID_IServiceProvider) = S_OK\n");
            *ppvObject = (IServiceProvider*)this;
            return S_OK;
        }
        if (riid == IID_IOleWindow)
        {
            //DBG_TRACE(__FUNCTION__ "(riid = IID_IOleWindow) = S_OK\n");
            *ppvObject = (IOleWindow*)this;
            return S_OK;
        }
        if (riid == IID_IShellBrowser)
        {
            //DBG_TRACE(__FUNCTION__ "(riid = IID_IShellBrowser) = S_OK\n");
            *ppvObject = (IShellBrowser*)this;
            return S_OK;
        }
        DBG_TRACE(__FUNCTION__ "(riid = %S) = E_NOINTERFACE\n", DumpIID(riid).GetString());
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)()
    {
        return 2L;
    }

    STDMETHOD_(ULONG, Release)()
    {
        return 1L;
    }
    // -IUnknown


    // +IOleWindow
    STDMETHOD(GetWindow)(HWND* phwnd)
    {
        *phwnd = m_hWnd;
        return S_OK;
    }

    STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode)
    {
        DBG_TRACE(__FUNCTION__ "(fEnterMode = %d)\n", fEnterMode);
        return E_NOTIMPL;
    }
    // -IOleWindow

    // +IShellBrowser
    STDMETHOD(InsertMenusSB)(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
    {
        DBG_TRACE(__FUNCTION__ "(hmenuShared = %p, lpMenuWidths = %p)\n", hmenuShared, lpMenuWidths);
        return E_NOTIMPL;
    }

    STDMETHOD(SetMenuSB)(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject)
    {
        DBG_TRACE(__FUNCTION__ "(hmenuShared = %p, holemenuRes = %p, hwndActiveObject = %p)\n", hmenuShared, holemenuRes, hwndActiveObject);
        return E_NOTIMPL;
    }

    STDMETHOD(RemoveMenusSB)(HMENU hmenuShared)
    {
        DBG_TRACE(__FUNCTION__ "(hmenuShared = %p)\n", hmenuShared);
        return E_NOTIMPL;
    }

    STDMETHOD(SetStatusTextSB)(LPCWSTR pszStatusText)
    {
        DBG_TRACE(__FUNCTION__ "(pszStatusText = %s)\n", pszStatusText);
        return E_NOTIMPL;
    }

    STDMETHOD(EnableModelessSB)(BOOL fEnable)
    {
        DBG_TRACE(__FUNCTION__ "(fEnable = %d)\n", fEnable);
        return E_NOTIMPL;
    }

    STDMETHOD(TranslateAcceleratorSB)(MSG* pmsg, WORD wID)
    {
        DBG_TRACE(__FUNCTION__ "(pmsg = %p, wID = %u)\n", pmsg, wID);
        return E_NOTIMPL;
    }

    STDMETHOD(BrowseObject)(PCUIDLIST_RELATIVE pidl, UINT wFlags)
    {
        Location tmp;
        HRESULT hr;

        if (wFlags & SBSP_RELATIVE)
        {
            tmp.Pidl.Attach(ILCombine(m_Current.Pidl, pidl));
            hr = m_Current.Folder->BindToObject(pidl, NULL, IID_PPV_ARGS(&tmp.Folder));
            if (FAILED(hr))
                return hr;
        }
        else if (wFlags & SBSP_PARENT)
        {
            tmp.Pidl.Attach(::ILClone(pidl));
            ::ILRemoveLastID(tmp.Pidl);
        }
        else
        {
            tmp.Pidl.Attach(::ILClone(pidl));
        }


        if (ILIsEqual(tmp.Pidl, m_Current.Pidl))
        {
            // We are already here!
            return S_OK;
        }

        if (!tmp.Folder)
        {
            hr = _ShellFolderFromPidl(pidl, tmp.Folder);
            if (FAILED(hr))
                return hr;
        }

        m_Current = tmp;

        FOLDERSETTINGS fs = {};
        if (m_spView)
        {
            m_spView->GetCurrentInfo(&fs);
        }
        else
        {
            fs.ViewMode = FVM_AUTO;
            fs.fFlags = FWF_DESKTOP;
        }

        CComPtr<IShellView> last = m_spView;
        m_spView.Release();
        hr = m_Current.Folder->CreateViewObject(m_hWnd, IID_PPV_ARGS(&m_spView));

        RECT rcView;
        GetClientRect(&rcView);
        hr = m_spView->CreateViewWindow(last, &fs, this, &rcView, &m_hView);

        if (last)
        {
            last->UIActivate(SVUIA_DEACTIVATE);
            last->DestroyViewWindow();
            last.Release();
        }

        m_spView->UIActivate(SVUIA_ACTIVATE_FOCUS);

        return S_OK;
    }

    STDMETHOD(GetViewStateStream)(DWORD grfMode, IStream** ppStrm)
    {
        DBG_TRACE(__FUNCTION__ "(grfMode = %d, ppStrm = %p)\n", grfMode, ppStrm);
        return E_NOTIMPL;
    }

    STDMETHOD(GetControlWindow)(UINT id, HWND* phwnd)
    {
        *phwnd = 0;
        DBG_TRACE(__FUNCTION__ "(id = %d, phwnd = %p)\n", id, phwnd);
        return S_FALSE;
    }

    STDMETHOD(SendControlMsg)(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam,LRESULT* pret)
    {
        DBG_TRACE(__FUNCTION__ "(id = %d, uMsg = 0x%x, wParam = 0x%x, lParam = 0x%x)\n", id, uMsg, wParam, lParam);
        return E_NOTIMPL;
    }

    STDMETHOD(QueryActiveShellView)(IShellView** ppshv)
    {
        DBG_TRACE(__FUNCTION__ "(ppshv = %p)\n", ppshv);
        ATLASSERT(m_spView);
        *ppshv = m_spView;
        (*ppshv)->AddRef();
        return S_OK;
    }

    STDMETHOD(OnViewWindowActive)(IShellView* pshv)
    {
        DBG_TRACE(__FUNCTION__ "(pshv = %p)\n", pshv);
        return E_NOTIMPL;
    }

    STDMETHOD(SetToolbarItems)(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags)
    {
        DBG_TRACE(__FUNCTION__ "(lpButtons = %p, nButtons = %u, uFlags = 0x%x)\n", lpButtons, nButtons, uFlags);
        return E_NOTIMPL;
    }
    // -IShellBrowser


    LRESULT OnCreate(UINT /*nMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
    {
        CComHeapPtr<ITEMIDLIST> desktop;
        (void)SHGetSpecialFolderLocation(NULL, CSIDL_DESKTOP, &desktop);
        BrowseObject(desktop, SBSP_ABSOLUTE);

        return 0L;
    }

    LRESULT OnDestroy(UINT /*nMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
    {
        if (m_spView)
        {
            m_spView->UIActivate(SVUIA_DEACTIVATE);
            m_spView->DestroyViewWindow();
            m_spView.Release();
        }

        ::PostQuitMessage(0);
        return 0L;
    }
};

extern "C"
int WINAPI _tWinMain(_In_ HINSTANCE /*hInstance*/, _In_opt_ HINSTANCE /*hPrevInstance*/, _In_ LPTSTR /*lpCmdLine*/, _In_ int nShowCmd)
{
    _CrtSetDbgFlag(_CrtSetDbgFlag(_CRTDBG_REPORT_FLAG) | _CRTDBG_LEAK_CHECK_DF);

    HRESULT hr = OleInitialize(NULL);
    if (FAILED(hr))
        return hr;

    CExplorerHost wnd;

    wnd.Create(NULL);
    wnd.ShowWindow(nShowCmd);

    MSG msg;
    while (GetMessage(&msg, 0, 0, 0) > 0)
    {
        // Give the View a chance to translate accelerators
        if (wnd.v_MayTranslateAccelerator(&msg) != S_OK)
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    OleUninitialize();

    return 0;
}

