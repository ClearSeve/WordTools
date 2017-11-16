// WordTools.cpp : DLL 导出的实现。


#include "stdafx.h"
#include "resource.h"
#include "WordTools_i.h"
#include "dllmain.h"


using namespace ATL;

// 用于确定 DLL 是否可由 OLE 卸载。
STDAPI DllCanUnloadNow(void)
{
			return _AtlModule.DllCanUnloadNow();
	}

// 返回一个类工厂以创建所请求类型的对象。
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID* ppv)
{
		return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - 在系统注册表中添加项。
STDAPI DllRegisterServer(void)
{
	// 注册对象、类型库和类型库中的所有接口
	HRESULT hr = _AtlModule.DllRegisterServer();
		return hr;
}

// DllUnregisterServer - 在系统注册表中移除项。
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
		return hr;
}

// DllInstall - 按用户和计算机在系统注册表中逐一添加/移除项。
STDAPI DllInstall(BOOL bInstall, _In_opt_  LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			ATL::AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{	
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}


// WordTools.cpp : CWordTools 的实现

#include "stdafx.h"
#include "WordTools.h"


// CWordTools


// CWordTools

////////////////////////////////////////////////////////////////////////////////
// IObjectWithSite

HRESULT STDMETHODCALLTYPE CWordTools::SetSite(IUnknown *pUnkSite)
{
	HRESULT hr = __super::SetSite(pUnkSite);

	if (SUCCEEDED(hr) && pUnkSite) // pUnkSite is NULL when band is being destroyed
	{
		CComQIPtr<IOleWindow> spOleWindow = pUnkSite;

		if (spOleWindow)
		{
			HWND hwndParent = NULL;
			hr = spOleWindow->GetWindow(&hwndParent);

			if (SUCCEEDED(hr))
			{
				m_wndContentWindow.Create(hwndParent /*,static_cast<IDeskBand*>(this), pUnkSite, &m_dateFormat*/);
				if (!m_wndContentWindow.IsWindow())
					hr = E_FAIL;
			}
		}
	}

	return hr;
}



////////////////////////////////////////////////////////////////////////////////
// IDeskBand

HRESULT STDMETHODCALLTYPE CWordTools::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO *pdbi)
{
	ATLTRACE(atlTraceCOM, 2, _T("IDeskBand::GetBandInfo\n"));

	HRESULT hr = E_INVALIDARG;

	if (pdbi)
	{
		m_dwBandID = dwBandID;

		if (pdbi->dwMask & DBIM_MINSIZE)
		{
			pdbi->ptMinSize.x = 200;
			pdbi->ptMinSize.y = 30;
		}

		if (pdbi->dwMask & DBIM_MAXSIZE)
		{
			pdbi->ptMaxSize.y = -1;
		}

		if (pdbi->dwMask & DBIM_INTEGRAL)
		{
			pdbi->ptIntegral.y = 1;
		}

		if (pdbi->dwMask & DBIM_ACTUAL)
		{
			pdbi->ptActual.x = 200;
			pdbi->ptActual.y = 30;
		}

		if (pdbi->dwMask & DBIM_TITLE)
		{
			// Don't show title by removing this flag.
			pdbi->dwMask &= ~DBIM_TITLE;
		}

		if (pdbi->dwMask & DBIM_MODEFLAGS)
		{
			pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_VARIABLEHEIGHT;
		}

		if (pdbi->dwMask & DBIM_BKCOLOR)
		{
			// Use the default background color by removing this flag.
			pdbi->dwMask &= ~DBIM_BKCOLOR;
		}

		hr = S_OK;
	}
	return hr;
}

////////////////////////////////////////////////////////////////////////////////
// IDeskBand2
HRESULT STDMETHODCALLTYPE CWordTools::CanRenderComposited(BOOL *pfCanRenderComposited)
{
	if (pfCanRenderComposited)
		*pfCanRenderComposited = TRUE;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CWordTools::SetCompositionState(BOOL fCompositionEnabled)
{
	m_bCompositionEnabled = fCompositionEnabled;
	m_wndContentWindow.Invalidate();

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CWordTools::GetCompositionState(BOOL *pfCompositionEnabled)
{
	if (pfCompositionEnabled)
		*pfCompositionEnabled = m_bCompositionEnabled;

	return S_OK;
}


////////////////////////////////////////////////////////////////////////////////
// IDockingWindow
HRESULT STDMETHODCALLTYPE CWordTools::ShowDW(BOOL fShow)
{
	ATLTRACE(atlTraceCOM, 2, _T("IDockingWindow::ShowDW\n"));

	if (m_wndContentWindow)
		m_wndContentWindow.ShowWindow(fShow ? SW_SHOW : SW_HIDE);

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CWordTools::CloseDW(DWORD /*dwReserved*/)
{
	ATLTRACE(atlTraceCOM, 2, _T("IDockingWindow::CloseDW\n"));

	if (m_wndContentWindow)
	{
		m_wndContentWindow.ShowWindow(SW_HIDE);
		m_wndContentWindow.DestroyWindow();
	}

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CWordTools::ResizeBorderDW(LPCRECT prcBorder, IUnknown *punkToolbarSite, BOOL /*fReserved*/)
{
	ATLTRACE(atlTraceCOM, 2, _T("IDockingWindow::ResizeBorderDW\n"));

	if (!m_wndContentWindow) return S_OK;

	CComQIPtr<IDockingWindowSite> spDockingWindowSite = punkToolbarSite;

	if (spDockingWindowSite)
	{
		BORDERWIDTHS bw = { 0 };
		bw.top = bw.bottom = ::GetSystemMetrics(SM_CYBORDER);
		bw.left = bw.right = ::GetSystemMetrics(SM_CXBORDER);

		HRESULT hr = spDockingWindowSite->RequestBorderSpaceDW(
			static_cast<IDeskBand*>(this), &bw);

		if (SUCCEEDED(hr))
		{
			HRESULT hr = spDockingWindowSite->SetBorderSpaceDW(
				static_cast<IDeskBand*>(this), &bw);

			if (SUCCEEDED(hr))
			{
				m_wndContentWindow.MoveWindow(prcBorder);
				return S_OK;
			}
		}
	}

	return E_FAIL;
}

////////////////////////////////////////////////////////////////////////////////
// IOleWindow
HRESULT STDMETHODCALLTYPE CWordTools::GetWindow(
	/* [out] */ HWND *phwnd)
{
	ATLTRACE(atlTraceCOM, 2, _T("IOleWindow::GetWindow\n"));

	if (phwnd) *phwnd = m_wndContentWindow;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CWordTools::ContextSensitiveHelp(
	/* [in] */ BOOL /*fEnterMode*/)
{

	ATLTRACE(atlTraceCOM, 2, _T("IOleWindow::ContextSensitiveHelp\n"));

	ATLTRACENOTIMPL("IOleWindow::ContextSensitiveHelp");
}
