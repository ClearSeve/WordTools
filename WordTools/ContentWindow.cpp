#include "stdafx.h"
#include "ContentWindow.h"

#include <Uxtheme.h>
#pragma comment(lib,"UxTheme.lib")

#include <vssym32.h>

CContentWindow::CContentWindow()
{
	m_nNumber = 0;
}


CContentWindow::~CContentWindow()
{
}

LRESULT CContentWindow::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// 获取系统颜色和系统字体
	GetSystemColor();
	GetSystemFont();

	// 设置定时器
	SetTimer(10001, 100);
	return 0;
}

LRESULT CContentWindow::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	
	PAINTSTRUCT ps;
	HDC hDC = BeginPaint(&ps);

	RECT rcClient = { 0 };
	GetClientRect(&rcClient);

	SelectObject(hDC, m_hSystemFont);
	SetBkMode(hDC, TRANSPARENT);

	//SetTextColor(hDC, m_clrSystemColor);
	DrawText(hDC, m_strText.GetBuffer(), m_strText.GetLength(), &rcClient, DT_CENTER | DT_SINGLELINE | DT_VCENTER);

	EndPaint(&ps);
	return 0;
}

LRESULT CContentWindow::OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{

	m_strText.Format(_T("%d"), m_nNumber++);

	Invalidate();
	
	return 0;
}

VOID CContentWindow::GetSystemColor()
{
	HTHEME hTheme = ::OpenThemeData(NULL, VSCLASS_TASKBAND);
	ATLASSERT(hTheme);

	if (hTheme)
	{
		const HRESULT hr = ::GetThemeColor(hTheme,TDP_GROUPCOUNT, 0, TMT_TEXTCOLOR, &m_clrSystemColor);
		ATLASSERT(SUCCEEDED(hr));

		::CloseThemeData(hTheme);
	}
}

VOID CContentWindow::GetSystemFont()
{
	HTHEME hTheme = ::OpenThemeData(NULL, VSCLASS_REBAR);
	ATLASSERT(hTheme);

	if (hTheme)
	{
		LOGFONT lf = { 0 };
		const HRESULT hr = ::GetThemeFont(hTheme,
			NULL, RP_BAND, 0, TMT_FONT, &lf);
		ATLASSERT(SUCCEEDED(hr));

		if (SUCCEEDED(hr))
		{
			m_hSystemFont = ::CreateFontIndirect(&lf);
			ATLASSERT(m_hFont);
		}

		::CloseThemeData(hTheme);
	}

}
