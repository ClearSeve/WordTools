#include "stdafx.h"
#include "ContentWindow.h"

#include <Uxtheme.h>
#pragma comment(lib,"UxTheme.lib")

#include <vssym32.h>
#include <atltypes.h>

CContentWindow::CContentWindow()
{
	m_nNumber = 0;
}


CContentWindow::~CContentWindow()
{
}


class DoubleBufferPaint
{
public:
	DoubleBufferPaint(HWND hWnd, DWORD dwRop = SRCCOPY) : m_hWnd(hWnd), m_dwRop(dwRop)
	{
		m_hdc = ::BeginPaint(m_hWnd, &m_ps);

		int x, y, cx, cy;
		InitDims(&x, &y, &cx, &cy);

		m_hdcMem = ::CreateCompatibleDC(m_hdc);
		m_hbmMem = ::CreateCompatibleBitmap(m_hdc, cx, cy);
		m_hbmSave = ::SelectObject(m_hdcMem, m_hbmMem);
		::SetWindowOrgEx(m_hdcMem, x, y, NULL);
	}

	~DoubleBufferPaint()
	{
		int x, y, cx, cy;
		InitDims(&x, &y, &cx, &cy);

		::BitBlt(m_hdc, x, y, cx, cy, m_hdcMem, x, y, m_dwRop);

		::SelectObject(m_hdcMem, m_hbmSave);
		::DeleteObject(m_hbmMem);
		::DeleteDC(m_hdcMem);

		::EndPaint(m_hWnd, &m_ps);
	}

	HDC GetDC() const { return m_hdcMem; }
	const RECT& GetPaintRect() const { return m_ps.rcPaint; }

private:
	void InitDims(int* px, int* py, int* pcx, int* pcy) const
	{
		*px = m_ps.rcPaint.left;
		*py = m_ps.rcPaint.top;
		*pcx = m_ps.rcPaint.right - m_ps.rcPaint.left;
		*pcy = m_ps.rcPaint.bottom - m_ps.rcPaint.top;
	}

private:
	HWND m_hWnd;
	DWORD m_dwRop;

	HDC m_hdc;
	HDC m_hdcMem;
	HGDIOBJ m_hbmMem;
	HGDIOBJ m_hbmSave;
	PAINTSTRUCT m_ps;

private:
	// not copyable
	DoubleBufferPaint(const DoubleBufferPaint&);
	DoubleBufferPaint& operator=(const DoubleBufferPaint&);
};

LRESULT CContentWindow::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// 获取系统颜色和系统字体
	GetSystemColor();
	GetSystemFont();

	// 设置定时器
	SetTimer(10001, 100);

	m_nWidth = 200;
	m_nHeight = 30;
	return 0;
}

LRESULT CContentWindow::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	DoubleBufferPaint Double(m_hWnd);

	//PAINTSTRUCT ps;
	//HDC hDC = BeginPaint(&ps);

	CRect rect = { 0 };
	GetClientRect(&rect);

	SelectObject(Double.GetDC(), m_hSystemFont);
	SetBkMode(Double.GetDC(), TRANSPARENT);
	
	SetTextColor(Double.GetDC(), m_clrSystemColor);

	CRect rectText(rect);
	int height = DrawText(Double.GetDC(),m_strText.GetBuffer(), m_strText.GetLength(),rectText, DT_CENTER | DT_WORDBREAK | DT_CALCRECT | DT_EDITCONTROL); // 获得文本高度  
	rect.OffsetRect(0, (rect.Height() - height) / 2);
	DrawText(Double.GetDC(), m_strText.GetBuffer(), m_strText.GetLength(), rect, DT_RIGHT | DT_WORDBREAK);


	return 0;
}

LRESULT CContentWindow::OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{

	m_strText.Format(_T("数据1:%d\r\n数据2:%d"), m_nNumber++);

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
