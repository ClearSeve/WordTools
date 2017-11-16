#pragma once
#include <atlstr.h>
using namespace ATL;

class CContentWindow : public CWindowImpl<CContentWindow>
{
public:
	CContentWindow();
	~CContentWindow();

public:
	BEGIN_MSG_MAP(CContentWindow)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_PAINT,OnPaint)
		MESSAGE_HANDLER(WM_TIMER,OnTimer)
		
	END_MSG_MAP()
public:
	LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

public:
	

private:
	VOID GetSystemColor();
	VOID GetSystemFont();

private:
	DWORD m_nNumber;
	HFONT m_hSystemFont;
	COLORREF m_clrSystemColor;
	CString m_strText;

};

