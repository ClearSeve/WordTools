#pragma once
using namespace ATL;

class CContentWindow : public CWindowImpl<CContentWindow>
{
public:
	CContentWindow();
	~CContentWindow();

public:
	BEGIN_MSG_MAP(CContentWindow)
		MESSAGE_HANDLER(WM_PAINT,OnPaint)
	END_MSG_MAP()
public:
	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
};

