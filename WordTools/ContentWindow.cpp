#include "stdafx.h"
#include "ContentWindow.h"


CContentWindow::CContentWindow()
{
}


CContentWindow::~CContentWindow()
{
}

LRESULT CContentWindow::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	
	PAINTSTRUCT ps;
	HDC hDC = BeginPaint(&ps);

	::Rectangle(hDC, 0, 0, 50, 50);

	EndPaint(&ps);
	return 0;
}