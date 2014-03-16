// TextBox.cpp : Implementation of CTextBox
#include "stdafx.h"
using namespace std;

/////////////////////////////////////////////////////////////////////////////
// CTextBox
const wstring CTextBox::m_LeftMargin = L"  ";

LRESULT CTextBox::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	PAINTSTRUCT ps; 
    HDC hdc = BeginPaint( &ps); 

	// Get vertical scroll bar position
	SCROLLINFO si;
	si.cbSize = sizeof (si);
	si.fMask  = SIF_POS;
	GetScrollInfo( SB_VERT, &si);
	int yPos = si.nPos;

    SelectObject(hdc, m_Font); 
	SetBkMode(hdc, TRANSPARENT);

	int y = 0;
	if ( m_Lines.size() > 0 )
	{
		// Find painting limits
		int FirstLine = max (0, yPos + ps.rcPaint.top / m_nLineHeight);
		int LastLine = min ( (int)m_Lines.size() - 1, yPos + ps.rcPaint.bottom / m_nLineHeight);

		for (int i = FirstLine; i <= LastLine; i++)
		{
			wstring Line(m_LeftMargin);
			Line += m_Lines[i];

			y = m_nLineHeight * (i - yPos);

			// Calculate the rectangle that the line lies within
			// DrawText should extend the rectangle according to text size
			RECT LineRect = { 0 /*left*/, y /*top*/, ps.rcPaint.right /*right*/, y + m_nLineHeight /*bottom*/ };
			FillRect(hdc, &LineRect, m_BackgroundBrush);
			DrawText(hdc, Line.c_str(), Line.size(), &LineRect, DT_TOP); 
		}
	}
	y += m_nLineHeight;

	// Fill the rest of the region if theres no text there
	if( y < ps.rcPaint.bottom )
	{
		RECT Rect = { 0 /*left*/, y /*top*/, ps.rcPaint.right /*right*/, ps.rcPaint.bottom /*bottom*/ };
		FillRect(hdc, &Rect, m_BackgroundBrush);
	}
	
	// Indicate that painting is finished
	EndPaint(&ps);
	return 0;
}

LRESULT CTextBox::OnScroll(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	SCROLLINFO si;

	// Get all the vertical scroll bar information
	si.cbSize = sizeof (si);
	si.fMask  = SIF_ALL;
	GetScrollInfo ( SB_VERT, &si);

	// Save the position for comparison later on
	int yPos = si.nPos;
	switch (LOWORD (wParam))
	{
		// user clicked the HOME keyboard key
		case SB_TOP:
			si.nPos = si.nMin;
			break;
              
		// user clicked the END keyboard key
		case SB_BOTTOM:
			si.nPos = si.nMax;
			break;
              
		// user clicked the top arrow
		case SB_LINEUP:
			si.nPos -= 1;
			break;
              
		// user clicked the bottom arrow
		case SB_LINEDOWN:
			si.nPos += 1;
			break;
              
		// user clicked the shaft above the scroll box
		case SB_PAGEUP:
			si.nPos -= si.nPage;
			break;
              
		// user clicked the shaft below the scroll box
		case SB_PAGEDOWN:
			si.nPos += si.nPage;
			break;
              
		// user dragged the scroll box
		case SB_THUMBTRACK:
			si.nPos = si.nTrackPos;
			break;
              
		default:
			break; 
		}

		// Set the position and then retrieve it.  Due to adjustments
		//   by Windows it may not be the same as the value set.
		si.fMask = SIF_POS;
		SetScrollInfo ( SB_VERT, &si, TRUE);
		GetScrollInfo ( SB_VERT, &si);
		// If the position has changed, scroll window and update it
		if (si.nPos != yPos)
		{
			RECT rc;
			GetClientRect(&rc);
			ScrollWindowEx( 0, (m_nLineHeight * (yPos - si.nPos)), NULL, NULL, NULL, &rc, SW_INVALIDATE );
			UpdateWindow ();
		}
	return 0;
}

LRESULT CTextBox::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	HDC hdc; 
	TEXTMETRIC tm; 

	// Select the font we're going to use for this window
	m_Font = (HFONT)GetStockObject(ANSI_VAR_FONT); 

	hdc = GetDC();

	// Extract font dimensions from the text metrics. 
    SelectObject(hdc, m_Font); 
	GetTextMetrics (hdc, &tm); 
	m_nLineHeight = tm.tmHeight + tm.tmExternalLeading; 
 
	// Background brush
	m_BackgroundBrush = (HBRUSH)GetStockObject(WHITE_BRUSH);

	// Free the device context. 
	ReleaseDC (hdc); 

	// Initialize the scroll bar
	SCROLLINFO si;
	si.cbSize = sizeof( si );
	si.fMask  = SIF_PAGE | SIF_POS | SIF_RANGE;
	si.nMin	  = 0;
	si.nMax	  = 0;
	si.nPage  = 0; 
	si.nPos   = 0; 
	SetScrollInfo( SB_VERT, &si, TRUE );

	return 0;
}

void CTextBox::Clear()
{
	m_Lines.clear();
}

void CTextBox::DisplayString( LPCOLESTR s )
{
	static wstring ThisLine;

	int nLines = m_Lines.size();
	LPCOLESTR cp, lp;
	for( cp = lp = s; *cp; cp++)
	{
		if( *cp == L'\n' ) 
		{
			ThisLine += wstring( lp, (cp==lp)?0:(cp-lp) );
			m_Lines.push_back( ThisLine );
			ThisLine.erase();
			lp = cp + 1;
		}
	}

	if( cp != lp ) 
		ThisLine += wstring( lp, cp-lp );

	//  Only update scroll bar and display text after a new line has been added
	if( nLines != m_Lines.size() )
	{
		SCROLLINFO si;
		si.cbSize = sizeof (si);
		si.fMask  = SIF_ALL;
		GetScrollInfo ( SB_VERT, &si);
		si.nMax	= m_Lines.size();
		SetScrollInfo( SB_VERT, &si, TRUE );

		PostMessage(WM_VSCROLL,SB_LINEDOWN);

		InvalidateRgn(NULL);
		UpdateWindow();
	}
}


LRESULT CTextBox::OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	SCROLLINFO si;

	// Retrieve the dimensions of the client area. 
	int yClient = HIWORD (lParam); 
	int xClient = LOWORD (lParam); 
 
	// Set the vertical scrolling range and page size
	si.cbSize = sizeof(si); 
	si.fMask  = SIF_RANGE | SIF_PAGE; 
	si.nMin   = 0; 
	si.nMax   = m_Lines.size() - 1; 
	si.nPage  = yClient / m_nLineHeight; 
	SetScrollInfo(SB_VERT, &si, TRUE); 
 
	return 0; 
}