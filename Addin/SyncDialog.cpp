// SyncDialog.cpp : Implementation of CCSyncDialog
#include "stdafx.h"
using namespace std;

/////////////////////////////////////////////////////////////////////////////
// CSyncDialog

CSyncDialog::CSyncDialog( CStore* Store )
	: m_Store( Store ), m_Frame(0), m_SyncState(0), m_BackgroundSyncronizing(false)
{
	// Load the 2 frames of the rotating syncronization animation
	m_SyncImage[0] =(HBITMAP)LoadImage(_Module.GetResourceInstance(),
		MAKEINTRESOURCE(IDB_SYNC1),IMAGE_BITMAP,0,0,0);
	m_SyncImage[1] =(HBITMAP)LoadImage(_Module.GetResourceInstance(),
		MAKEINTRESOURCE(IDB_SYNC2),IMAGE_BITMAP,0,0,0);
}

// Called by a create()
LRESULT CSyncDialog::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
//	m_BackgroundBrush = (HBRUSH)GetStockObject(WHITE_BRUSH);
	SetTimer( IDT_BACKGROUNDSYNC_TIMER, BACKGROUNDINITIAL_SYNC_TIMEOUT, (TIMERPROC) NULL);
	return 0;
}

bool CSyncDialog::TerminateSyncThread()
{
	if( m_Store->IsThreadActive() ) 
	{
		CStore::TerminateAllSyncThreads.Exchange(1);
		if( m_Store->WaitForCompletion(50) ) {
			TRACE("Terminated thread\n");
			return true;
		} else {
			TRACE("FAILED to Terminate thread\n");
			return false;
		}
	}
	return true;
}

// Terminate any background sync & do a user one
// Also display the sync window so they can see whats going on
// Exit when user presses the CLOSE button
void CSyncDialog::DoUserSync()
{	
	// Only 1 syncronize window allowed at a time
	if( m_SyncState > 0 ) return;
	m_SyncState = 1;

	m_TextBox.Clear();
	m_TextBox.Create( m_hWnd );  
	m_TextBox.MoveWindow( 20, 80, 380, 170 );

	SetTimer( IDT_ANNIMATION_TIMER, ANNIMATION_TIMEOUT, (TIMERPROC) NULL);
	ShowWindow(SW_SHOWNORMAL);
}

// Do a background sync unless a user sync is in operation
void CSyncDialog::DoBackgroundSync()
{	
	if( (m_SyncState==0) && ! m_Store->IsThreadActive() )
		m_Store->StartThread();
	m_Store->m_Addin->UpgradeCheck(); // Also check for upgrades
}

LRESULT CSyncDialog::OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    switch (wParam) 
    { 
    case IDT_ANNIMATION_TIMER: 

		if( ++m_Frame == MAX_FRAMES ) m_Frame = 0;

		if( m_SyncState==1 ) {
			if( TerminateSyncThread() ) {
				m_Store->RegisterStatusDisplay( &m_TextBox );
				m_TextBox.Clear();
				m_TextBox.DisplayString(L"Started server synchronization\n");
				m_Store->StartThread();
				m_SyncState = 2;
			}
		} else if( (m_SyncState==2 ) && !m_Store->IsThreadActive() ) {
			KillTimer(IDT_ANNIMATION_TIMER);
			m_Store->UnRegisterStatusDisplay();
			m_SyncState = 3;
		}

		InvalidateRgn(NULL);
		UpdateWindow();
		break; 
    case IDT_BACKGROUNDSYNC_TIMER: 
		if(BACKGROUNDSYNC_ENABLED) {
			DoBackgroundSync();
			SetTimer(IDT_BACKGROUNDSYNC_TIMER, BACKGROUNDSYNC_TIMEOUT, (TIMERPROC) NULL);
		}
		break;
	} 
	return 0;
}

LRESULT CSyncDialog::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	PAINTSTRUCT ps; 
    BeginPaint( &ps); 

	HDC hdc = CreateCompatibleDC( ps.hdc );
	SelectObject(hdc, m_SyncImage[m_Frame]);

	BITMAP ImageSize;
	GetObject(m_SyncImage[m_Frame],sizeof(BITMAP), &ImageSize );
	//StretchBlt( ps.hdc, 50, 18, 80, 80, hdc, 0, 0, ImageSize.bmWidth, ImageSize.bmHeight, SRCCOPY );
	BitBlt( ps.hdc, 20, 14, 48, 48, hdc, 0, 0, SRCCOPY );

	DeleteDC(hdc);
	EndPaint(&ps);
	return 0;
}

// Called when user presses the CLOSE button
LRESULT CSyncDialog::OnClose(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	ShowWindow(SW_HIDE);

	m_TextBox.DestroyWindow();

	if( (m_SyncState>0) && (m_SyncState<3) ) {
	
		KillTimer(IDT_ANNIMATION_TIMER);
		
		m_Store->UnRegisterStatusDisplay();

		TerminateSyncThread();
	}

	m_SyncState = 0;

	return 0;
}

// Called when window destroyed
LRESULT CSyncDialog::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	KillTimer( IDT_BACKGROUNDSYNC_TIMER );
	TerminateSyncThread();
	return 0;
}

CSyncDialog::~CSyncDialog()
{
}

