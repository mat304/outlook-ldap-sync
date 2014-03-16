#ifndef __CSYNCDIALOG_H_
#define __CSYNCDIALOG_H_

class CSyncDialog : 
	public CDialogImpl<CSyncDialog>
{
	const static UINT IDT_ANNIMATION_TIMER = 0;
	const static UINT IDT_BACKGROUNDSYNC_TIMER = 1;
	const static UINT ANNIMATION_TIMEOUT = 150; 
	const static bool BACKGROUNDSYNC_ENABLED = false; // Whether to do background sync
	const static UINT BACKGROUNDINITIAL_SYNC_TIMEOUT = 6000; // First background sync after 6sec
	const static UINT BACKGROUNDSYNC_TIMEOUT = 900000; // Subsequent background sync every 15min
	const static int MAX_FRAMES = 2;

	CStore* m_Store;
	int m_Frame;
	int m_SyncState;
	bool m_BackgroundSyncronizing;
	CTextBox m_TextBox;
	HBITMAP m_SyncImage[MAX_FRAMES];
//	HBRUSH m_BackgroundBrush;

public:

	CSyncDialog( CStore * Store );
	~CSyncDialog();

	void DoUserSync();
	void DoBackgroundSync();
	bool TerminateSyncThread();

	enum { IDD = IDD_CSYNCDIALOG };

BEGIN_MSG_MAP(CSyncDialog)
	MESSAGE_HANDLER(WM_PAINT, OnPaint)
	MESSAGE_HANDLER(WM_TIMER, OnTimer)
	COMMAND_ID_HANDLER(IDC_CLOSE, OnClose)
	MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
END_MSG_MAP()

private:

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnTimer(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnClose(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
};

#endif //__CSYNCDIALOG_H_
