#ifndef __TEXTBOX_H_
#define __TEXTBOX_H_

class CTextBox : 
	public CAxDialogImpl<CTextBox>,
	public IDisplay
{

	vector<wstring> m_Lines;
	int m_nLineHeight;
	HFONT m_Font;
	HBRUSH m_BackgroundBrush;
	const static wstring m_LeftMargin;

public:

	void DisplayString( LPCOLESTR str);

	// IDisplay functions
	virtual IDisplay& operator<<( LPCOLESTR s) { DisplayString(s); return *this; }
	virtual IDisplay& operator<<( const wstring& s) { DisplayString(s.c_str()); return *this; }

	CTextBox() : m_nLineHeight(1)
	{
	}

	~CTextBox()
	{
	}

	void Clear();

	enum { IDD = IDD_TEXTBOX };

BEGIN_MSG_MAP(CTextBox)
	MESSAGE_HANDLER(WM_PAINT, OnPaint)
	MESSAGE_HANDLER(WM_VSCROLL, OnScroll)
	MESSAGE_HANDLER(WM_SIZE, OnSize)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
END_MSG_MAP()

//	DECLARE_WND_SUPERCLASS(NULL, L"EDIT")

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

private:

	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnScroll(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnSize(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
};

#endif //__TEXTBOX_H_
