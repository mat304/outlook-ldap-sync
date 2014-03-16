#ifndef __CFGDIALOG_H_
#define __CFGDIALOG_H_

class CCfgDialog : 
	public CAxDialogImpl<CCfgDialog>
{
	CCfg* m_Cfg;
	CStore * m_Store;
	HBITMAP m_Logo;

	bool DecodeUserID( const wstring userid, wstring& userdn, wstring& basedn ) const;
	const wstring EncodeUserID( const wstring userdn ) const;

public:

	CCfgDialog( CCfg* cfg, CStore * store ) : m_Cfg(cfg), m_Store(store) {}
	~CCfgDialog() {}

	enum { IDD = IDD_CFGDIALOG };

BEGIN_MSG_MAP(CCfgDialog)
	MESSAGE_HANDLER(WM_PAINT, OnPaint)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	COMMAND_ID_HANDLER(IDOK, OnOK)
	COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
END_MSG_MAP()

	LRESULT OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
};

#endif //__CFGDIALOG_H_
