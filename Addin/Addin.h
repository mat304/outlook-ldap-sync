#ifndef __ADDIN_H_
#define __ADDIN_H_

#define ADDIN_VERSION L"1.0.37"

extern _ATL_FUNC_INFO OnClickButtonInfo;
extern _ATL_FUNC_INFO OnItemsEventsInfo;

typedef IDispEventSimpleImpl<1, CAddin, &Office::DIID__CommandBarButtonEvents> CommandButton1Events;
typedef IDispEventSimpleImpl<2, CAddin, &Office::DIID__CommandBarButtonEvents> CommandMenuEvents;
//typedef IDispEventSimpleImpl<3, CAddin, &Outlook::DIID_ItemsEvents> ItemsEvents;

/////////////////////////////////////////////////////////////////////////////
// CAddin
class ATL_NO_VTABLE CAddin : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAddin, &CLSID_Addin>,
	public ISupportErrorInfo,
	public IDispatchImpl<IAddin, &IID_IAddin, &LIBID_GROUPWARELib>,
	public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects>,
	public CommandButton1Events,
	public CommandMenuEvents /*,
	public ItemsEvents */
{
	CStore* m_pOutlookStore;
	CCfg* m_Cfg;
	CDisplay m_DebugLog;
	CTextBox * m_DebugWindow;
	CLogger * m_DebugFile;
	CSyncDialog* m_SyncDlg;
	Outlook::_Application * m_spApplication;
	int m_OutlookMajorVersion;
	bool m_FirstUpgradeCall;

	CAutoUpgrade* m_AutoUpgrade;
	static LPCOLESTR m_VersionStr; 
	int m_VersionNum; // Version in numerical format

	void DisplayMsg( LPCOLESTR str);
	void ConfigureOutlookExplorer(CComPtr<Outlook::_Explorer>& spExplorer);

	CComQIPtr<Office::_CommandBarButton> m_spSyncButton;
	CComQIPtr<Office::_CommandBarButton> m_spMenuButton;

	bool CommandButton1EventsEnabled;
	bool CommandMenuEventsEnabled;
	void __stdcall OnClickButton(IDispatch * /*Office::_CommandBarButton**/ Ctrl, VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickMenu(IDispatch * /*Office::_CommandBarButton**/ Ctrl, VARIANT_BOOL * CancelDefault);

	// Event handlers
// Can't seem to get these ones to work !
//	void _stdcall OnItemAdd(IDispatch * Item);
//	void _stdcall OnItemChange(IDispatch * Item);
//	void _stdcall OnItemRemove();

public:

BEGIN_SINK_MAP(CAddin)
SINK_ENTRY_INFO(1, Office::DIID__CommandBarButtonEvents, 0x01, OnClickButton, &OnClickButtonInfo)
SINK_ENTRY_INFO(2, Office::DIID__CommandBarButtonEvents, 0x01, OnClickMenu, &OnClickButtonInfo)
//SINK_ENTRY_INFO(3, Outlook::DIID_ItemsEvents, 0x0000f002, OnItemChange, &OnItemsEventsInfo)
//SINK_ENTRY_EX(3, Outlook::DIID_ItemsEvents, 0x0000f003, OnItemRemove)
END_SINK_MAP()

	CAddin() 
	    : m_pOutlookStore(NULL),
		m_Cfg(NULL),
		m_DebugWindow(NULL),
		m_DebugFile(NULL),
		m_SyncDlg(NULL),
		m_spApplication(NULL),
		m_FirstUpgradeCall(true),
		m_AutoUpgrade(NULL),
		CommandButton1EventsEnabled(false),
		CommandMenuEventsEnabled(false)
	{
	}

DECLARE_CLASSFACTORY()
DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CAddin)
	COM_INTERFACE_ENTRY(IAddin)
//DEL 	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY2(IDispatch, IAddin)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

	// _IDTExtensibility2
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);

	void UpgradeCheck(const bool QuickDownload = false);

	int GetOutlookMajorVersion() const { return m_OutlookMajorVersion; }

	// Return the major version part of a standard "a.b.c.d" version string
	int GetMajorVersion(LPCOLESTR VersionStr) const;

	// Return this addin's version number
	LPCOLESTR GetVersion();

	static void Quit();
};

#endif //__ADDIN_H_
