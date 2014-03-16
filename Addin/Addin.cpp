// Addin.cpp : Implementation of CAddin extra
#include "stdafx.h"
using namespace std;

static Outlook::_Application * m_QuitPtr; 

/////////////////////////////////////////////////////////////////////////////
// CAddin

STDMETHODIMP CAddin::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IAddin
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if ( ::InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

/******************************************************************
 * IDTExtensibility2 Implementation
 *
 ******************************************************************/
/******************************************************************
 * CMyAddin::OnConnection()
 *
 *  Called by Office app when the addin is being loaded. ConnectMode
 *  identifies whether the addin is being loaded:
 *
 *     0  -- after startup (by the end user)
 *     1  -- during startup (normal mode)
 *     2  -- from an external source (like a VBA macro)
 *     3  -- from the command line
 *
 *  Also passed to the addin are two dispatch pointers, one to the
 *  host application and the other to the object that represents
 *  this addin in the AddIns collection. OnConnection is where the
 *  addin saves these pointers for use in calling the host back. 
 *  The pointers should be released when OnDisconnection is called.
 *
 ******************************************************************/
STDMETHODIMP CAddin::OnConnection(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom)
{
    if (NULL == Application) return E_POINTER;

	// Save a reference to the application object...
	Application->QueryInterface(Outlook::IID__Application,(PVOID*)&m_spApplication);
	if (NULL == m_spApplication) return E_POINTER;

	m_QuitPtr = m_spApplication;

 // If we are connecting during startup, we should wait for OnStartupComplete
 // before modifying the user-interface and prompting the user. Otherwise, we
 // can call OnStartupComplete to do the work now...
    if (ConnectMode != ext_cm_Startup)
        OnStartupComplete(custom);

    return S_OK;
}

/******************************************************************
 * CMyAddin::OnDisconnection()
 *
 *  Called by Office app when the addin is being unloaded. The
 *  RemoveMode parameter will notify the Addin if it is being unloaded
 *  by the end user (1) or the application is shutting down (0).
 *
 ******************************************************************/
STDMETHODIMP CAddin::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	// If this is not because of host shutdown, make sure we have cleaned
	// up our command bar...
    if (RemoveMode != ext_dm_HostShutdown)
        OnBeginShutdown(custom);

	// Release the application pointer...
    if (NULL != m_spApplication) 
	{
		if(CommandButton1EventsEnabled)
			CHECKHR(CommandButton1Events::DispEventUnadvise((IDispatch*)m_spSyncButton));

		if(CommandMenuEventsEnabled)
			CHECKHR(CommandMenuEvents::DispEventUnadvise((IDispatch*)m_spMenuButton));

		m_spApplication->Release();
	}

    return S_OK;
}

LPCOLESTR CAddin::m_VersionStr = ADDIN_VERSION;
LPCOLESTR CAddin::GetVersion() 
{ 
#if 0
	wstring OurFileName( m_Cfg->GetDLLDir() + L"\\GroupWare.dll" );

	DWORD Handle;
	DWORD VersionLS = 0, VersionMS = 0;
	DWORD BufLen = GetFileVersionInfoSize( const_cast<LPWSTR>(OurFileName.c_str()), &Handle );
	if( BufLen > 0 ) {
		char * lpBuf = new char[BufLen];
		if( GetFileVersionInfo( const_cast<LPWSTR>(OurFileName.c_str()), NULL, BufLen, lpBuf ) ) {
			VS_FIXEDFILEINFO * FileInfo;
			UINT FileInfoLen;
			if( VerQueryValue( lpBuf, L"\\", (LPVOID*)&FileInfo, &FileInfoLen ) ) {
				VersionLS = FileInfo->dwFileVersionLS;
				VersionMS = FileInfo->dwFileVersionMS;
			}
		}
		delete[] lpBuf;
	}
#endif
	return m_VersionStr; 
}

STDMETHODIMP CAddin::OnStartupComplete(SAFEARRAY * * custom)
{

	// No explorer if you open outlook in inline mode (eg click on mailto: webpage link)
	CComPtr<Outlook::_Explorer> spExplorer = NULL;
	CHECKHR(m_spApplication->ActiveExplorer(&spExplorer));
	if( spExplorer == NULL )
		return S_OK;

	CComBSTR OutlookVersion;
	CHECKHR(m_spApplication->get_Version(&OutlookVersion));
	m_OutlookMajorVersion = GetMajorVersion(OutlookVersion.m_str);
	if( m_OutlookMajorVersion < 9 )
	{
		DisplayMsg(L"Versions of Outlook older than Outlook 2000 are not supported.\nPlease upgrade Outlook.");
		return S_OK;
	}

	// Start auto-upgrade as early as possible !
	m_AutoUpgrade = new CAutoUpgrade();
	m_VersionNum = m_AutoUpgrade->EnumVerStr(m_VersionStr);

	m_Cfg = new CCfg();
	if( !m_Cfg->OK() )
	{
		DisplayMsg( L"Unable to read user settings. Please reinstall or repair.");
		return S_OK;
	}

	if( !m_Cfg->RefreshDLLDir() )
	{
		DisplayMsg( L"Unable to determine product installation root.");
		return S_OK;
	}

	// Open debug log if we're debugging
	if( m_Cfg->GetLogFileName() != L"" )
	{
		if( m_Cfg->GetLogFileName() == L"win" )
		{
			m_DebugWindow = new CTextBox();
			m_DebugWindow->Create(GetActiveWindow());
			m_DebugWindow->MoveWindow( 40, 190, 400, 400 );
			m_DebugLog.RegisterDisplay(m_DebugWindow);
		}
		else
		{
			m_DebugFile = CLogger::Make(m_Cfg->GetLogFileName());
			m_DebugLog.RegisterDisplay(m_DebugFile);
		}
		m_DebugLog << L"Started GroupWare for Outlook debug log\n";
	}
	// .. otherwise dbg messages go to "null display"
	
	// Check for upgrades now. Download immediately if we crashed last time Outlook shutdown
	// Had to remove this because web client timeout is too long and outlook user interface is not
	// responsive until we return from OnConnection() fn. First upgrade check now done at first sync time.
	if(m_Cfg->GetABEND()) UpgradeCheck( true ); // well only do it if we abended last time

	m_Cfg->SetABEND(true);

	try {
		ConfigureOutlookExplorer(spExplorer);
	} catch ( _com_error& Err )
	{
		const wstring ErrMsg(Err.ErrorMessage());
		m_DebugLog << L"ConfigureOutlookExplorer COM error " << ErrMsg << L"\n";
		return S_OK;
	}

	m_pOutlookStore = new CStore( this, L"Personal Folders", L"Public Folders", m_Cfg );

	if(m_DebugWindow)
		m_pOutlookStore->RegisterDebugDisplay( m_DebugWindow );
	else if(m_DebugFile)
		m_pOutlookStore->RegisterDebugDisplay( m_DebugFile );

	if( !m_Cfg->ValidConfig() )
	{
		DisplayMsg( L"This is the first time you have used GroupWare on this PC.\n"
					L"Please review your configuration settings on the next screen.\n" );
		m_Cfg->SetupNewUser();
		m_Cfg->UserConfig(m_pOutlookStore);
	}

	// Start up the sync window. (This also does the background syncronization)
	m_SyncDlg = new CSyncDialog( m_pOutlookStore );
	m_SyncDlg->Create( GetActiveWindow() );

	return S_OK;
}

STDMETHODIMP CAddin::OnBeginShutdown(SAFEARRAY * * custom)
{

	if( m_Cfg ) {
	
		if( m_Cfg->OK() )
		{
			USES_CONVERSION;

			if( m_DebugWindow )
			{
				if( m_pOutlookStore ) m_pOutlookStore->UnRegisterDebugDisplay();
				m_DebugLog.UnRegisterDisplay();
				delete m_DebugWindow;
				m_DebugWindow = NULL;
			}

			if( m_SyncDlg ) 
			{
				if(m_SyncDlg->IsWindow()) 
					m_SyncDlg->DestroyWindow();
				//delete m_SyncDlg;
				m_SyncDlg = NULL;
			}

			CStoreArchive File(W2CA((m_Cfg->GetDTFileName()).c_str()),false);
			if( m_pOutlookStore && (!File.SaveHeader() || !m_pOutlookStore->Save(&File)))
				DisplayMsg( L"Failed to write deletion tracking file. Some deletion operations may be lost." );

			if( m_DebugFile )
			{
				m_DebugLog << L"Finished GroupWare for Outlook debug log\n\n\n";
				if(m_pOutlookStore) m_pOutlookStore->UnRegisterDebugDisplay();
				m_DebugLog.UnRegisterDisplay();
				delete m_DebugFile;
				m_DebugFile = NULL;
			}

			if( m_pOutlookStore ) delete m_pOutlookStore;

			m_Cfg->SetABEND(false); // We shut down cleanly
		}

		delete m_Cfg;

		m_Cfg = NULL;
	}

	if( m_AutoUpgrade )
	{
		delete m_AutoUpgrade;
		m_AutoUpgrade = NULL;
	}

	return S_OK;
}

void CAddin::DisplayMsg(LPCOLESTR str)
{
	MessageBox( GetActiveWindow(), str, L"GroupWare server", MB_OK);
}

int CAddin::GetMajorVersion(LPCOLESTR Str) const
{ 
	long int a,b,c,d;
	if(swscanf(Str,L"%ld.%ld.%ld.%ld",&a,&b,&c,&d)!=4)
		return 0;
	return a; 
}

void CAddin::ConfigureOutlookExplorer(CComPtr<Outlook::_Explorer>& spExplorer)
{
	CComPtr<Office::_CommandBars> spCmdBars = NULL; 
	CHECKHR(spExplorer->get_CommandBars(&spCmdBars));
	CHECKNOTNULL(spCmdBars);

	CComVariant vtTrue(VARIANT_TRUE);
	CComPtr<Office::CommandBar> spCmdBar = NULL;
	CComPtr<Office::CommandBar> spOldCmdBar = NULL;
	CComBSTR OurBarName(L"GroupWare");
	CComBSTR OldBarName(L"GroupWare Server"); 

	for( int i = 0; i < spCmdBars->Count; i++)
	{
		CComVariant index(i);
		CComPtr <Office::CommandBar> bar;
		if( FAILED( spCmdBars->get_Item( index, &bar)) || !bar) continue;

		CComBSTR barName;
		CHECKHR(bar->get_Name(&barName));
		if( barName == OldBarName ) {
			spOldCmdBar = bar;
		} else if( barName == OurBarName ) {
			spCmdBar = bar;
		}
	}

	// Migrate old command bar
	if( spOldCmdBar && !spCmdBar) {
		spCmdBar = spOldCmdBar;
		spCmdBar->PutName(OurBarName.m_str);
	} else if( spOldCmdBar && spCmdBar) {
		spOldCmdBar->Delete();
	}

	if( !spCmdBar ) {
		CComVariant vName(OurBarName);
		CComVariant vPos(Office::msoBarTop); 
		spCmdBar = spCmdBars->Add(vName, vPos, vtMissing, vtMissing);
		CHECKNOTNULL( spCmdBar );
		spCmdBar->put_RowIndex(spCmdBars->Count);
	}

	spCmdBar->put_Enabled(VARIANT_TRUE);
	spCmdBar->put_Visible(VARIANT_TRUE);

	// Now add the Synchronize button
	CComPtr < Office::CommandBarControls> spBarControls = NULL;
	spBarControls = spCmdBar->GetControls();
	CHECKNOTNULL( spBarControls );

	CComVariant vToolBarType(Office::msoControlButton);
	CComPtr < Office::CommandBarControl> m_spSyncButtonControl = NULL; 
	m_spSyncButtonControl = spBarControls->Add(vToolBarType, vtMissing, vtMissing, vtMissing, vtTrue); 
	CHECKNOTNULL(m_spSyncButtonControl);

	m_spSyncButton = m_spSyncButtonControl;

	// Set the button style
	m_spSyncButton->PutStyle(Office::msoButtonIconAndCaption);
	m_spSyncButton->PutVisible(VARIANT_TRUE); 
	m_spSyncButton->PutFaceId(688/*2131*/);
	m_spSyncButton->PutCaption(OLESTR("Synchronize")); 
	m_spSyncButton->PutEnabled(VARIANT_TRUE);
	m_spSyncButton->PutTooltipText(OLESTR("Synchronize with GroupWare Server")); 
	m_spSyncButton->PutTag(OLESTR("Synchronize")); 

#ifdef ATTEMPT_PASTEFACE

	// This hardly ever works. Best to use built in Outlook FaceIds

	HBITMAP hBmp =(HBITMAP)LoadImage(_Module.GetResourceInstance(),
		MAKEINTRESOURCE(IDR_LOGO),IMAGE_ICON,32,32,LR_MONOCHROME);

	// Risky operation putting icon onto button !
	if( OpenClipboard(GetActiveWindow()) && EmptyClipboard() )
	{
		SetClipboardData(CF_BITMAP, hBmp);

		try {
			m_spSyncButton->PasteFace(); 
		} catch (  _com_error& )
		{
			// If this fails then just continue with no button icon
		}

		CloseClipboard();
	}

#endif

	CHECKHR(CommandButton1Events::DispEventAdvise((IDispatch*)m_spSyncButton));
	CommandButton1EventsEnabled = true;

	CComPtr<Office::CommandBar> spMenuBar;
	CHECKHR(spCmdBars->get_ActiveMenuBar(&spMenuBar));

	// get menu as CommandBarControls 
	CComPtr < Office::CommandBarControls> spCmdCtrls = NULL;
	spCmdCtrls = spMenuBar->GetControls(); 
	CHECKNOTNULL(spCmdCtrls);

	// we want to add a menu entry to Outlook's 6th(Tools) menu 	//item
	CComVariant vItem(5);

	CComPtr < Office::CommandBarControl> spCmdCtrl = NULL;
	spCmdCtrl = spCmdCtrls->GetItem(vItem);
	CHECKNOTNULL(spCmdCtrl);

	IDispatchPtr spDisp = NULL;
	spDisp = spCmdCtrl->GetControl(); 
	CHECKNOTNULL(spDisp);

	// a CommandBarPopup interface is the actual menu item
	CComQIPtr < Office::CommandBarPopup> ppCmdPopup(spDisp);  
	CHECKNOTNULL(ppCmdPopup);

	CComPtr < Office::CommandBarControls> spCmdBarCtrls = NULL; 
	spCmdBarCtrls = ppCmdPopup->GetControls();
	CHECKNOTNULL(spCmdBarCtrls);

	CComVariant vMenuType(1); // type of control - menu
	CComVariant vMenuPos(6);  
	CComVariant vMenuEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
	CComVariant vMenuShow(VARIANT_TRUE); // menu should be visible
	CComVariant vMenuTemp(VARIANT_TRUE); // menu is temporary		
		
	CComPtr < Office::CommandBarControl> spNewMenu = NULL;
	// now create the actual menu item and add it
	spNewMenu = spCmdBarCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp); 
	CHECKNOTNULL(spNewMenu);

	spNewMenu->PutCaption(OLESTR("GroupWare Server"));
	spNewMenu->PutEnabled(VARIANT_TRUE);
	spNewMenu->PutVisible(VARIANT_TRUE); 
	
	//we'd like our new menu item to look cool and display
	// an icon. Get menu item  as a CommandBarButton
	m_spMenuButton = spNewMenu;
	CHECKNOTNULL(m_spMenuButton);

	m_spMenuButton->PutStyle(Office::msoButtonIconAndCaption);
//	m_spMenuButton->put_FaceId(688);
	
#ifdef ATTEMPT_PASTEFACE

	// Risky operation putting icon onto button !
	if( OpenClipboard(GetActiveWindow()) && EmptyClipboard() )
	{
		SetClipboardData(CF_BITMAP, hBmp);

		try {
			m_spMenuButton->PasteFace(); 
		} catch (  _com_error& )
		{
			// If this fails then just continue with no button icon
		}

		CloseClipboard();
	}

//	DeleteObject(hBmp);	

#endif 

	CHECKHR(CommandMenuEvents::DispEventAdvise((IDispatch*)m_spMenuButton));
	CommandMenuEventsEnabled = true;

	// Now add some event handlers for when Contacts get added or deleted
//	CComPtr<Outlook::_Items> pItems;
//	CHECKHR(pFolder->get_Items(&pItems));
//	CHECKHR(ItemsEvents::DispEventAdvise((IDispatch*)pItems));
}

_ATL_FUNC_INFO OnClickButtonInfo ={CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BYREF | VT_BOOL}};

// Menu bar button has been pressed
void __stdcall CAddin::OnClickButton(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	if( m_SyncDlg ) m_SyncDlg->DoUserSync();
}

// Menu tools dropdown button has been pressed
void __stdcall CAddin::OnClickMenu(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	if( m_Cfg ) m_Cfg->UserConfig(m_pOutlookStore);
}

//_ATL_FUNC_INFO OnItemsEventsInfo = {CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};

// Event called whenever the Contact item Save or Save As buttons are pressed
//void __stdcall CAddin::OnItemAdd( IDispatch * p)
//{
//	MessageBox(NULL, (LPCTSTR)"Wrote item", (LPCTSTR)"OnClickButton", MB_SETFOREGROUND);
//}

// Event called whenever the Contact item Save or Save As buttons are pressed
//void __stdcall CAddin::OnItemRemove()
//{
//	MessageBox(NULL, (LPCTSTR)"Wrote item", (LPCTSTR)"OnClickButton", MB_SETFOREGROUND);
//}

// Event called whenever the Contact item Save or Save As buttons are pressed
//void __stdcall CAddin::OnItemChange(IDispatch * p)
//{
//	MessageBox(NULL, (LPCTSTR)"Wrote item", (LPCTSTR)"OnClickButton", MB_SETFOREGROUND);
//}

// Event called whenever a Contact item is deleted
//void __stdcall CAddin::OnItemDelete(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
//{
//	MessageBox(NULL, (LPCTSTR)"Wrote item", (LPCTSTR)"OnClickButton", MB_OK);
//}

// This does everything required to auto-upgrade the product. 
// It should be called at startup & about every 5mins afterwards
inline const wchar_t * BoolChar(const bool b) { return b?L"T":L"F"; }
void CAddin::UpgradeCheck(const bool QuickDownload )
{
	const wstring DLLFile( m_Cfg->GetDLLDir()+L"\\GroupWare.dll" );
	const wstring BackupDLLFile( m_Cfg->GetDLLDir()+L"\\GroupWare.dll.bak" );

	if( m_Cfg->GetUpgradeUrl() == L"" )
	{
		m_DebugLog << L"\n\nUpgradeCheck user disable (" << BoolChar(m_FirstUpgradeCall) << BoolChar(QuickDownload) << L")\n";
		return;
	}

	m_DebugLog << L"\n\nUpgradeCheck STARTED (" << BoolChar(m_FirstUpgradeCall) << BoolChar(QuickDownload) << L")\n";

	// Delete the backup file if this is the first time we have been called since Outlook started
	//  .. otherwise the presence of a backup file means that upgrades are disabled (until user restarts outlook)
	bool BackupFile = false;
	WIN32_FIND_DATA FindFileData;
	HANDLE hFind = FindFirstFile( BackupDLLFile.c_str(), &FindFileData );
	if (hFind != INVALID_HANDLE_VALUE) 
	{
		// Find first file returns the filename without the directory bit
		const wstring File(FindFileData.cFileName);
		if( wstring(L"GroupWare.dll.bak") == File ) BackupFile = true;
		FindClose(hFind);
	}

	if( BackupFile && m_FirstUpgradeCall )
	{
		if( DeleteFile(BackupDLLFile.c_str()) == TRUE )
			m_DebugLog << L"Deleted " << BackupDLLFile << L"\n";
	}
	else if( BackupFile )
	{
		m_DebugLog << L"Upgrade disabled\n";
		return;
	}
	
	m_FirstUpgradeCall = false;

	const bool UsingBITS = (!QuickDownload && m_AutoUpgrade->BitsAvailable());
	GUID UpgradeJob;

	if( UsingBITS ) 
	{
		// Try to complete any existing BITS jobs
		UpgradeJob = m_Cfg->GetUpgradeJob();
		if( UpgradeJob != GUID_NULL )
		{
			bool InProgress, CompletedFailed, CompletedSuccess;
			if( m_AutoUpgrade->CheckDownloadProgress( UpgradeJob, InProgress, CompletedFailed, CompletedSuccess ) == 0 )
			{
				m_DebugLog << L"BITS upgrade status " << BoolChar(InProgress) << BoolChar(CompletedFailed) << BoolChar(CompletedSuccess) << L"\n";
				if( CompletedSuccess || CompletedFailed )
					m_Cfg->SetUpgradeJob(GUID_NULL);
			}
			else
				m_DebugLog << L"BITS CheckDownloadProgress FAILED\n";
		}
	}

	const wstring UpgradePattern(L"gw4ol-[0-9]+\\.[0-9]+\\.[0-9]+\\.upgrade");

	m_DebugLog << L"Checking " << m_Cfg->GetUpgradeUrl() 
			   << L" for upgrade files like " << UpgradePattern  << L"\n";

	int BestVersion = m_VersionNum;
	wstring BestFileName;
	list<wstring> Matches;
	if( m_AutoUpgrade->Check( m_Cfg->GetUpgradeUrl() + _T("/index.html"), UpgradePattern, Matches ) == 0 ) 
	{
		for( list<wstring>::const_iterator i = Matches.begin(); i!=Matches.end(); ++i) 
		{
			const wstring& FileName(*i);
			const int Version = m_AutoUpgrade->EnumFileVerStr(FileName);
			if( Version > BestVersion ) 
			{
				BestVersion = Version;
				BestFileName = FileName;
			}
		}
	}

	if( BestVersion > m_VersionNum ) 
	{
		const wstring DownloadFileUrl( m_Cfg->GetUpgradeUrl()+L"/"+BestFileName );
		const wstring LocalFile( m_Cfg->GetDLLDir()+L"\\"+BestFileName );
		
		// Check if upgrade has been downloaded
		bool FileDownloaded = false;
		hFind = FindFirstFile( LocalFile.c_str(), &FindFileData );
		if (hFind != INVALID_HANDLE_VALUE) 
		{	
			// Find first file returns the filename without the directory bit
			const wstring File(FindFileData.cFileName);
			if( BestFileName == File ) FileDownloaded = true;
			FindClose(hFind);
		}

		if( FileDownloaded )
			m_DebugLog << L"Downloaded upgrade " << LocalFile << L"\n";
		else
			m_DebugLog << L"Want to download upgrade " << DownloadFileUrl << L"\n";

		if( !FileDownloaded )
		{
			if( UsingBITS )
			{
				if( ( m_Cfg->GetUpgradeJob() == GUID_NULL )
				 && ( m_AutoUpgrade->StartDownload( DownloadFileUrl, LocalFile, UpgradeJob ) == 0 ) )
				{
					m_DebugLog << L"Started BITS download from " << DownloadFileUrl << L" to " << LocalFile << L"\n";
					m_Cfg->SetUpgradeJob(UpgradeJob);
				}
			}
			else
			{
				m_DebugLog << L"Started quick download from " << DownloadFileUrl << L" to " << LocalFile << L"\n";
				DisplayMsg(L"An upgrade for this product has been detected and will be downloaded now.");
				if( m_AutoUpgrade->Download( DownloadFileUrl, LocalFile ) )
					DisplayMsg(L"Failed to download upgrade");
				else
					FileDownloaded = true;
			}
		}
		if( FileDownloaded )
		{
			m_DebugLog << L"Swapping " << DLLFile << L" with " << LocalFile << L"\n";
			if( ( MoveFile(DLLFile.c_str(),BackupDLLFile.c_str()) == TRUE )
				&& ( MoveFile(LocalFile.c_str(),DLLFile.c_str()) == TRUE ) )
				DisplayMsg( L"Successfully downloaded product upgrade. Upgrade will take effect when Outlook is next restarted." );
			else
				DisplayMsg( L"Upgrade failed. Please reinstall GroupWare Addin." );
		}
	}
	m_DebugLog << L"Finished UpgradeCheck\n";
}

void CAddin::Quit()
{
	m_QuitPtr->Quit();
}
