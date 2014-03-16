#include "stdafx.h"
using namespace std;

CStore::CStore(CAddin * Addin, const wstring StoreName, const wstring PublicSubFolder, CCfg* cfg )
	:m_Addin(Addin)
	,m_StoreName(StoreName)
	,m_PublicSubFolder(PublicSubFolder)
	,m_pMapiSession(NULL)
	,m_pMapiStore(NULL)
	,m_Cfg(cfg)
	,m_FirstSync(true)
	,m_pNameSpace(NULL)
	,m_pStoreFolders(NULL)
{

	// Pretend that we've already syncronized if its the first time a user has run GW4OL
	if( m_Cfg->GetFirstTime() )
		m_FirstSync = false;

	// This fails when using Win98. 
	// Looks like theres something missing in ole32.dll.
	CHECKHR(Init());

	SetApartmentToEnter();
}

CStore::~CStore(void)
{
	if( m_pMapiStore ) delete m_pMapiStore;
}

void CStore::ReadConfig()
{
	bool Invalidate = false;
	CHECKTRUE( m_Cfg->Refresh(Invalidate) );
	if(Invalidate)
	{
		map<wstring,CFolder*>::iterator i;
		for(i=m_Folders.begin(); i!=m_Folders.end(); ++i)
			i->second->InvalidateIndexes();
	}
}

void CStore::ScanLdapFolders( CLdapSession& sess
							, const wstring IPMClass
							, set<wstring>& Folders )
{
	CLdapSearch objSearch( sess );
	LPCOLESTR apszUserAttrsSD[] = { L"cn", NULL };
	wstring SearchExpr(L"(&(objectClass=OlFolder)(IPMClass=");
	SearchExpr += IPMClass;
	SearchExpr += L"))";

	list<wstring> BaseDNs;
	BaseDNs.push_back(m_Cfg->GetUserDN());
	BaseDNs.insert(BaseDNs.end(), m_GroupDNs.begin(), m_GroupDNs.end() );

	list<wstring>::iterator li;
	for( li=BaseDNs.begin(); li!=BaseDNs.end(); ++li)
	{
		wstring& BaseDN = *li;

		try
		{
			objSearch.Search( BaseDN.c_str()
						 ,LDAP_SCOPE_ONELEVEL
						 ,SearchExpr.c_str()
						 ,apszUserAttrsSD );

			CLdapObject objLdapEntry;
			while( objSearch.GetNextObject( objLdapEntry ) )
			{
				CLdapAttribute objLdapAttr;
				CLdapValue objValue;

				CHECKTRUE( objLdapEntry.GetAttribute( L"cn", objLdapAttr ) );
				CHECKTRUE(objLdapAttr.ValueCount() == 1 );
				objValue = objLdapAttr(); // assumes 1 value
				wstring Name((LPCOLESTR)objValue);
				Folders.insert(Name);
//				dbg << L"Found LDAP folder " << Name << L"\n";
			}
		} catch (CLdapException& err)
		{
			dbg << L"Failed to get folders index " << err.Description() << L"\n";
			throw CError( L"Failed to get folders index");
		}
	}
}

// This refreshes m_GroupDNs which is a list of basedns for the groups that user is a member of
//  and m_GroupMap which is a associative array of Public Folder name to basedn
void CStore::RefreshGroups( CLdapSession& sess )
{
	m_GroupDNs.clear();
	m_GroupMap.clear();

	CLdapSearch GroupSearch( sess );
	LPCOLESTR apszGroupAttrsSD[] = { L"group", L"GWPublicFolder", L"GWPublicFolderRW", NULL };
	
	try
	{
		GroupSearch.Search( m_Cfg->GetUserDN().c_str()
						,LDAP_SCOPE_BASE
						,NULL
						,apszGroupAttrsSD );

		CLdapObject GroupEntry;
		while( GroupSearch.GetNextObject( GroupEntry ) )
		{
			CLdapAttribute GroupAttr;
			CLdapAttribute FolderAttr;
			if( GroupEntry.GetAttribute( L"group", GroupAttr ) ) 
			{
				for( int i = 0; i < GroupAttr.ValueCount(); i++ )
				{
					CLdapValue GroupValue(GroupAttr[i]);
					wstring GroupDN = m_Cfg->GetGroupDN((LPCOLESTR)GroupValue);
					m_GroupDNs.push_back(GroupDN);

					CLdapSearch FolderSearch( sess );
					LPCOLESTR apszFolderAttrsSD[] = { L"GWPublicFolder", L"GWPublicFolderRW", NULL };
					FolderSearch.Search( GroupDN.c_str()
									,LDAP_SCOPE_BASE
									,NULL
									,apszFolderAttrsSD );
					
					CLdapObject FolderEntry;
					while( FolderSearch.GetNextObject( FolderEntry ) )
					{
						if( FolderEntry.GetAttribute( L"GWPublicFolder", FolderAttr ) ) 
						{
							for( int i = 0; i < FolderAttr.ValueCount(); i++ )
							{
								CLdapValue FolderValue(FolderAttr[i]);
								wstring FolderName((LPCOLESTR)FolderValue);
								m_GroupMap[FolderName] = GroupDN;
								m_PublicFolderAccessMap[FolderName] = PFA_READ_ACCESS;
							}
						}
						if( FolderEntry.GetAttribute( L"GWPublicFolderRW", FolderAttr ) )
						{
							for( int i = 0; i < FolderAttr.ValueCount(); i++ )
							{
								CLdapValue FolderValue(FolderAttr[i]);
								wstring FolderName((LPCOLESTR)FolderValue);
								m_GroupMap[FolderName] = GroupDN;
								m_PublicFolderAccessMap[FolderName] = PFA_WRITE_ACCESS;
							}
						}
					}
				}
			}
			if( GroupEntry.GetAttribute( L"GWPublicFolder", FolderAttr ) ) 
			{
				for( int i = 0; i < FolderAttr.ValueCount(); i++ )
				{
					CLdapValue FolderValue(FolderAttr[i]);
					wstring FolderName((LPCOLESTR)FolderValue);
					m_PublicFolderAccessMap[FolderName] = PFA_READ_ACCESS;
				}
			}
			if( GroupEntry.GetAttribute( L"GWPublicFolderRW", FolderAttr ) )
			{
				for( int i = 0; i < FolderAttr.ValueCount(); i++ )
				{
					CLdapValue FolderValue(FolderAttr[i]);
					wstring FolderName((LPCOLESTR)FolderValue);
					m_PublicFolderAccessMap[FolderName] = PFA_WRITE_ACCESS;
				}
			}
		}
	} catch (CLdapException& err )
	{
		dbg << L"Failed to get group list " << err.Description() << L"\n";
		throw CError( L"Failed to get group list");
	}
}

// This will return either the LDAP base DN in which to store entries for the Outlook Folder
// or it will return L"" in which case the user is not authorized to share the named folder
const wstring CStore::GetDNFromFolderName(const wstring Name)
{
	wstring SubFolder(Name);
	wstring TopFolder;
	
	// If folder name doesn't start with "Public Folders" then this is a personal folder
	// that needs to be stored under UserDN
	SplitFirst( SubFolder, TopFolder );
	if( TopFolder != m_PublicSubFolder )
		return L"cn=" + Name + L"," + m_Cfg->GetUserDN();

	// This is a public folder (because its name starts with "Public Folder")
	// but we only syncronize public folders that correspond to LDAP GWPublicFolder attrs
	map<wstring,wstring>::iterator mi = m_GroupMap.find(Name);
	if( mi == m_GroupMap.end() )
		return L"";

	return L"cn=" + Name + L"," + mi->second;
}

const wstring CStore::GetEntryIDFromFolderName(const wstring Name)
{
	CComPtr<Outlook::_Folders> pFolders = m_pStoreFolders;
	wstring FolderName(Name);
	wstring EntryID;
	bool Matched;
	do {
		wstring First;
		SplitFirst( FolderName, First );
		RemoveEscapeSequences( First );

		Matched = false;
		Outlook::MAPIFolder * pFolder = NULL;
		CHECKHR(pFolders->GetFirst(&pFolder));
		while( pFolder && !Matched )
		{
			CComBSTR BSTR = NULL;
			CHECKHR(pFolder->get_Name(&BSTR));
			if( BSTR && (wstring(BSTR.m_str)==First) )
			{	
				Matched = true;
				if( FolderName == L"" )
				{
					CComBSTR EntryIDBSTR = NULL;
					CHECKHR(pFolder->get_EntryID(&EntryIDBSTR));
					if( EntryIDBSTR ) EntryID = wstring(EntryIDBSTR.m_str);
				}
				else 
				{
					pFolders.Release();
					if( FAILED(pFolder->get_Folders(&pFolders)) )
						pFolders.p = NULL;
				}
				break;
			}
			pFolder->Release();
			CHECKHR(pFolders->GetNext(&pFolder));
		}
	} while ( Matched && (FolderName!=L"") && pFolders.p);

	if( FolderName != L"" || !Matched ) 
		return L"";

	return EntryID;
}

const bool CStore::IsPublicFolder( const wstring Name )
{
	wstring SubFolder(Name);
	wstring TopFolder;
	SplitFirst( SubFolder, TopFolder );
	return ( TopFolder == m_PublicSubFolder );
}

const bool CStore::HavePublicFolderWriteAccess( const wstring Name )
{
	if( !IsPublicFolder(Name) ) return false;
	map<wstring,DWORD>::iterator mi = m_PublicFolderAccessMap.find(Name);
	if( mi == m_PublicFolderAccessMap.end() ) return false;
	const DWORD& Access = mi->second; 
	return ((Access & PFA_WRITE_ACCESS ) > 0 ) ? true : false;
}

void CStore::SetFolderAccessRights( const wstring EntryID, const bool Read, const bool Modify, const bool Delete )
{
	// Unfortunately this doesn't work unless its an Exchange folder
	// TODO: Work out if there is some other way of doing it
	//CMAPIEntryID MAPIFolderEntryID;
	//MAPIFolderEntryID.Set( EntryID );
	//CMAPIFolder MAPIFolder(m_pMapiStore);
	//MAPIFolder.OpenFolder( MAPIFolderEntryID );
	//MAPIFolder.SetAccessRights(Read,Modify,Delete);
}

Outlook::OlDefaultFolders CStore::GetOlDefaultFolders( const wstring IPMClass)
{
	if( IPMClass == L"IPM.Contact" ) return Outlook::olFolderContacts;
	else if( IPMClass == L"IPM.Calendar" ) return Outlook::olFolderCalendar;
	else if( IPMClass == L"IPM.StickyNote" ) return Outlook::olFolderNotes;
	else if( IPMClass == L"IPM.Task" ) return Outlook::olFolderTasks;
	return Outlook::olFolderInbox; 
}

Outlook::OlItemType CStore::GetOlItemType( const wstring IPMClass)
{
	if( IPMClass == L"IPM.Contact" ) return Outlook::olContactItem;
	else if( IPMClass == L"IPM.Calendar" ) return Outlook::olAppointmentItem;
	else if( IPMClass == L"IPM.StickyNote" ) return Outlook::olNoteItem;
	else if( IPMClass == L"IPM.Task" ) return Outlook::olTaskItem;
	return Outlook::olMailItem;
}

void CStore::ScanLocalFolders( const wstring IPMClass
							 , Outlook::_Folders * pFolders
							 , set<wstring>& Folders
							 , const wstring ParentFolderName )
{
	Outlook::MAPIFolder* pMAPIFolder = NULL;
	for( pFolders->GetFirst(&pMAPIFolder); pMAPIFolder; pFolders->GetNext(&pMAPIFolder) )
	{
		CComPtr<Outlook::MAPIFolder> pFolder;
		pFolder.Attach(pMAPIFolder);
		CComBSTR NameBSTR = NULL;
		CHECKHR(pFolder->get_Name(&NameBSTR));
		if( NameBSTR )
		{
			wstring Name(NameBSTR.m_str);
			AddEscapeSequences( Name, wstring(1,m_FolderNameDelimiter) );
			Name = ParentFolderName + Name;
			CComBSTR MsgClass;
			if( SUCCEEDED(pFolder->get_DefaultMessageClass(&MsgClass)) && (wstring(MsgClass.m_str)==IPMClass) )
				Folders.insert(Name);
			Outlook::_Folders * pChildFolders = NULL;
			CHECKHR(pMAPIFolder->get_Folders(&pChildFolders));
			if( pChildFolders ) {
				ScanLocalFolders( IPMClass, pChildFolders, Folders, Name + m_FolderNameDelimiter );
				pChildFolders->Release();
			}
		}
	}
}

class FolderRuleFunctor: public CMatrix::IRule
{
	CLdapSession& m_sess;
	CDisplay& dbg;
	CStore* m_Store;
	const wstring m_IPMClass;
public:
	FolderRuleFunctor( CLdapSession& sess, CDisplay& out, const wstring IPMClass, CStore* Store ) 
		:m_sess(sess), dbg(out), m_Store(Store), m_IPMClass(IPMClass) 
	{}
	virtual ~FolderRuleFunctor() {}
	virtual void Do(  const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring FolderName )
	{
		try {
			m_Store->ApplyRule( Rule, LocalOp, LdapOp, FolderName, m_sess, m_IPMClass );
		} catch ( CError& e ) {
			dbg << L"FOLDER SYNCHRONIZATION ERROR " << e.GetDescription() << L"\n";
			return;
		} catch ( CLdapException& e ) {
			dbg << L"LDAP FOLDER SYNCHRONIZATION ERROR " << e.Description() << L"\n";
			return;
		}
	}
};

void CStore::ApplyRule( const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring FolderName, CLdapSession& sess, const wstring IPMClass )
{
	CANCELLATION_POINT

	bool WriteAllowed = ( !IsPublicFolder(FolderName) 
						|| ( IsPublicFolder(FolderName) && HavePublicFolderWriteAccess(FolderName) ) );

	dbg << L"\nFolder : " << FolderName << L"\n";
	dbg << L"-LDAP(" << CMatrix::ChangeTypeStr[LdapOp] << L") LOCAL(" << CMatrix::ChangeTypeStr[LocalOp] << L")\n";
	dbg << L"Action : " << CMatrix::RuleStr[Rule] << L"\n";
	if( !WriteAllowed ) dbg << L"[No write]\n";

	switch( Rule ) 
	{
	case CMatrix::RTCOC:
		AddLdapFolder(FolderName, sess, IPMClass);
		if( !AddFolderSyncronizer( IPMClass, FolderName ) )
		{
			DeleteLdapFolder(FolderName, sess);
			throw CError(L"Failed to add folder "+FolderName);
		}
		break;
	case CMatrix::RTDL:
		if( WriteAllowed ) DeleteLdapFolder(FolderName, sess);
		DeleteFolderSyncronizer(FolderName);
		break;
	case CMatrix::RTDO: 
		DeleteLocalFolder(FolderName );
		DeleteFolderSyncronizer(FolderName);
		break;
	case CMatrix::RTCLC: 
		AddLocalFolder( FolderName, FolderName, IPMClass, m_pStoreFolders );
		if( !AddFolderSyncronizer( IPMClass, FolderName ) )
		{
			DeleteLocalFolder(FolderName );
			throw CError(L"Failed to add folder "+FolderName);
		}
		break;
	case CMatrix::RTCMR:
		if( !AddFolderSyncronizer( IPMClass, FolderName ) )
			throw CError(L"Failed to add folder "+FolderName);
		break;
	case CMatrix::RTN:
		break;
	default:
		CHECKTRUE(false); // Bug if we get here
		break;
	}
	CMatrix::CorrectIndex( Rule, FolderName, LdapOp, LocalOp, m_LdapFolders, m_LocalFolders);

	// Now syncronise the entries in the folder (unless folder was deleted !)
	map<wstring,CFolder*>::iterator mi = m_Folders.find(FolderName);
	if( mi != m_Folders.end() )
	{
		stat << L"Synchronizing " << FolderName << L" ";
		if( mi->second->Syncronize(	 sess
								   , stat
								   , dbg
								   , m_pNameSpace
								   , m_pMapiStore
								   , WriteAllowed ) )
			stat << L"OK\n";
		else 
			stat << L"\n";

		// Write protect public folders
//		if( IsPublicFolder( FolderName ) ) 
//			SetFolderAccessRights( GetEntryIDFromFolderName(FolderName), true, false, false ); 
	}
}

void CStore::Syncronize( const wstring IPMClass, CLdapSession& sess )
{	
	set<wstring>::iterator si, tsi;

	set<wstring> LdapFolders, LdapAdded, LdapDeleted, LdapModified, LdapUnModified;
	ScanLdapFolders( sess, IPMClass, LdapFolders );

	// Quietly forget any public folders that we aren't authorized to share
	for( tsi=si=LdapFolders.begin(); si!=LdapFolders.end(); tsi=si )
	{
		++si;
		if( GetDNFromFolderName(*tsi) == L"" )
			LdapFolders.erase(tsi);
	}
	for( tsi=si=m_LdapFolders.begin(); si!=m_LdapFolders.end(); tsi=si )
	{
		++si;
		if( GetDNFromFolderName(*tsi) == L"" )
		{
			DeleteFolderSyncronizer(*tsi);
			m_LdapFolders.erase(tsi);
		}
	}
	
	// Entries that existed last time but don't exist now should be deleted
	set_difference(	  m_LdapFolders.begin()
					, m_LdapFolders.end()
					, LdapFolders.begin()
					, LdapFolders.end()
					, inserter(LdapDeleted, LdapDeleted.begin()) );
	
	// Prevent deletion of standard outlook folders
	for( tsi=si=LdapDeleted.begin(); si!=LdapDeleted.end(); tsi=si )
	{
		++si;
		const wstring& FldrName(*tsi);
		if( FldrName  == L"Contacts" ) {
			m_LdapFolders.erase(FldrName);
			LdapDeleted.erase(FldrName);
		}
	}

	// Entries that didn't exist last time and exist now should be added
	set_difference(   LdapFolders.begin()
					, LdapFolders.end()
					, m_LdapFolders.begin()
					, m_LdapFolders.end()
					, inserter(LdapAdded, LdapAdded.begin()) );

	// Entries that existed last time and still exist this time
	set_intersection( LdapFolders.begin()
					, LdapFolders.end()
					, m_LdapFolders.begin()
					, m_LdapFolders.end()
					, inserter(LdapUnModified, LdapUnModified.begin()) );

	set<wstring> LocalFolders, LocalAdded, LocalDeleted, LocalUnModified, LocalModified;
	ScanLocalFolders( IPMClass, m_pStoreFolders, LocalFolders );

	// Quietly forget any public folders that we aren't authorised to share
	for( tsi=si=LocalFolders.begin(); si!=LocalFolders.end(); tsi=si )
	{
		++si;
		if( GetDNFromFolderName(*tsi) == L"" )
			LocalFolders.erase(tsi);
	}
	for( tsi=si=m_LocalFolders.begin(); si!=m_LocalFolders.end(); tsi=si )
	{
		++si;
		if( GetDNFromFolderName(*tsi) == L"" )
		{
			DeleteFolderSyncronizer(*tsi);
			m_LocalFolders.erase(tsi);
		}
	}

	// Entries that existed last time but don't exist now should be deleted
	set_difference(	  m_LocalFolders.begin()
					, m_LocalFolders.end()
					, LocalFolders.begin()
					, LocalFolders.end()
					, inserter(LocalDeleted, LocalDeleted.begin()) );
	
	// Entries that didn't exist last time and exist now should be added
	set_difference(   LocalFolders.begin()
					, LocalFolders.end()
					, m_LocalFolders.begin()
					, m_LocalFolders.end()
					, inserter(LocalAdded, LocalAdded.begin()) );

	// Entries that existed last time and still exist this time
	set_intersection( LocalFolders.begin()
					, LocalFolders.end()
					, m_LocalFolders.begin()
					, m_LocalFolders.end()
					, inserter(LocalUnModified, LocalUnModified.begin()) );

	CMatrix Matrix;
	Matrix.Intersection(  &LocalAdded, &LocalModified, &LocalDeleted, &LocalUnModified
						, &LdapAdded, &LdapModified, &LdapDeleted, &LdapUnModified
						, FolderRuleFunctor(sess,dbg,IPMClass,this) );
}

void CStore::PromptUserForPassword()
{
	MessageBox(NULL,L"Authentication failure.\nPlease check your userid and password are correct.", L"GroupWare Server", MB_OK);
	m_Cfg->UserConfig(this);
}

CCPLWinAtomicCounter CStore::TerminateAllSyncThreads; 

// The thread fn()
void CStore::Proc()
{	
	TerminateAllSyncThreads.Exchange(0);

	TRACE("Thread started\n");

//	stat << L"Started server synchronization\n";

	// Check essential config settings
	if( (m_Cfg->GetServerName().size()==0)
		|| (m_Cfg->GetUserPassword().size()==0) )
	{
		PromptUserForPassword();
		return;
	}

	try {

		m_pMapiSession = new CMAPISession();
		wstring CurrentUserProfile;
		CHECKHR(m_pMapiSession->GetCurrentUserProfile( CurrentUserProfile ));
		CHECKHR(m_pMapiSession->Logon( CurrentUserProfile, L"" /* No password ! */ ));

		CMAPIEntryID PersonalFoldersStore;
		m_pMapiStore = new CMAPIStore(m_pMapiSession);
		CHECKHR(m_pMapiStore->GetDefaultStore( PersonalFoldersStore ));
		CHECKHR(m_pMapiStore->OpenStore( PersonalFoldersStore));

		m_StoreID = CComBSTR( PersonalFoldersStore.GetStr().c_str() );

		CComPtr<Outlook::_Application> pApp;
		HRESULT hr = ::CoCreateInstance( Outlook::CLSID_Application, NULL, CLSCTX_LOCAL_SERVER,
				Outlook::IID__Application, (LPVOID*)&pApp);

		// First level of folder hierarchy corresponds to MAPI Stores (PST file)
		// Get default MAPI store which corresponds to "Personal folders" (unless its been renamed!)
		CComPtr<Outlook::_Folders> pFolders;
		CHECKHR(pApp->GetNamespace(_bstr_t("MAPI"), &m_pNameSpace));
		CHECKHR(m_pNameSpace->get_Folders(&pFolders));
		if( pFolders ) 
		{
			Outlook::MAPIFolder* pMAPIFolder = NULL;
			for( pFolders->GetFirst(&pMAPIFolder); pMAPIFolder; pFolders->GetNext(&pMAPIFolder) )
			{
				CComBSTR StoreID = NULL;
				CHECKHR(pMAPIFolder->get_StoreID(&StoreID));
				if( StoreID && (StoreID == m_StoreID) )
				{
					CHECKHR(pMAPIFolder->get_Folders(&m_pStoreFolders));
				}
				pMAPIFolder->Release();
			}
		}

		if( !m_pStoreFolders )
			throw CError(L"Failed to find MAPI store");

		// Start up an LDAP session
		CLdapSession sess( m_Cfg->GetServerName().c_str() );

		dbg << L"Connecting to GroupWare server " << m_Cfg->GetServerName() << L"\n";
		stat << L"Connecting to GroupWare server (" << m_Cfg->GetServerName() << L") ...";

		CLdapCredentials creds( m_Cfg->GetUserDN().c_str(), m_Cfg->GetUserPassword().c_str() );
		creds.SetAuthType( CLdapCredentials::AuthTypeSimple);
		CLdapError e = sess.Bind( creds);
		stat << L"\n";
		if( (e==CLdapError(LDAP_TIMEOUT)) || (e==CLdapError(LDAP_SERVER_DOWN)) )
		{
			throw CError(L"Failed to contact server");
		} 
		else if( e==CLdapError(LDAP_INVALID_CREDENTIALS) )
		{
			PromptUserForPassword();
			throw CError(L"Bad userid/password");
		}
		else if ( !e.Success() )
		{
			throw CError(e.Description());
		}

		dbg << L"CONNECTED\n";

		RefreshGroups(sess);

		CANCELLATION_POINT

		if( m_FirstSync )
		{
			dbg << L"First synchronization\n";
			USES_CONVERSION;
			CStoreArchive File(W2CA((m_Cfg->GetDTFileName()).c_str()));
			if( !File.HeaderOK() || !Load(&File) )
				dbg << L"WARNING: Failed to open DT file so merging LDAP and local entries\n";
			m_FirstSync = false;
		}

		Syncronize( L"IPM.Contact", sess );

	}
	catch( CError& e ) {
		stat << L"ERROR: " << e.GetDescription() << L"\n";
		dbg << L"Synchronization failed: " << e.GetDescription() << L"\n";
	}
//	catch(...) {
//		dbg << L"Synchronization failed due to unknown exception\n";
//	}

	if( m_pStoreFolders ) {
		m_pStoreFolders->Release();
		m_pStoreFolders = NULL;
	}

	if( m_pNameSpace ) {
		m_pNameSpace->Release();
		m_pNameSpace = NULL;
	}

	if( m_pMapiStore ) {
		delete m_pMapiStore;
		m_pMapiStore = NULL;
	}

	if( m_pMapiSession ) {
		delete m_pMapiSession;
		m_pMapiSession = NULL;
	}

	m_Cfg->SetFirstTime( false );

	TRACE("Thread ended\n");

	dbg << L"Synchronization completed\n";
	stat << L"Synchronization completed\n";
}

CFolder* CStore::AddFolderSyncronizer( const wstring IPMClass, const wstring FolderName )
{
	map<wstring,CFolder*>::iterator mi = m_Folders.find(FolderName);
	if( mi != m_Folders.end() ) return mi->second;

	wstring BaseDN = GetDNFromFolderName( FolderName );
	if(BaseDN == L"") return NULL;

	wstring FolderEntryID = GetEntryIDFromFolderName( FolderName );
	if( FolderEntryID == L"" ) return NULL;

	CFolder * NewFolder = new CFolder(    this
										, FolderName
										, FolderEntryID
										, IPMClass
										, GetOlDefaultFolders(IPMClass)
										, GetOlItemType(IPMClass)
										, BaseDN 
										, m_StoreID );
	
	m_Folders[FolderName] = NewFolder;
	
	return NewFolder;
}

void CStore::DeleteFolderSyncronizer( const wstring FolderName )
{
	map<wstring,CFolder*>::iterator mi = m_Folders.find(FolderName);
	if( mi == m_Folders.end() ) return;
	delete mi->second;
	m_Folders.erase(mi);
}

void CStore::AddLocalFolder( const wstring FullName, wstring Name, const wstring IPMClass, Outlook::_Folders * pFolders )
{
	wstring First;
	wstring EntryID;
	SplitFirst( Name, First );
	RemoveEscapeSequences( First );

	// First check if a folder with this name already exists
	Outlook::_Folders * pChildFolders = NULL;
	Outlook::MAPIFolder * pFolder = NULL;
	CHECKHR(pFolders->GetFirst(&pFolder));
	while( pFolder )
	{
		CComBSTR NameBSTR = NULL;
		CHECKHR(pFolder->get_Name(&NameBSTR));
		if( NameBSTR && (wstring(NameBSTR.m_str)==First) ) {
			if( pChildFolders ) pChildFolders->Release();
			CHECKHR(pFolder->get_Folders(&pChildFolders));
		}

		CComBSTR EntryIDBSTR = NULL;
		CHECKHR(pFolder->get_EntryID(&EntryIDBSTR));
		if( EntryIDBSTR ) EntryID = wstring(EntryIDBSTR.m_str);

		pFolder->Release();
		CHECKHR(pFolders->GetNext(&pFolder));
	}

	// If a folder with this name doesn't exist then add one
	if( !pChildFolders )
	{
		if( IsPublicFolder( FullName ) ) SetFolderAccessRights( EntryID, true, true, true ); // EntryID of parent folder
		CHECKHR( pFolders->Add( CComBSTR(First.c_str()), CComVariant(GetOlDefaultFolders(IPMClass)), &pFolder ) ); 
		if( IsPublicFolder( FullName ) ) SetFolderAccessRights( EntryID, true, false, false ); // EntryID of parent folder
		CHECKHR( pFolder->get_Folders(&pChildFolders) );
	}

	CHECKNOTNULL(pChildFolders);

	if( Name != L"" ) 
		AddLocalFolder( FullName, Name, IPMClass, pChildFolders );

	pChildFolders->Release();
}

void CStore::DeleteLocalFolder( const wstring FolderName )
{
	CFolder * pFolder = m_Folders[FolderName];
	CComPtr<Outlook::MAPIFolder> spMAPIFolder;
	CHECKHR( m_pNameSpace->GetFolderFromID( CComBSTR(pFolder->GetEntryID().c_str())
											, CComVariant(m_StoreID)
											, &spMAPIFolder ) );
	CHECKHR(spMAPIFolder->Delete());
}

void CStore::AddLdapFolder( const wstring Name, CLdapSession& sess, const wstring IPMClass )
{
	wstring FolderDN(GetDNFromFolderName(Name));
	CLdapObject objFolder( FolderDN.c_str() );
	objFolder.PutAttribute( L"ObjectClass", L"OlFolder" );
	objFolder.PutAttribute( L"cn", Name.c_str() );
	objFolder.PutAttribute( L"IPMClass", IPMClass.c_str() );
	CLdapError e = objFolder.Put( sess );
	if( !e.Success() ) {
		wstring msg(L"Failed to create folder on server : ");
		msg += e.Description();
		dbg << msg << L"\n";
	}
	dbg << L"Added LDAP folder " << Name << L"\n";
}

void CStore::DeleteLdapFolder( const wstring FolderName, CLdapSession& sess)
{
	CLdapError Err;
	wstring DN = GetDNFromFolderName(FolderName);

    // Must delete the entries in this folder before we can delete the folder itself
	CLdapSearch EntrySearch( sess );
	
	try
	{
		EntrySearch.Search(	DN.c_str()
						,LDAP_SCOPE_ONELEVEL
						,L"ObjectClass=OlEntry"
						,NULL );

		CLdapObject Entry;
		while( EntrySearch.GetNextObject( Entry ) )
		{
			Err = Entry.Delete( sess );
			if( ! Err.Success() )
			{
				throw CError( L"Failed while deleting folder " + FolderName);
				dbg << L"Failed to delete entry " << Entry.GetDN() << L"\n ERROR: " << Err.Description() << L"\n";
			}
		}
	
		CLdapObject Folder(DN.c_str());
		Err = Folder.Delete( sess );
		if( ! Err.Success() )
		{
			throw CError( L"Failed to delete folder " + FolderName);
			dbg << L"Failed to delete folder " << Folder.GetDN() << L"\n ERROR: " << Err.Description() << L"\n";
		}
	} catch (CLdapException& err )
	{
		dbg << L"Failed to delete folder " << err.Description() << L"\n";
		throw CError( L"Failed to delete folder");
	}
}

const wchar_t CStore::m_FolderNameDelimiter = L'/';
const wstring CStore::m_RFC2254ReservedChars = L"*()";

// Add escape sequences as per RFC2254
void CStore::AddEscapeSequences( wstring& Str, const wstring Reserved )
{
	wstring::size_type s;
	char buf[4];
	wstring ReservedChars(Reserved+L"\\"); // Always escape the escape char
	for( s=Str.find_first_of(ReservedChars);s!=string::npos;s=Str.find_first_of(ReservedChars,s+3))
	{
		USES_CONVERSION;
		const char * EscapedChar = W2CA(wstring(1,Str[s]).c_str());
		sprintf(buf,"\\%02X",(int)EscapedChar[0]);
		Str.erase(s,1);
		Str.insert(s,A2CW(buf));
	}
}

// Remove escape sequences as per RFC2254
void CStore::RemoveEscapeSequences( wstring& Str )
{
	wstring::size_type s;
	for( s=Str.find_first_of(L"\\");s!=string::npos;s=Str.find_first_of(L"\\",s+1))
	{
		USES_CONVERSION;
		string EscapedChar(1,(char)wcstoul(Str.substr(s+1,2).c_str(),NULL,16));
		Str = Str.substr(0,s)+wstring(A2CW(EscapedChar.c_str()))+Str.substr(s+3);
	}
}

void CStore::SplitFirst( wstring& Str, wstring& First, const wchar_t Delimiter )
{
    string::size_type s = Str.find_first_of(wstring(1,Delimiter));
    First = Str.substr(0,s);
    if(s==string::npos) Str=L"";
	else Str = Str.substr(s+1);
}

// Load a single record from an archive
void CStore::FolderRecord::Load( CStoreArchive* in, const int Length )
{
	CStringRecord StrRec( CStoreArchive::CStringRecord );
	in->Load(StrRec);
	m_IPMClass = StrRec.Get();
	in->Load(StrRec);
	m_FolderName = StrRec.Get();
}

// Save a single record in an archive
void CStore::FolderRecord::Save( CStoreArchive* out) const
{
	out->Save( CStringRecord( m_IPMClass, CStoreArchive::CStringRecord ) );
	out->Save( CStringRecord( m_FolderName, CStoreArchive::CStringRecord ) );
}

bool CStore::Save( CStoreArchive* out ) 
{
	bool rc = true;

	set<wstring>::iterator si;
	for( si = m_LocalFolders.begin(); si != m_LocalFolders.end() && rc; ++si)
		rc = out->Save( CStringRecord( *si, CStoreArchive::CStoreLocalFolder ) );

	for( si = m_LdapFolders.begin(); si != m_LdapFolders.end() && rc; ++si)
		rc = out->Save( CStringRecord( *si, CStoreArchive::CStoreLdapFolder ) );

	map<wstring,CFolder*>::iterator mi;
	for( mi = m_Folders.begin(); mi != m_Folders.end() && rc; ++mi)
	{
		CFolder* pFolder = mi->second;
		rc = out->Save( FolderRecord( pFolder->GetIPMClass()
									, mi->first
									, CStoreArchive::FolderRecord ) );
		dbg << L"Saved folder " << mi->first << L"\n";
		pFolder->Save( out );
	}

	return rc;
}

bool CStore::Load( CStoreArchive* in ) 
{
	
	CStringRecord LocalIndexRec( CStoreArchive::CStoreLocalFolder );
	in->ReWind();
	while(in->Load(LocalIndexRec))
		m_LocalFolders.insert(LocalIndexRec.Get());
	
	CStringRecord LdapIndexRec( CStoreArchive::CStoreLdapFolder );
	in->ReWind();
	while(in->Load(LdapIndexRec))
		m_LdapFolders.insert(LdapIndexRec.Get());

	FolderRecord rec(CStoreArchive::FolderRecord);
	in->ReWind();
	while(in->Load(rec))
		if( ! AddFolderSyncronizer( rec.m_IPMClass, rec.m_FolderName ) )
		{
			// Folder has disappeared since last time or its a public folder and rights have 
			// been revoked since last time. Either way we should forget it.
			m_LocalFolders.erase( rec.m_FolderName );
			m_LdapFolders.erase( rec.m_FolderName );
			dbg << L"Ignoring folder " << rec.m_FolderName << L"\n";
		}
		else
			dbg << L"Loaded folder " << rec.m_FolderName << L"\n";

	map<wstring,CFolder*>::iterator mi;
	for( mi = m_Folders.begin(); mi != m_Folders.end(); ++mi)
		mi->second->Load(in);

	return true;
}

