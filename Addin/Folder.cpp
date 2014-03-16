#include "stdafx.h"
using namespace std;

CFolder::CFolder( CStore const * Store
				, const wstring FolderName
				, const wstring FolderEntryID
				, const wstring IPMClass
				, const Outlook::OlDefaultFolders OlFolderType
				, const Outlook::OlItemType ItemType
				, const wstring BaseDN 
				, CComBSTR StoreID )
	:m_Store(Store)
	,m_FolderName(FolderName)
	,m_FolderEntryID(FolderEntryID)
	,m_IPMClass(IPMClass)
	,m_OlFolderType(OlFolderType)
	,m_OlItemType(ItemType)
	,m_BaseDN(BaseDN)
	,m_pNameSpace(NULL)
	,m_Folder(NULL)
	,m_StoreID(StoreID)
	,m_pMapiFolder(NULL)
{
}

CFolder::~CFolder()
{
	// Free up any map entries that we've allocated
	map<wstring,CEntry*>::iterator i;
	for( i = m_Entries.begin(); i != m_Entries.end(); ++i )
		delete i->second;

	if(m_Folder) m_Folder->Release();
	if(m_pMapiFolder) delete m_pMapiFolder;
}

// Build a set containing the indexes of every Outlook entry
// Compare this set to the set obtained from the last time this fn was called
// Return lists of entries that have been deleted, modified or added
void CFolder::ScanLocalEntries(  set<wstring>& Added
										  ,set<wstring>& Modified
										  ,set<wstring>& Deleted
										  ,set<wstring>& UnModified
										  ,CDisplay& dbg )
{
	set< wstring > LocalEntries;
	set< wstring > ModifiedLocalEntries;

	CComPtr<Outlook::_Items> pItems;
	CHECKHR(m_Folder->get_Items(&pItems));

	IDispatch * pDisp;
	for( pItems->GetFirst(&pDisp); pDisp; pItems->GetNext(&pDisp)) 
	{
		CEntry * pEntry = CEntry::Make( m_OlFolderType, m_Store->m_Addin->GetOutlookMajorVersion() );
		if( !pEntry->CheckOlPtr(pDisp) )
		{
			// entry probably an unexpected type
			delete pEntry;
			continue; 
		}
		
		pEntry->SetIndex( pDisp);
		pEntry->SetTimeStamp( pDisp);
		pEntry->SetEntryName( pDisp);

		// Add index into local entries set
		LocalEntries.insert(pEntry->GetIndex());

		// Add or update the entry in our map
		map< wstring, CEntry* >::const_iterator iterExistingEntry = m_Entries.find( pEntry->GetIndex() );
		if( iterExistingEntry == m_Entries.end() )
		{
			// Add new entry
			pEntry->SetNewTimeStamp( pEntry->GetTimeStamp() );
			m_Entries[ pEntry->GetIndex() ] = pEntry;
			dbg << L"Adding local entry " << pEntry->GetIndex() << L"\n";
		}
		else
		{
			// Check if existing entry has been modified
			CEntry * pExistingEntry = iterExistingEntry->second;

			pExistingEntry->SetEntryNameStr( pEntry->GetEntryShortName().c_str() );

			if( pExistingEntry->GetTimeStamp() != pEntry->GetTimeStamp() )
				ModifiedLocalEntries.insert(pEntry->GetIndex());

			pExistingEntry->SetNewTimeStamp( pEntry->GetTimeStamp() );

			dbg << L"Updating local entry " << pEntry->GetIndex() << L"\n";
			delete pEntry;
		}
	}

	// Entries that existed last time but don't exist now should be deleted
	set_difference(	  m_LocalEntries.begin()
					, m_LocalEntries.end()
					, LocalEntries.begin()
					, LocalEntries.end()
					, inserter(Deleted, Deleted.begin()) );

	// Entries that didn't exist last time and exist now should be added
	set_difference(   LocalEntries.begin()
					, LocalEntries.end()
					, m_LocalEntries.begin()
					, m_LocalEntries.end()
					, inserter(Added, Added.begin()) );

	// Entries that existed last time and still exist this time
	set<wstring> ContinueToExist;
	set_intersection( LocalEntries.begin()
					, LocalEntries.end()
					, m_LocalEntries.begin()
					, m_LocalEntries.end()
					, inserter(ContinueToExist, ContinueToExist.begin()) );

	// Entries that continue to exist but have an updated timestamp since last time
	// have been modified
	set_intersection( ContinueToExist.begin()
					, ContinueToExist.end()
					, ModifiedLocalEntries.begin()
					, ModifiedLocalEntries.end()
					, inserter(Modified, Modified.begin()) );

	// Entries that continue to exist but haven't been modified are unmodified
	set_difference( ContinueToExist.begin()
				  , ContinueToExist.end()
				  , ModifiedLocalEntries.begin()
  				  , ModifiedLocalEntries.end()
				  , inserter(UnModified, UnModified.begin()) );
}

// Build a set containing the indexes of every LDAP entry
// Compare this set to the set obtained from the last time this fn was called
// Return lists of entries that have been deleted, modified or added
void CFolder::ScanLdapEntries( CLdapSession& sess 
										,set<wstring>& Added
										,set<wstring>& Modified
										,set<wstring>& Deleted
										,set<wstring>& UnModified
										,CDisplay& dbg )
{
	set< wstring > LdapEntries;
	set< wstring > ModifiedLdapEntries;
    list<CEntry *> NewEntries;

	CLdapSearch objSearch( sess );
	LPCOLESTR apszUserAttrsSD[] = { L"cn", L"modifyTimestamp", NULL };
	
	try
	{
		objSearch.Search(	m_BaseDN.c_str()
						,LDAP_SCOPE_ONELEVEL
						,L"objectClass=*"
						,apszUserAttrsSD );

		CLdapObject objLdapEntry;
		while( objSearch.GetNextObject( objLdapEntry ) )
		{
			CEntry * pEntry = CEntry::Make( m_OlFolderType, m_Store->m_Addin->GetOutlookMajorVersion() );

			CLdapAttribute objLdapAttr;
			CLdapValue objValue;

			wstring dn(objLdapEntry.GetDN());
			wstring szIndex(dn,0,dn.find(L','));

			// LDAP returns cn which we can use as the entry's CommonName
			CHECKTRUE( objLdapEntry.GetAttribute( L"cn", objLdapAttr ) );
			CHECKTRUE(objLdapAttr.ValueCount() == 1 );
			objValue = objLdapAttr(); // assumes 1 value
			wstring szCn((LPCOLESTR)objValue);

			// LDAP returns modifyTimestamp which we can use as the entry's LdapTimeStamp
			CHECKTRUE( objLdapEntry.GetAttribute( L"modifyTimestamp", objLdapAttr ) );
			CHECKTRUE(objLdapAttr.ValueCount() == 1 );
			objValue = objLdapAttr(); // assumes 1 value
			LPCOLESTR szTimeStamp = (LPCOLESTR)objValue;
			
			// Set this entry up
			pEntry->SetIndexStr( szIndex.c_str() );
			pEntry->SetEntryNameStr( szCn.c_str() );
			pEntry->SetLdapTimeStamp( CEntry::LdapTimeToGMTVariantTime(wstring(szTimeStamp)) );
			NewEntries.push_back(pEntry);
		}
	} catch ( CLdapException& )
	{
		// Something went wrong. We may only have a partial index
		for( list<CEntry*>::const_iterator i = NewEntries.begin(); i != NewEntries.end(); ++i )
			delete *i;
		throw CError( L"Failed read complete index" );
	}

	// LDAP search completed normally so this is a complete index
	for( list<CEntry*>::iterator i = NewEntries.begin(); i != NewEntries.end(); ++i )
	{
		CEntry * pEntry = *i;

		LdapEntries.insert( pEntry->GetIndex() );

		map< wstring, CEntry* >::const_iterator iterExistingEntry = m_Entries.find( pEntry->GetIndex() );
		if( iterExistingEntry == m_Entries.end() )
		{
			// New entry
			pEntry->SetNewLdapTimeStamp( pEntry->GetLdapTimeStamp() );
			m_Entries[ pEntry->GetIndex() ] = pEntry;
			dbg << L"Adding LDAP entry " << pEntry->GetIndex() << L"\n";
		
			// Note that OlIndex need not be present since there is no Outlook entry.
		}
		else
		{
			CEntry * pExistingEntry = iterExistingEntry->second;

			pExistingEntry->SetEntryNameStr( pEntry->GetEntryShortName().c_str() );

			if( pExistingEntry->GetLdapTimeStamp() != pEntry->GetLdapTimeStamp() ) 
				ModifiedLdapEntries.insert(pEntry->GetIndex());

			// Note that OlIndex will have already been setup by localScan

			pExistingEntry->SetNewLdapTimeStamp( pEntry->GetLdapTimeStamp() );
			dbg << L"Updating LDAP entry " << pEntry->GetIndex() << L"\n";
			delete pEntry;
		}
	}

	// Entries that existed last time but don't exist now should be deleted
	set_difference(	  m_LdapEntries.begin()
					, m_LdapEntries.end()
					, LdapEntries.begin()
					, LdapEntries.end()
					, inserter(Deleted, Deleted.begin()) );

	// Entries that didn't exist last time and exist now should be added
	set_difference(   LdapEntries.begin()
					, LdapEntries.end()
					, m_LdapEntries.begin()
					, m_LdapEntries.end()
					, inserter(Added, Added.begin()) );

	// Entries that existed last time and still exist this time
	set<wstring> ContinueToExist;
	set_intersection( LdapEntries.begin()
					, LdapEntries.end()
					, m_LdapEntries.begin()
					, m_LdapEntries.end()
					, inserter(ContinueToExist, ContinueToExist.begin()) );

	// Entries that continue to exist but have an updated timestamp since last time
	// have been modified
	set_intersection( ContinueToExist.begin()
					, ContinueToExist.end()
					, ModifiedLdapEntries.begin()
					, ModifiedLdapEntries.end()
					, inserter(Modified, Modified.begin()) );

	// Entries that continue to exist but haven't been modified are unmodified
	set_difference( ContinueToExist.begin()
				  , ContinueToExist.end()
				  , ModifiedLdapEntries.begin()
  				  , ModifiedLdapEntries.end()
				  , inserter(UnModified, UnModified.begin()) );
}

class EntryRuleFunctor: public CMatrix::IRule
{
	CLdapSession& m_sess;
	CDisplay& stat;
	CDisplay& dbg;
	CFolder* m_Folder;
	bool* m_pNoErrors;
	bool m_WriteAllowed;
public:
	EntryRuleFunctor( CLdapSession& sess, CDisplay& Stat, CDisplay& Dbg, CFolder* Folder, bool* pNoErrors, bool WriteAllowed ) 
		:m_sess(sess), stat(Stat), dbg(Dbg), m_Folder(Folder), m_pNoErrors(pNoErrors), m_WriteAllowed(WriteAllowed) {}
	virtual ~EntryRuleFunctor() {}
	virtual void Do(  const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring EntryIndex )
	{
		m_Folder->ApplyRule(Rule, LocalOp, LdapOp, EntryIndex, m_sess, stat, dbg, m_pNoErrors, m_WriteAllowed);
	}
};

void CFolder::ApplyRule( const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring EntryIndex, CLdapSession& sess, CDisplay& stat, CDisplay& dbg, bool* pNoErrors, bool WriteAllowed )
{
	CEntry* pEntry = m_Entries[EntryIndex];
	
	CANCELLATION_POINT

	try {

		if( Rule != CMatrix::RTN )
		{
			dbg << L"Entry : " << pEntry->GetEntryName() << L"\n";
			dbg << L"-LDAP(" << CMatrix::ChangeTypeStr[LdapOp] << L") LOCAL(" << CMatrix::ChangeTypeStr[LocalOp] << L")\n";
		}
		
		switch( Rule ) 
		{
		case CMatrix::RTCOC:
			CopyToLdap( sess, pEntry, true ); 
			break;
		case CMatrix::RTCO:
			CopyToLdap( sess, pEntry, false ); 
			break;
		case CMatrix::RTDL:
			if( WriteAllowed ) DeleteLdap( sess, pEntry );
			break;
		case CMatrix::RTDO: 
			DeleteLocal( sess, pEntry ); 
			break;
		case CMatrix::RTCLC: 
			CopyFromLdap( sess, pEntry, true ); 
			break;
		case CMatrix::RTCL:
			CopyFromLdap( sess, pEntry, false ); 
			break;
		case CMatrix::RTCMR:
			dbg << L"-Outlook timestamp " << CEntry::GMTVariantTimeToStr(pEntry->GetNewTimeStamp()) << L"\n"
				<< L"-Server timestamp " << CEntry::GMTVariantTimeToStr(pEntry->GetNewLdapTimeStamp()) << L"\n";
			if( pEntry->CompareTimeStamps() < 0 ) CopyToLdap( sess, pEntry, false ); 
			else CopyFromLdap( sess, pEntry, false );
			break;
		default:
			break;
		}
		CMatrix::CorrectIndex( Rule, EntryIndex, LdapOp, LocalOp, m_LdapEntries, m_LocalEntries);
		if( CMatrix::WasModified( LdapOp, LocalOp ) ) pEntry->CommitNewTimeStamps();

	} catch ( CError& e ) {
		dbg << L"ENTRY SYNCHRONIZATION ERROR " << e.GetDescription() << L"\n";
		stat << L"\nERROR: " << pEntry->GetEntryName() << L" failed to " << CMatrix::RuleStr[Rule];
		*pNoErrors = false;
		return;
	} catch ( CLdapException& e ) {
		dbg << L"LDAP ENTRY SYNCHRONIZATION ERROR " << e.Description() << L"\n";
		stat << L"\nERROR: " << pEntry->GetEntryName() << L" failed to " << CMatrix::RuleStr[Rule];
		*pNoErrors = false;
		return;
	}
	if( Rule != CMatrix::RTN ) dbg << L"Action : " << CMatrix::RuleStr[Rule] << L"\n";
}

// Syncronize local and LDAP entries
bool CFolder::Syncronize( CLdapSession& sess, CDisplay& stat, CDisplay& dbg, Outlook::_NameSpace* NameSpace, CMAPIStore* pMapiStore, const bool WriteAllowed )
{
	m_pNameSpace = NameSpace;

	// Get the Outlook MAPIFolder from the FolderEntryID
	if(m_Folder) m_Folder->Release();
	CHECKHR(m_pNameSpace->GetFolderFromID( CComBSTR(m_FolderEntryID.c_str())
										, CComVariant( m_StoreID)
										, &m_Folder ) );
	CHECKNOTNULL(m_Folder);

	// Get the CMAPIFolder 
	CMAPIEntryID MAPIFolderEntryID;
	MAPIFolderEntryID.Set( m_FolderEntryID );

	if(m_pMapiFolder) delete m_pMapiFolder;
	m_pMapiFolder = new CMAPIFolder(pMapiStore);
	CHECKNOTNULL(m_pMapiFolder);

	CHECKHR( m_pMapiFolder->OpenFolder( MAPIFolderEntryID ));

	bool NoErrors = true;

	set<wstring> LocalAdded, LocalDeleted, LocalUnModified, LocalModified;
	ScanLocalEntries( LocalAdded
					, LocalModified
					, LocalDeleted
					, LocalUnModified
					, dbg );
	
	set<wstring> LdapAdded, LdapDeleted, LdapModified, LdapUnModified;
	ScanLdapEntries(  sess
					, LdapAdded
					, LdapModified
					, LdapDeleted
					, LdapUnModified
					, dbg );

	CMatrix Matrix;
	Matrix.Intersection(  &LocalAdded, &LocalModified, &LocalDeleted, &LocalUnModified
						, &LdapAdded, &LdapModified, &LdapDeleted, &LdapUnModified
						, EntryRuleFunctor(sess,stat,dbg,this,&NoErrors,WriteAllowed) );

	m_Folder->Release();
	m_Folder = NULL;

	delete m_pMapiFolder;
	m_pMapiFolder = NULL;

	return NoErrors;
}

void CFolder::CopyToLdap( CLdapSession& sess, CEntry* pEntry, const bool Create )
{
	// Outlook index must be set by ScanLocal
	CHECKTRUE( pEntry->GetOlIndex() != L"" );

	IDispatch * pDisp = NULL;
	if( FAILED( m_pNameSpace->GetItemFromID( CComBSTR(pEntry->GetOlIndex().c_str())
						, CComVariant( m_StoreID), &pDisp) ))
		throw CError( L"Failed attempt to find local copy of " + pEntry->GetEntryName() );
	CHECKNOTNULL( pDisp);

	pEntry->Pack( pDisp, m_pMapiFolder );
	pEntry->Send( sess, m_BaseDN.c_str(), Create );

//	pEntry->DebugBox();

	// When LDAP saves an entry it sets the entry's modifyTimestamp
	// There is clearly a variable delay between us sending the entry to LDAP
	// and LDAP timestamping the entry. This delay is in the region of 1 -> 10 secs
	// We must read the timestamp that LDAP has given the entry back from the server
	// and update our records.
	CLdapSearch objSearch( sess );
	LPCOLESTR apszUserAttrsSD[] = { L"modifyTimestamp", NULL };
	wstring szContactDN = pEntry->GetIndex() + L"," + m_BaseDN;
	CLdapObject objLdapEntry;

	try 
	{
		objSearch.Search( szContactDN.c_str(), LDAP_SCOPE_BASE, L"objectClass=*", apszUserAttrsSD );
		if( !objSearch.GetNextObject( objLdapEntry ) )
			throw CError( L"Failed attempt to read server timestamp for " + pEntry->GetEntryName() );
	} catch (CLdapException&  )
	{
		throw CError( L"Failed attempt to read server timestamp for " + pEntry->GetEntryName() );
	}

	CLdapAttribute objLdapAttr;
	CLdapValue objValue;
	CHECKTRUE( objLdapEntry.GetAttribute( L"modifyTimestamp", objLdapAttr ) );
	CHECKTRUE(objLdapAttr.ValueCount() == 1 );
	objValue = objLdapAttr(); // assumes 1 value
	LPCOLESTR szTimeStamp = (LPCOLESTR)objValue;

	// Update our records of the server timestamp
	pEntry->SetLdapTimeStamp( CEntry::LdapTimeToGMTVariantTime( wstring(szTimeStamp) ) );
	pEntry->SetNewLdapTimeStamp( pEntry->GetLdapTimeStamp() );
}

void CFolder::CopyFromLdap( CLdapSession& sess, CEntry* pEntry, const bool Create )
{	
	IDispatch * pDisp = NULL;

	if( Create ) 
	{
		// Add new Outlook entry
		CComPtr<Outlook::_Items> pItems;
		CHECKHR(m_Folder->get_Items(&pItems));
		if( FAILED(pItems->Add( CComVariant( pEntry->GetEntryClass()), &pDisp))) 
			throw CError( L"Failed attempt to add local copy of " + pEntry->GetEntryName() );
	}
	else
	{	
		// Outlook index must be set by ScanLocal
		CHECKTRUE( pEntry->GetOlIndex() != L"" );

		// Find existing Outlook entry so we can modify it
		if( FAILED( m_pNameSpace->GetItemFromID( CComBSTR( pEntry->GetOlIndex().c_str())
						, CComVariant( m_StoreID), &pDisp) ))
			throw CError( L"Failed attempt to find local copy of " + pEntry->GetEntryName() );
	}

	CHECKNOTNULL( pDisp);

	pEntry->Recv( sess, m_BaseDN.c_str() );
	if( pEntry->Validate() ) {

		pEntry->UnPack( pDisp, Create );

	//	pEntry->DebugBox();

		// When Outlook saves an entry it sets the entry's lastModified timestamp
		// Outlook is very quick to save entries but the operation isn't instananeous.
		// We must read the timestamp that Outlook has given the entry back and update
		// our records.
		pEntry->SetTimeStamp(pDisp);
		pEntry->SetNewTimeStamp( pEntry->GetTimeStamp() );
	}
}

// Delete an LDAP entry (which may or may not be there)
bool CFolder::DeleteLdap( CLdapSession& sess, CEntry* pEntry)
{

	CLdapSearch objSearch( sess );
	wstring szContactDN = pEntry->GetIndex() + L"," + m_BaseDN;
	CLdapObject objLdapEntry;

	try
	{
		objSearch.Search( szContactDN.c_str(), LDAP_SCOPE_BASE );

		if( !objSearch.GetNextObject( objLdapEntry ) )
			return false; // Nothing had to be deleted
	
		objLdapEntry.Delete( sess );
	}
	catch (CLdapException&  )
	{
		throw CError( L"Failed to delete entry");
	}
	return true; // We deleted the entry
}

// Delete local entry (which may or may not be there)
bool CFolder::DeleteLocal( CLdapSession& sess, CEntry* pEntry)
{

	IDispatch * pDisp = NULL;

	// Outlook index must be set by ScanLocal
	CHECKTRUE( pEntry->GetOlIndex() != L"" );

	if( FAILED( m_pNameSpace->GetItemFromID( CComBSTR( pEntry->GetOlIndex().c_str())
						, CComVariant( m_StoreID), &pDisp) ))
		return false; // Nothing had to be deleted

	CHECKNOTNULL( pDisp);

	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);
	CHECKHR(spContact->Delete());

	return true; // We deleted the entry
}

void CFolder::InvalidateIndexes()
{
	m_LocalEntries.erase( m_LocalEntries.begin(), m_LocalEntries.end() );
	m_LdapEntries.erase( m_LdapEntries.begin(), m_LdapEntries.end() );
}

// Load a single record from an archive
void CFolder::EntryRecord::Load( CStoreArchive* in, const int Length )
{
	CStringRecord StrRec( CStoreArchive::CStringRecord );
	in->fs.read( (char *)&m_ItemType, sizeof(m_ItemType));
	in->Load(StrRec);
	m_Index = StrRec.Get();
	in->Load(StrRec);
	m_OlIndex = StrRec.Get();
	in->Load(StrRec);
	m_Name = StrRec.Get();
	in->fs.read( (char *)&m_OlDate, sizeof(m_OlDate));
	in->fs.read( (char *)&m_LdapDate, sizeof(m_LdapDate));
	in->Load(StrRec);
	m_FolderName = StrRec.Get();
}

// Save a single record in an archive
void CFolder::EntryRecord::Save( CStoreArchive* out) const
{
	out->fs.write( (char *)&m_ItemType, sizeof(m_ItemType));
	out->Save( CStringRecord( m_Index, CStoreArchive::CStringRecord ) );
	out->Save( CStringRecord( m_OlIndex, CStoreArchive::CStringRecord ) );
	out->Save( CStringRecord( m_Name, CStoreArchive::CStringRecord ) );
	out->fs.write( (char *)&m_OlDate, sizeof(m_OlDate));
	out->fs.write( (char *)&m_LdapDate, sizeof(m_LdapDate));
	out->Save( CStringRecord( m_FolderName, CStoreArchive::CStringRecord ) );
}

void CFolder::FolderIndexRecord::Load( CStoreArchive* in, const int Length )
{
	CStringRecord StrRec( CStoreArchive::CStringRecord );
	in->Load(StrRec);
	m_Index = StrRec.Get();
	in->Load(StrRec);
	m_FolderName = StrRec.Get();
}

void CFolder::FolderIndexRecord::Save( CStoreArchive* out ) const
{
	out->Save( CStringRecord( m_Index, CStoreArchive::CStringRecord ) );
	out->Save( CStringRecord( m_FolderName, CStoreArchive::CStringRecord ) );
}

bool CFolder::Save( CStoreArchive* out)
{
	bool rc = true;

	set<wstring>::const_iterator si;
	for( si = m_LocalEntries.begin(); si != m_LocalEntries.end() && rc ; ++si )
		rc = out->Save( FolderIndexRecord( *si, m_FolderName, CStoreArchive::FolderLocalIndexRecord ) );

	for( si = m_LdapEntries.begin(); si != m_LdapEntries.end() && rc ; ++si )
		rc = out->Save( FolderIndexRecord( *si, m_FolderName, CStoreArchive::FolderLdapIndexRecord ) );

	map< wstring, CEntry* >::const_iterator mi;
	for( mi = m_Entries.begin(); mi != m_Entries.end() && rc ; ++mi )
	{
		CEntry * pEntry = mi->second;
		rc = out->Save( EntryRecord( m_OlItemType
									, mi->first
									, pEntry->GetOlIndex()
									, pEntry->GetEntryShortName()
									, pEntry->GetTimeStamp()
									, pEntry->GetLdapTimeStamp()
									, m_FolderName
									, CStoreArchive::EntryRecord ) );
	}

	return rc;
}

bool CFolder::Load( CStoreArchive* in ) 
{

	FolderIndexRecord LocalIndexRec(CStoreArchive::FolderLocalIndexRecord);
	in->ReWind();
	while(in->Load(LocalIndexRec))
		if(LocalIndexRec.m_FolderName == m_FolderName)
			m_LocalEntries.insert(LocalIndexRec.m_Index);
	
	FolderIndexRecord LdapIndexRec( CStoreArchive::FolderLdapIndexRecord );
	in->ReWind();
	while(in->Load(LdapIndexRec))
		if(LdapIndexRec.m_FolderName == m_FolderName)
			m_LdapEntries.insert(LdapIndexRec.m_Index);

	EntryRecord rec(CStoreArchive::EntryRecord);
	in->ReWind();
	while( in->Load(rec) )
	{
		if( rec.m_FolderName != m_FolderName ) continue;

		map< wstring, CEntry* >::const_iterator mi = m_Entries.find( rec.m_Index );
		if( mi == m_Entries.end() )
		{
			CEntry * pEntry = CEntry::Make( m_OlFolderType, m_Store->m_Addin->GetOutlookMajorVersion() );
			pEntry->SetIndexStr( rec.m_Index.c_str());
			pEntry->SetTimeStampDate( rec.m_OlDate );
			pEntry->SetLdapTimeStamp( rec.m_LdapDate );
			pEntry->SetEntryNameStr( rec.m_Name.c_str() );
			pEntry->SetOlIndexStr( rec.m_OlIndex.c_str());
			m_Entries[rec.m_Index] = pEntry;
		}
	}

	return true;
}
