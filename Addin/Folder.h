#if !defined(AFX_FOLDER_H__C33205D1_266D_44F9_8B1A_34F8E952417C__INCLUDED_)
#define AFX_FOLDER_H__C33205D1_266D_44F9_8B1A_34F8E952417C__INCLUDED_

class CFolder
{
	CStore const * m_Store;
	const wstring m_FolderName;
	const wstring m_FolderEntryID;
	const wstring m_IPMClass;
	const Outlook::OlDefaultFolders m_OlFolderType;
	const Outlook::OlItemType m_OlItemType;
	const wstring m_BaseDN;
	Outlook::_NameSpace* m_pNameSpace;
	Outlook::MAPIFolder* m_Folder;
	CComBSTR m_StoreID;
	CMAPIFolder * m_pMapiFolder;

	set< wstring > m_LocalEntries;
	set< wstring > m_LdapEntries;
	map< wstring, CEntry* > m_Entries;

	void ScanLdapEntries(    CLdapSession& sess
							,set<wstring>& Added
							,set<wstring>& Modified
							,set<wstring>& Deleted
							,set<wstring>& UnModified
							,CDisplay& dbg );

	void ScanLocalEntries( set<wstring>& Added
						  ,set<wstring>& Modified
						  ,set<wstring>& Deleted
						  ,set<wstring>& UnModified
						  ,CDisplay& dbg );

	void CopyFromLdap( CLdapSession& sess, CEntry* pEntry, const bool Create );
	void CopyToLdap( CLdapSession& sess, CEntry* pEntry, const bool Create );
	bool DeleteLdap( CLdapSession& sess, CEntry* pEntry);
	bool DeleteLocal( CLdapSession& sess, CEntry* pEntry);

	class EntryRecord : public IFileRecord
	{
	public:
		const LONG m_Id;
		ULONG m_ItemType;
		wstring m_Index;
		wstring m_OlIndex;
		wstring m_Name;
		DATE m_OlDate;
		DATE m_LdapDate;
		wstring m_FolderName;

		EntryRecord( const LONG Identifier )
			:m_Id(Identifier)
			{}
		EntryRecord( const Outlook::OlItemType ItemType
					, const wstring Index
					, const wstring OlIndex
					, const wstring Name
					, const DATE OlDate
					, const DATE LdapDate
					, const wstring FolderName
					, const LONG Identifier )
			:m_Id(Identifier)
			,m_ItemType(ItemType)
			,m_Index(Index)
			,m_OlIndex(OlIndex)
			,m_Name(Name)
			,m_OlDate(OlDate)
			,m_LdapDate(LdapDate)
			,m_FolderName(FolderName)
			{}

		virtual LONG GetIdentifier() const { return m_Id; }
		virtual void Load( CStoreArchive* in, const int Length );
		virtual void Save( CStoreArchive* out ) const;
	};

	class FolderIndexRecord : public IFileRecord
	{
	public:
		const LONG m_Id;
		wstring m_Index;
		wstring m_FolderName;

		FolderIndexRecord( const LONG Identifier )
			:m_Id(Identifier)
			{}
		FolderIndexRecord(  const wstring Index
					, const wstring FolderName
					, const LONG Identifier )
			:m_Id(Identifier)
			,m_Index(Index)
			,m_FolderName(FolderName)
			{}

		virtual LONG GetIdentifier() const { return m_Id; }
		virtual void Load( CStoreArchive* in, const int Length );
		virtual void Save( CStoreArchive* out ) const;
	};

public:

	CFolder(  CStore const * Store
		    , const wstring FolderName
			, const wstring FolderEntryID
			, const wstring IPMClass
			, const Outlook::OlDefaultFolders OlFolderType
			, const Outlook::OlItemType ItemType
			, const wstring BaseDN 
			, CComBSTR StoreID );
	virtual ~CFolder();

	bool Syncronize( CLdapSession& sess, CDisplay& stat, CDisplay& dbg, Outlook::_NameSpace* NameSpace, CMAPIStore* pMapiStore, const bool WriteAllowed );
	void ApplyRule( const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring EntryIndex, CLdapSession& sess, CDisplay& stat, CDisplay& dbg, bool* pNoErrors, bool WriteAllowed );

	// Call this to erase deletion tracking information
	// This is necessary if you modify ServerName or BaseDN so that entries
	// will be stored in a different place. Note that this is handled automatically
	// when you use the SetBaseDN() & SetServerName() fns
	void InvalidateIndexes();

	// Get misc folder details for CStore
	const wstring GetEntryID() const { return m_FolderEntryID; }
	const wstring GetIPMClass() const { return m_IPMClass; }
	const wstring GetBaseDN() const { return m_BaseDN; }

	// Must save our sets of indexes so that they can be restored after Outlook shutdown
	// This is necessary in order to detect deleted Outlook & LDAP items
	bool Save( CStoreArchive* out );
	bool Load( CStoreArchive* in );
};

#endif // !defined(AFX_FOLDER_H__C33205D1_266D_44F9_8B1A_34F8E952417C__INCLUDED_)
