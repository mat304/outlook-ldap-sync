#ifndef __STORE_H_
#define __STORE_H_

class CStore : public CCPLWinThread
{

public:
	CAddin * m_Addin;

private:
	const wstring m_StoreName;
	const wstring m_PublicSubFolder;
	CComBSTR m_StoreID;
	CMAPISession * m_pMapiSession;
	CMAPIStore * m_pMapiStore;
	CCfg * m_Cfg;
	bool m_FirstSync;
	Outlook::_NameSpace * m_pNameSpace;
	Outlook::_Folders * m_pStoreFolders;
	CDisplay dbg; // Normally goes to debug log
	CDisplay stat; // Normally gets displayed in a SyncWindow
	
	set<wstring> m_LocalFolders;
	set<wstring> m_LdapFolders;
	map<wstring,CFolder*> m_Folders;

	list<wstring> m_GroupDNs;
	map<wstring,wstring> m_GroupMap;

	static const PFA_READ_ACCESS = 0x0001; // presently includes modify & creates
	static const PFA_WRITE_ACCESS = 0x0002; // presently only adds delete
	map<wstring,DWORD> m_PublicFolderAccessMap;

	void RefreshGroups( CLdapSession& sess );

	void ScanLdapFolders(  CLdapSession& sess
						  ,const wstring IPMClass
						  ,set<wstring>& Folders );

	void ScanLocalFolders( const wstring IPMClass
						  ,Outlook::_Folders * pFolders
						  ,set<wstring>& Folders
						  ,const wstring ParentFolderName = L"" );

	void AddLocalFolder( const wstring FullName, wstring Name, const wstring IPMClass, Outlook::_Folders * pFolders );
	void DeleteLocalFolder( const wstring FullName );
	void AddLdapFolder( const wstring Name, CLdapSession& sess, const wstring IPMClass );
	void DeleteLdapFolder( const wstring FolderName, CLdapSession& sess);

	CFolder* AddFolderSyncronizer( const wstring IPMClass, const wstring FolderName );
	void DeleteFolderSyncronizer( const wstring FolderName );

	void Syncronize( const wstring IPMClass, CLdapSession& sess );

	class FolderRecord : public IFileRecord
	{
	public:
		const LONG m_Id;
		wstring m_IPMClass;
		wstring m_FolderName;

		FolderRecord( const LONG Identifier ) 
			:m_Id(Identifier)
			{}
		FolderRecord( const wstring IPMClass
					, const wstring FolderName
					, const LONG Identifier )
			:m_Id(Identifier)
			,m_IPMClass(IPMClass)
			,m_FolderName(FolderName)
			{}

		virtual LONG GetIdentifier() const { return m_Id; }
		virtual void Load( CStoreArchive* in, const int Length );
		virtual void Save( CStoreArchive* out ) const;
	};

	static const wchar_t m_FolderNameDelimiter;
	static const wstring m_RFC2254ReservedChars;
	static void AddEscapeSequences( wstring& Str, const wstring ReservedChars );
	static void RemoveEscapeSequences( wstring& Str );
	static void SplitFirst( wstring& Str, wstring& First, const wchar_t Delimiter = m_FolderNameDelimiter );
	const wstring GetDNFromFolderName(const wstring Name);
	const wstring GetEntryIDFromFolderName(const wstring Name);
	Outlook::OlDefaultFolders GetOlDefaultFolders( const wstring IPMClass);
	Outlook::OlItemType GetOlItemType( const wstring IPMClass);
	const bool IsPublicFolder( const wstring Name );
	const bool HavePublicFolderWriteAccess( const wstring Name );
	void SetFolderAccessRights( const wstring EntryID, const bool Read, const bool Modify, const bool Delete );

	void PromptUserForPassword();

public:

	CStore(CAddin * Addin, const wstring StoreName, const wstring PublicSubFolder, CCfg* cfg );
	virtual ~CStore(void);

	void ReadConfig();
	
	// When > 0 we terminate all syncronization threads asap
	static CCPLWinAtomicCounter TerminateAllSyncThreads; 

	void Proc();

	void ApplyRule( const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring FolderName, CLdapSession& sess, const wstring IPMClass );
	
	bool Save( CStoreArchive* out );
	bool Load( CStoreArchive* in );

	// Register textbox or logger here to display/log debug messages
	void RegisterDebugDisplay( IDisplay* out) { dbg.RegisterDisplay(out); }
	void UnRegisterDebugDisplay() { dbg.UnRegisterDisplay(); }

	// Register a status display here to receive the user status messages
	void RegisterStatusDisplay( IDisplay* out) { stat.RegisterDisplay(out); }
	void UnRegisterStatusDisplay() { stat.UnRegisterDisplay(); }
};

#define CANCELLATION_POINT { if(CStore::TerminateAllSyncThreads) return; }

#endif
