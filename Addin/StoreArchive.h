#pragma once

class CStoreArchive;
class IFileRecord
{
public:
	virtual LONG GetIdentifier() const = 0; 
	virtual void Load( CStoreArchive* in, const int Length ) = 0;
	virtual void Save( CStoreArchive* out ) const = 0;
};

// Standard string file record
class CStringRecord : public IFileRecord
{
	LONG m_Id;
	wstring m_Str;

public:
	CStringRecord( const LONG Identifier );
	CStringRecord( const wstring Str
				 , const LONG Identifier );

	virtual LONG GetIdentifier() const { return m_Id; }
	virtual void Load( CStoreArchive* in, const int Length );
	virtual void Save( CStoreArchive* out ) const;

	const wstring Get() const { return wstring(m_Str); }
};

class CStoreArchive
{
	struct Header
	{
		static const wchar_t * m_ArchiveTitle;
		static const USHORT m_ArchiveVersion;
		wchar_t m_Title[32];
		USHORT m_Version;
		LONG GetLength() const { 
			return sizeof(m_Title) + sizeof(m_Version); 
		}
		USHORT Load( fstream& in );
		void Save( fstream& out );
	} m_Header;

public:

	fstream fs;

	CStoreArchive( const char * FileName, const bool in = true );
	virtual ~CStoreArchive();

	static const LONG CStringRecord;
	static const LONG CStoreLocalFolder;
	static const LONG CStoreLdapFolder;
	static const LONG FolderRecord;
	static const LONG FolderLocalIndexRecord;
	static const LONG FolderLdapIndexRecord;
	static const LONG EntryRecord;

	bool HeaderOK();
	bool SaveHeader();
	void ReWind(); // Rewind to first record (after header)
	bool eof() const { return fs.eof(); }
	bool good() const { return fs.good(); }
	bool Load( IFileRecord& rp); 
	bool Save( const IFileRecord& rp);
};
