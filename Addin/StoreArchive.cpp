#include "stdafx.h"
using namespace std;

// Identifiers for different types of file record
const LONG CStoreArchive::CStringRecord = 1;
const LONG CStoreArchive::CStoreLocalFolder = 2;
const LONG CStoreArchive::CStoreLdapFolder = 3;
const LONG CStoreArchive::FolderRecord = 4;
const LONG CStoreArchive::FolderLocalIndexRecord = 5;
const LONG CStoreArchive::FolderLdapIndexRecord = 6;
const LONG CStoreArchive::EntryRecord = 7;

// These are the file format identifiers that this version of the software writes
// although older versions may be able to be read
const wchar_t * CStoreArchive::Header::m_ArchiveTitle = L"GroupWare";
const USHORT CStoreArchive::Header::m_ArchiveVersion = (USHORT)3;

// Load an archive header and return archive file format version
// Returns either the file format version number or 0 if it doesn't think
// that this file is for us.
USHORT CStoreArchive::Header::Load( fstream& in)
{
	in.read( (char *)&m_Title, sizeof(m_Title) );
	in.read( (char *)&m_Version, sizeof(m_Version) );
	if( wstring(m_ArchiveTitle) != wstring(m_Title) )
		return (USHORT)0; // File titles don't match
	return m_Version;
}

void CStoreArchive::Header::Save( fstream& out)
{
	lstrcpy( m_Title, m_ArchiveTitle );
	m_Version = m_ArchiveVersion;
	out.write( (char *)&m_Title, sizeof(m_Title) );
	out.write( (char *)&m_Version, sizeof(m_Version) );
}

CStoreArchive::CStoreArchive(  const char * FileName, const bool in )
{
	// File won't open in rw mode !
	if( in ) fs.open( FileName, (ios_base::binary|ios_base::in) );
	else fs.open( FileName, (ios_base::binary|ios_base::out|ios_base::trunc) );
}

CStoreArchive::~CStoreArchive()
{
	fs.close();
}

bool CStoreArchive::HeaderOK()
{
	if( !fs.is_open() ) return false;
	return (m_Header.Load(fs) == m_Header.m_ArchiveVersion);
}

bool CStoreArchive::SaveHeader()
{
	if( !fs.is_open() ) return false;
	m_Header.Save(fs);
	return fs.good();
}

void CStoreArchive::ReWind()
{
	fs.clear(); // Fuckers at MS - see http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B146445
	fs.seekg( m_Header.GetLength(), ios_base::beg );
}

// Loads the next standard record of a given type from an archive
bool CStoreArchive::Load( IFileRecord& rp)
{
	LONG Id = 0, Len = 0; 
	LONG mid = rp.GetIdentifier();
	while( (Id != rp.GetIdentifier()) && fs.good() && !fs.eof() )
	{
		fs.seekg( Len, ios_base::cur );
		fs.read( (char *)&Id, sizeof(LONG));
		fs.read( (char *)&Len, sizeof(LONG));
	}
	if( Id == rp.GetIdentifier() )
		rp.Load(this,Len);
	return (fs.good()&&!fs.eof());
}

// Saves a standard record
bool CStoreArchive::Save( const IFileRecord& rp )
{
	LONG Id = rp.GetIdentifier();
	LONG DataLen = 0;
	fs.write( (char *)&Id, sizeof(LONG));
	fstream::pos_type DataLenPos = fs.tellp();
	fs.write( (char *)&DataLen, sizeof(LONG)); // write dummy length
	fstream::pos_type DataStartPos = fs.tellp();
	rp.Save(this);
	fstream::pos_type DataEndPos = fs.tellp();

	// Now we know the record data length go back and write it into the header
	fs.clear();
	fs.seekp(DataLenPos);
	DataLen = DataEndPos - DataStartPos; 
	fs.write( (char *)&DataLen, sizeof(LONG)); 
	
	fs.clear();
	fs.seekp(DataEndPos);
	return fs.good();
}

CStringRecord::CStringRecord( const wstring Str
						    , const LONG Identifier )
	:m_Id(Identifier), m_Str(Str)
{
}

CStringRecord::CStringRecord( const LONG Identifier )
	:m_Id(Identifier)
{
}

void CStringRecord::Load( CStoreArchive* in, const int Length )
{
	char * buf = new char[Length];
	in->fs.read( buf, Length);
	m_Str = (wchar_t *)buf;
	delete buf;
}

void CStringRecord::Save( CStoreArchive* out ) const
{
	out->fs.write( (char *)m_Str.c_str(), (m_Str.size()+1)*sizeof(wchar_t));
}

