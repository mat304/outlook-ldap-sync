#pragma once

class CCfg
{
	class CKey 
	{
		CRegKey m_Key;

	public:

		CKey() {}
		CKey( LPCOLESTR Root );
		~CKey();

		bool Open( LPCOLESTR KeyName );
		bool Create( LPCOLESTR KeyName );
		bool Attach( HKEY hKey ) { m_Key.Attach( hKey ); return true; }
		bool SetDWORD( LPCOLESTR KeyName, const DWORD Value );
		bool SetString( LPCOLESTR KeyName, const wstring Value );
		bool SetGUID( LPCOLESTR KeyName, const GUID& Guid );
		bool SetBinary( LPCOLESTR KeyName, const unsigned char* Blob, const int BlobLen);
		bool Delete( LPCOLESTR KeyName );
		bool DeleteValue( LPCOLESTR KeyName );
		bool GetString( LPCOLESTR KeyName, wstring& Value );
		bool GetGUID( LPCOLESTR KeyName, GUID& Guid );
		bool GetBinary( LPCOLESTR KeyName, unsigned char* BlobBuf, int& BlobLen );
	};

	CKey m_Settings;

	bool m_OK;
	bool m_ValidConfig; 
	wstring m_DLLDir;

	// These settings are cached here so we can tell if user modifies them
	wstring m_ServerName;
	wstring m_BaseDN;
	wstring m_UserDN;
	wstring m_UserPassword;

public:

	CCfg();
	~CCfg();

	bool OK() const { return m_OK; }

	bool SetupNewUser();
	void UserConfig( CStore * store );
	const bool ValidConfig() const { return m_ValidConfig; }

	bool Refresh( bool& Modified ); // Just update the cached settings
	const wstring GetServerName() const { return m_ServerName; }
	const wstring GetBaseDN() const {  return m_BaseDN; }
	const wstring GetUserDN() const {  return m_UserDN; }
	const wstring GetGroupDN( const wstring GroupID ) const { return L"cn="+GroupID+L",ou=Groups,"+m_BaseDN; }
	const wstring GetUserPassword() const { return m_UserPassword; }

	bool RefreshDLLDir();
	const wstring GetDLLDir() const { return m_DLLDir; }

	bool GetFirstTime();
	const wstring GetDTFileName();
	const wstring GetLogFileName();
	const wstring GetUpgradeUrl();
	const GUID GetUpgradeJob();
	const bool GetABEND();

	bool SetServerName( const wstring Value );
	bool SetBaseDN( const wstring Value );
	bool SetUserDN( const wstring Value );
	bool SetUserPassword( const wstring Value );
	bool SetABEND( const bool Value );

	bool SetFirstTime( const bool Value ); 
	bool SetUpgradeJob( const GUID Job );
};
