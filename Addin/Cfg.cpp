#include "stdafx.h"
#include "cfg.h"
using namespace std;

CCfg::CKey::CKey( LPCOLESTR Root ) 
{
	HKEY hkIA;
	RegOpenKeyEx( HKEY_CURRENT_USER, Root, 0, KEY_ALL_ACCESS, &hkIA);
	m_Key.Attach(hkIA);
}

CCfg::CKey::~CKey()
{
}

bool CCfg::CKey::Open( LPCOLESTR KeyName )
{
	return ( m_Key.Open( m_Key.m_hKey, KeyName) == ERROR_SUCCESS );
}

bool CCfg::CKey::Create( LPCOLESTR KeyName)
{
	return ( m_Key.Create( m_Key.m_hKey, KeyName) == ERROR_SUCCESS );
}

bool CCfg::CKey::SetDWORD( LPCOLESTR KeyName, const DWORD Value)
{
	return( m_Key.SetDWORDValue( KeyName, Value) == ERROR_SUCCESS ); 
}

bool CCfg::CKey::SetString( LPCOLESTR KeyName, const wstring Value)
{
	return( m_Key.SetStringValue( KeyName, (LPCTSTR)Value.c_str()) == ERROR_SUCCESS );
}

bool CCfg::CKey::SetGUID( LPCOLESTR KeyName, const GUID& Guid )
{
	return( m_Key.SetGUIDValue( KeyName, Guid ) == ERROR_SUCCESS );
}

bool CCfg::CKey::SetBinary( LPCOLESTR KeyName, const unsigned char* Blob, const int BlobLen)
{
	return( m_Key.SetBinaryValue( KeyName, Blob, BlobLen) == ERROR_SUCCESS ); 
}

bool CCfg::CKey::Delete( LPCOLESTR KeyName )
{
	return( m_Key.RecurseDeleteKey( KeyName ) == ERROR_SUCCESS );
}

bool CCfg::CKey::DeleteValue( LPCOLESTR KeyName )
{
	return( m_Key.DeleteValue( KeyName ) == ERROR_SUCCESS );
}

bool CCfg::CKey::GetString( LPCOLESTR KeyName, wstring& Value )
{
	Value = L"";
	wchar_t WValue[128];
	ULONG ValueLen = sizeof(WValue);
	if( ( m_Key.QueryStringValue( KeyName, (LPTSTR)WValue, &ValueLen) != ERROR_SUCCESS ) || (ValueLen == 0) )
		return false;
	Value = WValue;
	return true;
}

bool CCfg::CKey::GetGUID( LPCOLESTR KeyName, GUID& Guid )
{
	return ( m_Key.QueryGUIDValue( KeyName, Guid ) == ERROR_SUCCESS );
}

bool CCfg::CKey::GetBinary( LPCOLESTR KeyName, unsigned char* BlobBuf, int& BlobLen )
{
	return( m_Key.QueryBinaryValue( KeyName, (void*)BlobBuf, (ULONG*)&BlobLen ) == ERROR_SUCCESS );
}

CCfg::CCfg()
	:m_OK(false), m_ValidConfig(false)
{
	// Default registry keys are under HKEY_LOCAL_MACHINE for all users install
	// If this is the first time this user has used the addin then we need to set them
	// up with a default user configuration
	HKEY hkey, hkeyLM;
	long HKCURC = RegOpenKeyEx( HKEY_CURRENT_USER, L"Software\\GroupWare", 0, KEY_ALL_ACCESS, &hkey);
	long HKLMRC = RegOpenKeyEx( HKEY_LOCAL_MACHINE, L"Software\\GroupWare", 0, KEY_ALL_ACCESS, &hkeyLM);
	if( (HKCURC != ERROR_SUCCESS) && (HKLMRC == ERROR_SUCCESS) )
	{	
		HKCURC = RegCreateKeyEx( HKEY_CURRENT_USER, L"Software\\GroupWare", 
						0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
		if( HKCURC == ERROR_SUCCESS ) 
		{
			RegCloseKey(hkeyLM);
			m_Settings.Attach(hkey);
			m_OK = true; // .. but need to call SetupNewUser()
			return;
		}
	}
	if( HKLMRC == ERROR_SUCCESS ) RegCloseKey(hkeyLM);
	if( HKCURC != ERROR_SUCCESS ) return;

	m_OK = true;

	m_Settings.Attach(hkey);

	// Check user config is valid
	wchar_t Buf[64];
	int BufLen = sizeof(Buf);
	wstring Value;
	m_ValidConfig = (  m_Settings.GetString( L"ServerName",Value )
					&& m_Settings.GetString( L"BaseDN", Value )
					&& m_Settings.GetString( L"UserDN", Value )
					&& m_Settings.GetBinary( L"Password", (unsigned char *)Buf, BufLen )
					&& m_Settings.GetString( L"UpgradeUrl", Value )
					&& m_Settings.GetString( L"FirstTime", Value )
					&& m_Settings.GetString( L"DTFileName", Value ) );

	bool Dummy;
	if( m_ValidConfig ) Refresh(Dummy);
}

CCfg::~CCfg()
{
}

bool CCfg::Refresh( bool& Modified )
{
	Modified = false;
	wstring Value;

	if( ! m_Settings.GetString( L"ServerName", Value ) ) return false;
	if( Value != m_ServerName )	Modified = true;
	m_ServerName = Value;

	if( ! m_Settings.GetString( L"BaseDN", Value ) ) return false;
	if( Value != m_BaseDN )	Modified = true;
	m_BaseDN = Value;

	if( ! m_Settings.GetString( L"UserDN", Value ) ) return false;
	if( Value != m_UserDN )	Modified = true;
	m_UserDN = Value;

	wchar_t Buf[64];
	int BufLen = sizeof(Buf);
	if( ! m_Settings.GetBinary( L"Password", (unsigned char *)Buf, BufLen ) ) return false;
	for( UINT i=0; i<(BufLen/sizeof(wchar_t)); i++ ) Buf[i] -= 0xa8;
	m_UserPassword = wstring(Buf,(BufLen/sizeof(wchar_t)));

	return true;
}

bool CCfg::SetupNewUser()
{
	CKey DefaultSettings;
	HKEY hkeyLM;
	long HKLMRC = RegOpenKeyEx( HKEY_LOCAL_MACHINE, L"Software\\GroupWare", 0, KEY_ALL_ACCESS, &hkeyLM);
	if( HKLMRC != ERROR_SUCCESS ) return false;
	DefaultSettings.Attach(hkeyLM);

	wstring Value;

	if( DefaultSettings.GetString(L"ServerName",Value) ) 
		m_Settings.SetString(L"ServerName",Value);

	if( DefaultSettings.GetString(L"BaseDN",Value) )
		m_Settings.SetString(L"BaseDN",Value);

	if( DefaultSettings.GetString(L"UserDN",Value) )
		m_Settings.SetString(L"UserDN",Value);

	if( DefaultSettings.GetString(L"DTFileName",Value) )
		m_Settings.SetString(L"DTFileName",Value);

	if( DefaultSettings.GetString( L"UpgradeUrl", Value ) )
		m_Settings.SetString( L"UpgradeUrl", Value );

	SetUserPassword(L"");

	GUID Guid = { 0, 0, 0, 0};
	m_Settings.SetGUID( L"UpgradeJob", Guid );
	m_Settings.SetString(L"FirstTime",L"yes");

	m_ValidConfig = true;

	bool Dummy;
	Refresh(Dummy);

	return true;
}

void CCfg::UserConfig( CStore * store )
{
	CCfgDialog dlg( this, store );
	dlg.DoModal();
}

bool CCfg::SetServerName( const wstring Value ) 
{ 
	return m_Settings.SetString(L"ServerName", Value); 
}

bool CCfg::SetBaseDN( const wstring Value ) 
{  
	return m_Settings.SetString(L"BaseDN", Value); 
}

bool CCfg::SetUserDN( const wstring Value ) 
{
	return m_Settings.SetString(L"UserDN", Value); 
}

bool CCfg::SetUserPassword( const wstring Value ) 
{ 	
	wchar_t* Buf = new wchar_t[Value.size()];
	memcpy( Buf, (char *)Value.c_str(), Value.size()*sizeof(wchar_t) );
	for( UINT i=0; i<Value.size(); i++ ) Buf[i] += 0xa8;
	return m_Settings.SetBinary(L"Password", (unsigned char *)Buf, Value.size()*sizeof(wchar_t)); 
}

bool CCfg::SetFirstTime( const bool Value ) 
{ 
	return 	m_Settings.SetString( L"FirstTime", Value ? L"yes" : L"no" );
}

bool CCfg::SetUpgradeJob( const GUID Job )
{
	return 	m_Settings.SetGUID( L"UpgradeJob", Job );
}

bool CCfg::GetFirstTime()
{ 
	wstring Value;
	if( ! m_Settings.GetString( L"FirstTime", Value ) ) return true;
	return ( Value == L"yes" );
}

const wstring CCfg::GetDTFileName()
{
	wstring Value, Path;
	// Prepend users appdata path to DT filename
	if( ! m_Settings.GetString( L"DTFileName", Value ) ) return L"C:\\";
	CKey appdata(L"Volatile Environment");
	if( !appdata.GetString( L"APPDATA", Path ) ) return L"C:\\";
	return Path + wstring(L"\\") + Value;
}

const wstring CCfg::GetLogFileName()
{ 
	wstring Value;
	m_Settings.GetString( L"LogFileName", Value );
	return Value; 
}

const GUID CCfg::GetUpgradeJob()
{ 
	GUID Value;
	m_Settings.GetGUID( L"UpgradeJob", Value );
	return Value; 
}

const wstring CCfg::GetUpgradeUrl()
{ 
	wstring Value;
    m_Settings.GetString( L"UpgradeUrl", Value );
	return Value; 
}

bool CCfg::RefreshDLLDir()
{
	// Find out where the hell windows installed us
	CKey Classes;
	HKEY hkey;
	// Lookup the CLSID entry for ourselves
	if(RegOpenKeyEx( HKEY_CLASSES_ROOT, L"CLSID\\{b33767cf-240a-4ec5-a978-ae685707ba6c}\\InprocServer32", 0, KEY_READ, &hkey) != ERROR_SUCCESS )
		return false;

	wchar_t Val[128];
	LONG ValLen = sizeof(Val);
	if( RegQueryValue( hkey, NULL, (LPWSTR)&Val, &ValLen ) != ERROR_SUCCESS )
	{
		RegCloseKey( hkey );
		return false;
	}

	RegCloseKey( hkey );

	wstring FilePath(Val,ValLen);

	// Chop the "\GroupWare.dll" bit off the end
	wstring::size_type LastSlash = FilePath.find_last_of(L'\\');
	if( (LastSlash == string::npos) || (LastSlash == 0))
		return false;

	// Chop the preceeding quote off the front
	if( FilePath[0] == L'"' )
		m_DLLDir = FilePath.substr(1,--LastSlash);
	else
		m_DLLDir = FilePath.substr(0,LastSlash);

	return true;
}

bool CCfg::SetABEND( const bool Value )
{
	return 	m_Settings.SetString( L"ABENDCheck", Value ? L"yes" : L"no" );
}

const bool CCfg::GetABEND()
{
	wstring Value;
    if( ! m_Settings.GetString( L"ABENDCheck", Value ) )
		return false;
	return(Value==L"yes"); 
}




