// CfgDialog.cpp : Implementation of CCfgDialog
#include "stdafx.h"
#include "cfgdialog.h"
using namespace std;

bool CCfgDialog::DecodeUserID( const wstring userid, wstring& userdn, wstring& basedn ) const
{
	// Convert userid = john.matthews@groupware.com
	// to userdn = uid=john.matthews,cn=Users,dc=groupware,dc=com
	// & basedn = dc=groupware,dc=com
	wstring::size_type s = userid.find_first_of(L"@");
	if( s == string::npos ) return false;

	wstring user = userid.substr(0,s);
	wstring company = userid.substr(s+1);

	bool NoDots = true;
	for( s=company.find_first_of(L".") ;s!=string::npos ;s=company.find_first_of(L".",s) )
	{
		company[s] = L',';
		company.insert(s+1, L"dc=");
		NoDots = false;
	}
	if( NoDots ) return false;

	userdn = L"uid="+user+L",cn=Users,dc="+company;
	basedn = L"dc="+company;

	return true;
}

const wstring CCfgDialog::EncodeUserID( const wstring userdn ) const
{
	// Convert userdn = uid=john.matthews,cn=Users,dc=groupware,dc=com
	// into userid = john.matthews@groupware.com
	wstring::size_type s = userdn.find(L"uid=");
	wstring::size_type c = userdn.find(L",cn=Users,dc=");
	if( (s==string::npos) || (c==string::npos) )
		return L"";

	wstring user = userdn.substr(s+4,c-(s+4));
	wstring company = userdn.substr(c+13);

	for( s=company.find(L",dc=") ;s!=string::npos ;s=company.find(L",dc=",s) )
	{
		company[s] = L'.';
		company.erase(s+1,3);
	}

	return user+L"@"+company;
}

LRESULT CCfgDialog::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	// Set the product name and version static text
	//wchar_t ProductVersion[ 128 ];
	// Doesn't work under Win98
	//CHECKNOTZERO(LoadString(_Module.GetResourceInstance(),
	//	IDS_VERSIONSTR, ProductVersion, sizeof(ProductVersion)));
	const wstring VerStr( wstring(L"( V") + m_Store->m_Addin->GetVersion() + wstring(L" )") );
	SetDlgItemText( IDC_VERSION, VerStr.c_str() );
	SetDlgItemText( IDC_SERVERNAME, m_Cfg->GetServerName().c_str() );
	SetDlgItemText( IDC_USERID, EncodeUserID(m_Cfg->GetUserDN()).c_str() );
	SetDlgItemText( IDC_PASSWORD, m_Cfg->GetUserPassword().c_str() );

	//m_Logo =(HBITMAP)LoadImage(_Module.GetResourceInstance(),
	//	MAKEINTRESOURCE(IDB_LOGO),IMAGE_BITMAP,0,0,0);

	return 1;  // Let the system set the focus
}

LRESULT CCfgDialog::OnPaint(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	PAINTSTRUCT ps; 
    BeginPaint( &ps); 
	//HDC hdc = CreateCompatibleDC( ps.hdc );
	//SelectObject(hdc, m_Logo);

	//BITMAP bm;
	//GetObject(m_Logo,sizeof(BITMAP), &bm );
//	StretchBlt( ps.hdc, 41, 12, 48, 48, hdc, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY );
	//BitBlt( ps.hdc, 51, 12, 78, 42, hdc, 0, 0, SRCCOPY );

	//DeleteDC(hdc);
	EndPaint(&ps);
	return 0;
}

LRESULT CCfgDialog::OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	wchar_t Value[128];

	GetDlgItemText( IDC_SERVERNAME, Value, sizeof(Value) );
	if( lstrlen(Value) == 0 )
	{
		MessageBox( L"Bad GroupWare server name\n"
					L"Please enter network name of server (eg ldap.company.com)" 
				, L"GroupWare Server", MB_OK);
		return 1;
	}
	m_Cfg->SetServerName( Value );

	GetDlgItemText( IDC_USERID, Value, sizeof(Value) );
	wstring basedn, userdn;
	if( ! DecodeUserID( Value, userdn, basedn ) )
	{
		MessageBox( L"Bad user id!\n"
					L"Please enter user id in email format (eg first.last@company.com)" 
				, L"GroupWare Server", MB_OK);
		return 1;
	}
	m_Cfg->SetUserDN( userdn );
	m_Cfg->SetBaseDN( basedn );

	GetDlgItemText( IDC_PASSWORD, Value, sizeof(Value));
	if( (lstrlen(Value) == 0) || (wstring(Value) == L"UNSET") )
	{
		MessageBox( L"Bad password\n"
					L"Please enter your usual password" 
				, L"GroupWare Server", MB_OK);
		return 1;
	}
	m_Cfg->SetUserPassword( Value );

	m_Store->ReadConfig();

	EndDialog(wID);
	return 0;
}

LRESULT CCfgDialog::OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// Ignore the new values
	EndDialog(wID);
	return 0;
}

LRESULT CCfgDialog::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	DeleteObject(m_Logo);
	return 0;
}
