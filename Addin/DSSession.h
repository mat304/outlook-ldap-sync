// DSSession.h: interface for the CDSSession class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DSSESSION_H__07111B16_30AD_41F7_89F6_F38B0D552F2C__INCLUDED_)
#define AFX_DSSESSION_H__07111B16_30AD_41F7_89F6_F38B0D552F2C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ldap.h"

class CDSSession : public CLdapSession  
{
	CLdapString m_BaseDN;
public:
	CDSSession( LPCOLESTR server, LPCOLESTR baseDN );
	virtual ~CDSSession();
	CLdapString GetBaseDN() const { return m_BaseDN; }
	void GetContacts();
};

#endif // !defined(AFX_DSSESSION_H__07111B16_30AD_41F7_89F6_F38B0D552F2C__INCLUDED_)
