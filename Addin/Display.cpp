#include "stdafx.h"
using namespace std;

IDisplay& CDisplay::operator<<( LPCOLESTR s) 
{
	if( m_RealDisplay ) 
	{	
		CCPLWinCriticalSection::GUARD Lock(m_Crit);
		IDisplay& out = *m_RealDisplay;
		out << s;
	}
	return *this; 
}

IDisplay& CDisplay::operator<<( const wstring& s) 
{ 
	if( m_RealDisplay ) 
	{
		CCPLWinCriticalSection::GUARD Lock(m_Crit);
		IDisplay& out = *m_RealDisplay;
		out << s;
	}
	return *this; 
}

void CDisplay::RegisterDisplay(IDisplay* pDisp)
{
	CCPLWinCriticalSection::GUARD Lock(m_Crit);
	m_RealDisplay = pDisp;
}

void CDisplay::UnRegisterDisplay()
{
	CCPLWinCriticalSection::GUARD Lock(m_Crit);
	m_RealDisplay = NULL;
}
