// Error.cpp: implementation of the CError class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
using namespace std;


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CError::CError( const wstring& desc ) 
	: m_Desc(desc) 
{
}

CError::~CError()
{
}

const wstring CError::GetDescription() 
{ 
	return m_Desc; 
}