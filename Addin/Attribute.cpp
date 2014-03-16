// Attribute.cpp: implementation of the CAttribute class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
using namespace std;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CAttribute::CAttribute()
	: m_BERVAL_valid(false)
{
}

CAttribute::CAttribute( const CComVariant& Value)
	: m_BERVAL_valid(false), m_Value(Value)
{
}

CAttribute::CAttribute( const wstring& str)
	: m_BERVAL_valid(false)
{
	VARIANT var;
	::VariantInit( &var);
	V_BSTR(&var) = ::SysAllocString(const_cast<wchar_t *>(str.c_str()));
	V_VT(&var) = VT_BSTR;
	m_Value.Attach( &var);
}

// Copy constructor
CAttribute::CAttribute( const CAttribute& attr)
	: m_Value( attr.m_Value )
{
	if( attr.m_BERVAL_valid)
	{
		m_BERVAL.bv_len = attr.m_BERVAL.bv_len;
		m_BERVAL.bv_val = new char[m_BERVAL.bv_len];
		memcpy( m_BERVAL.bv_val, attr.m_BERVAL.bv_val, m_BERVAL.bv_len);
	}
	m_BERVAL_valid = attr.m_BERVAL_valid;
}

CAttribute::~CAttribute()
{
	if( m_BERVAL_valid) 
		delete [] m_BERVAL.bv_val;
}

void CAttribute::SetValue( const CComVariant& Value)
{
	m_Value = Value;
}

void CAttribute::SetValue( const LDAP_BERVAL * BerVal, const VARTYPE vt)
{
	// Copy the BERVAL and create a CComVariant of the supplied type
	m_BERVAL.bv_len = BerVal->bv_len;
	m_BERVAL.bv_val = new char[ m_BERVAL.bv_len ];
	memcpy( m_BERVAL.bv_val, BerVal->bv_val, m_BERVAL.bv_len);
	m_BERVAL_valid = true;

	VARIANT var;
	::VariantInit( &var);
	
	switch(vt) {
	case VT_BSTR:
		{
			// BerVal must include the NULL string terminator
			BSTR bstr = reinterpret_cast<wchar_t * const>(BerVal->bv_val);
			if (bstr == NULL) {
				V_BSTR(&var) = NULL;
			}
			else {
				V_BSTR(&var) = ::SysAllocStringByteLen(reinterpret_cast<char*>(bstr),
									/*::SysStringByteLen(bstr)*/ BerVal->bv_len);
			}
			V_VT(&var) = VT_BSTR;
		}
		break;
	case VT_DATE:
		CHECKTRUE( BerVal->bv_len == sizeof(DATE) );
		V_DATE(&var) = *(reinterpret_cast<const DATE *>(BerVal->bv_val));
		V_VT(&var) = VT_DATE;
		break;
	default:
		CHECKTRUE(false);
		break;
	}

	m_Value.Attach( &var);
}

void CAttribute::GetValue( CComVariant& Value) const
{
	Value = m_Value; 
}

void CAttribute::GetValue( const LDAP_BERVAL * * pValue)
{
	// Get value as a BERVAL. Must convert CComVariant into BERVAL if not valid
	if( m_BERVAL_valid) 
	{
		*pValue = &m_BERVAL;
		return;
	}

	CComBSTR str;
	switch( m_Value.vt) {
	case VT_BSTR:
		CHECKHR(str.AppendBSTR(m_Value.bstrVal));
		m_BERVAL.bv_len = (str.Length() + 1) * sizeof(wchar_t);
		m_BERVAL.bv_val = new char[m_BERVAL.bv_len];
		memcpy( m_BERVAL.bv_val, str.m_str, (str.Length() * sizeof(wchar_t)) );
		memset( &m_BERVAL.bv_val[str.Length() * sizeof(wchar_t)], 0, sizeof(wchar_t));
		m_BERVAL_valid = true;
		break;
	case VT_DATE:
		m_BERVAL.bv_len = sizeof(DATE);
		m_BERVAL.bv_val = new char[m_BERVAL.bv_len];
		*((DATE *)m_BERVAL.bv_val) = m_Value.date;
		m_BERVAL_valid = true;
		break;
	case VT_EMPTY:
		m_BERVAL.bv_len = 0;
		m_BERVAL.bv_val = NULL;
		m_BERVAL_valid = true;
		break;
	default:
		CHECKTRUE(false);
		break;
	}

	*pValue = &m_BERVAL;
}

const wstring CAttribute::GetValueStr() const
{
	CComVariant vTemp(m_Value);
	CHECKHR( vTemp.ChangeType( VT_BSTR));
	return wstring(vTemp.bstrVal);
}

const DATE CAttribute::GetValueDate() const
{
	CComVariant vTemp(m_Value);
	CHECKHR( vTemp.ChangeType( VT_DATE));
	return vTemp.date;
}

// Assignment operator only copies attribute value
CAttribute& CAttribute::operator=(const CAttribute& attr)
{
	if( this == &attr) return *this; // self assignment

	m_Value = attr.m_Value;

	if( m_BERVAL_valid) 
		delete [] m_BERVAL.bv_val;

	if( attr.m_BERVAL_valid)
	{
		m_BERVAL.bv_len = attr.m_BERVAL.bv_len;
		m_BERVAL.bv_val = new char[m_BERVAL.bv_len];
		memcpy( m_BERVAL.bv_val, attr.m_BERVAL.bv_val, m_BERVAL.bv_len);
	}
	m_BERVAL_valid = attr.m_BERVAL_valid;

	return *this;
}


