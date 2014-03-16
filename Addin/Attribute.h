#if !defined(AFX_ATTRIBUTE_H__5521FD88_310B_45EB_9838_462A4EFD0A76__INCLUDED_)
#define AFX_ATTRIBUTE_H__5521FD88_310B_45EB_9838_462A4EFD0A76__INCLUDED_

// This doesn't do much other than to hold a value that 
// can be easily ported between LDAP attribute and Outlook.property formats
class CAttribute  
{
	CComVariant m_Value;
	bool m_BERVAL_valid; 
	LDAP_BERVAL m_BERVAL;

public:	

	CAttribute();
	CAttribute( const CComVariant& Value);
	CAttribute( const CAttribute& attr);
	CAttribute( const wstring& str);
	virtual ~CAttribute();

	void SetValue( const CComVariant& Value);
	void SetValue( const LDAP_BERVAL * BerVal, const VARTYPE vt);


	void GetValue( CComVariant& Value) const;
	void GetValue( const LDAP_BERVAL** Value);
	const wstring GetValueStr() const;
	const DATE GetValueDate() const;

	CAttribute& operator=(const CAttribute& attr);
};

#endif // !defined(AFX_ATTRIBUTE_H__5521FD88_310B_45EB_9838_462A4EFD0A76__INCLUDED_)
