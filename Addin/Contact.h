#if !defined(AFX_CONTACT_H__28D0B907_39AA_4286_9CC9_00E5D013ACE5__INCLUDED_)
#define AFX_CONTACT_H__28D0B907_39AA_4286_9CC9_00E5D013ACE5__INCLUDED_

class CContact : public CEntry
{

	static AttribMapping ContactMap[];
	int MaxLdapAttribMap();

	typedef struct _MapiMapping {
		LPCOLESTR szOutlookName;
		CMAPIMessage::MAPIProperty_t eMAPIProp;
	} MapiMapping;

	static MapiMapping ContactMapiMap[];

	int MaxMapiMapping() const;

#ifdef MSO9_COMPATIBLE
	HRESULT oldGetContactValue( CComQIPtr<Outlook::_ContactItem>& spContact, const wstring& Name, CComVariant& Value);
	HRESULT oldPutContactValue( CComQIPtr<Outlook::_ContactItem>& spContact, const wstring& Name, CComVariant& Value);
#endif

public:

	CContact();
	virtual ~CContact() { FlushAllAttributes(); }

	virtual bool CheckOlPtr( IDispatch * pDisp );
	virtual void SetIndex( IDispatch * pDisp );
	virtual void SetTimeStamp( IDispatch * pDisp );
	virtual void SetEntryName( IDispatch* pDisp);

	virtual void Pack( IDispatch * pDisp, CMAPIFolder * pMAPIFolder );
	virtual void UnPack( IDispatch * pDisp, const bool Create );
	virtual bool Validate();

	virtual CLdapString GetLdapFilter();
	virtual void GetLdapObjectClass(CLdapAttribute &rv);
};

#endif // !defined(AFX_CONTACT_H__28D0B907_39AA_4286_9CC9_00E5D013ACE5__INCLUDED_)
