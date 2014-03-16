#if !defined(AFX_LOCALENTRY_H__25868421_AFB4_42D0_AE5C_011EEB8A01F1__INCLUDED_)
#define AFX_LOCALENTRY_H__25868421_AFB4_42D0_AE5C_011EEB8A01F1__INCLUDED_

class CEntry  
{

protected:

	const static int MAX_LDAP_VALUES = 3;

	typedef struct _AttribMapping {
		VARTYPE iOutlookType;
		LPCOLESTR szOutlookName[MAX_LDAP_VALUES];
		LdapSyntax iLdapType;
		LPCOLESTR szLdapName;
	} AttribMapping;

	virtual int MaxLdapAttribMap() = 0;

private:
	LPCOLESTR m_EntryClass;

protected:
	AttribMapping * m_AttribMap;
	int m_OutlookVersion;
private:

	wstring m_sIndex;  // This is the master index and should be used for everything
	wstring m_sOlIndex;  // .. except Outlook & MAPI calls which should use the Outlook entry index
	DATE m_GMTTimeStamp; // Outlook item time (GMT)
	DATE m_GMTLdapTimeStamp; // LDAP item time (GMT)
	DATE m_NewGMTTimeStamp;   // Modified Outlook entry's new timestamp (temp storage)
	DATE m_NewGMTLdapTimeStamp;  // Modified LDAP entry's new timestamp (temp storage)
	map<wstring,const CAttribute*> m_Attributes;
	wstring m_EntryName;		// eg "John Matthews"
	wstring m_EntryTypeName;	// eg "Contact"

	const wstring FixLdapRecvStringSyntax( LPCOLESTR Str, LdapSyntax Syntax );
	const wstring FixLdapSendStringSyntax( LPCOLESTR Str, LdapSyntax Syntax );

	void stripPhoneNumber(CLdapObject& objContact,LPCOLESTR attrName, LPCOLESTR strippedAttrName);
	const string::size_type CEntry::findNDD(const wstring& s, const string::size_type pos, wstring& ndd, bool allowPreceedingDigits);


	// This should be read from ldap because we might sell some PBXs abroad at some time!
	static const wstring COUNTRY_CODE;
	static const wstring OUR_NDD;
	static const wstring IDD;

protected:

	CEntry( AttribMapping * map, LPCOLESTR EntryClass, LPCOLESTR EntryTypeName ) 
		: m_EntryClass(EntryClass),
		  m_AttribMap( map),
		  m_sIndex(L""), /* Special value meaning no index set yet */
		  m_sOlIndex(L""), /* Special value meaning no index set yet */
		  m_EntryTypeName(EntryTypeName)
	{}

	// Attribute fns
	void PutAttribute( LPCOLESTR Name, const CAttribute* Value );
	bool FindAttribute( LPCOLESTR Name, const CAttribute*& pValue );
	void EraseAttribute( LPCOLESTR Name );
	void FlushAllAttributes();

public:

	// Get/Set human readable description for the entry ( eg "Contact[John Matthews]")
	const wstring GetEntryShortName() const { return m_EntryName; }
	const wstring GetEntryTypeName() const { return m_EntryTypeName; }
	const wstring GetEntryName() const { return m_EntryTypeName + L"(" + m_EntryName + L")"; }
	virtual void SetEntryName( IDispatch* pDisp) = 0;
	void SetEntryNameStr( LPCOLESTR name ) { m_EntryName = name; }

	// Get the MAPI entry class ( eg "IPM.Contact")
	LPCOLESTR GetEntryClass() const;

	// Factory method. Create an entry of the appropriate types
	// eg) CContact
	static CEntry * Make( const Outlook::OlDefaultFolders EntryType, int OutlookVersion );
	virtual ~CEntry() { FlushAllAttributes(); }
	virtual bool CheckOlPtr( IDispatch * pDisp ) = 0;

	// These operate on m_sIndex & m_sOlIndex
	const wstring GetIndex() const { return m_sIndex; }
	const wstring GetOlIndex() const { return m_sOlIndex; }
	virtual void SetIndex( IDispatch * pDisp ) = 0;
	void SetIndexStr( LPCOLESTR szIndex );
	void SetOlIndexStr( LPCOLESTR szIndex );

	// These operate on the entry's Outlook timestamps
	virtual void SetTimeStamp( IDispatch * pDisp ) = 0;
	void SetTimeStampDate( const DATE& date );
	const DATE GetTimeStamp() const { return m_GMTTimeStamp; }
	const DATE GetNewTimeStamp() const { return m_NewGMTTimeStamp; }
	void SetNewTimeStamp( const DATE& date ) { m_NewGMTTimeStamp = date; }

	// These operate on the entry's LDAP timestamps
	void SetLdapTimeStamp( const DATE& date);
	const DATE GetLdapTimeStamp() const { return m_GMTLdapTimeStamp; }
	const DATE GetNewLdapTimeStamp() const { return m_NewGMTLdapTimeStamp; }
	void SetNewLdapTimeStamp( const DATE& date ) { m_NewGMTLdapTimeStamp = date; }

	// Compare LDAP and local entry timestamps
	// Returns : 0 if identical
	//			<0 if local entry is newer
	//			>0 if LDAP entry is newer
	const int CompareTimeStamps() const
	{ 
		if( m_NewGMTLdapTimeStamp == m_NewGMTTimeStamp )
			return 0;

		if( m_NewGMTLdapTimeStamp > m_NewGMTTimeStamp )
			return 1;

		return -1;
	}

	void CommitNewTimeStamps() 
	{ 
		m_GMTTimeStamp = m_NewGMTTimeStamp; 
		m_GMTLdapTimeStamp = m_NewGMTLdapTimeStamp;	
	}

	// Display the entry for debugging purposes
	void DebugBox();
	void LdapDebugBox(CLdapObject& obj);

	virtual void Pack( IDispatch * pDisp, CMAPIFolder * pMAPIFolder ) = 0;
	virtual void UnPack( IDispatch * pDisp, const bool Create ) = 0;
	virtual bool Validate() = 0;

	void Recv( CLdapSession& sess, LPCOLESTR BaseDN );
	void Send( CLdapSession& sess, LPCOLESTR BaseDN, const bool Create );

	virtual CLdapString GetLdapFilter() = 0;
	virtual void GetLdapObjectClass( CLdapAttribute& rv) = 0;

	// Utility fns for converting between different awkward time formats
	static const wstring GMTVariantTimeToLdapTime( const DATE& in);
	static const DATE LdapTimeToGMTVariantTime( const wstring& in);
	static const wstring GMTVariantTimeToStr( const DATE& in);
	static const DATE VariantTimeToGMTVariantTime( const DATE& in);
	static const DATE GMTVariantTimeToVariantTime( const DATE& in);
	static const DATE VariantTimeAddSeconds( const DATE& in, int Secs );
};

#endif // !defined(AFX_LOCALENTRY_H__25868421_AFB4_42D0_AE5C_011EEB8A01F1__INCLUDED_)
