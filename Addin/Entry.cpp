// Entry.cpp: implementation of the CLocalEntry class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include <time.h>
using namespace std;
	
const wstring CEntry::COUNTRY_CODE = L"61";
const wstring CEntry::OUR_NDD = L"0";
const wstring CEntry::IDD = L"0011";


// Add/Modify named attribute
void CEntry::PutAttribute( LPCOLESTR Name, const CAttribute* Value )
{
	CAttribute* pExistingValue;
	if( FindAttribute( Name, pExistingValue ) ) 
	{
		m_Attributes[wstring(Name)] = Value;
		delete pExistingValue; 
		return;
	}
	m_Attributes[wstring(Name)] = Value;
}

// Find named Attribute and return pointer to it
bool CEntry::FindAttribute( LPCOLESTR Name, const CAttribute*& pValue )
{
	map<wstring,const CAttribute*>::const_iterator i = m_Attributes.find( wstring(Name) );
	pValue = i->second;
	return ( i != m_Attributes.end() );
}
	
// Find NDD which will be in (<digit>) format and return the NDD
const string::size_type CEntry::findNDD(const wstring& s, const string::size_type pos, wstring& ndd, bool allowPreceedingDigits) {
	bool brackets = false;
	for(string::size_type i=pos; i<s.size(); i++) {
		wchar_t c = s[i];
		if(c==L'(') brackets=true;
		if(c==L')') {
			if(ndd.size()>0) 
				return i+1;
			else break;
		}
		if(c>=L'0'&&c<=L'9') {
			if(brackets) 
				ndd.append(1,c);
			else if(!allowPreceedingDigits) break;
		}
	}
	ndd.clear();
	return pos;
}
	
// Strip formatting from a phone number
void CEntry::stripPhoneNumber(CLdapObject& objContact, LPCOLESTR attrName, LPCOLESTR strippedAttrName)
{
	CLdapAttribute attr;
	if( !objContact.GetAttribute( attrName, attr) || (attr.ValueCount()!=1))
		return;
	wstring strippedNumber(attr[0]);
	// Strip out anything that isn't a digit + or ()
	wstring::size_type pos = strippedNumber.find_first_not_of(L"0123456789+()");
	while( pos!=wstring::npos) {
		strippedNumber = strippedNumber.replace(pos,1,L"");
		pos = strippedNumber.find_first_not_of(L"0123456789+()");
	}
	// Check international dial codes
	if((strippedNumber[0]=='+')&&(strippedNumber.find(COUNTRY_CODE)==1)) {
		// Interational notation for number within our country
		wstring ndd;
		string::size_type i = COUNTRY_CODE.size()+1;
		string::size_type nddEnd = findNDD(strippedNumber,i,ndd,false);
		if((nddEnd!=i)&&(ndd==OUR_NDD)) 
			strippedNumber.replace(0,nddEnd-1,wstring(L""));
		else
			strippedNumber.replace(0,i,wstring(L""));
		strippedNumber = OUR_NDD + strippedNumber;
	} else if(strippedNumber[0]=='+') {
		// Foreign number
		wstring ndd;
		string::size_type i = 1;
		string::size_type nddEnd = findNDD(strippedNumber,i,ndd,true);
		if(nddEnd!=i) {
			// Remove the ndd
			string::size_type start = strippedNumber.find_first_of(L"(");
			if((start!=string::npos)&&(nddEnd>start))
				strippedNumber.replace(start,nddEnd-start,wstring(L""));
		}
		// Add the idd
		strippedNumber = IDD + strippedNumber;
	}
	// Strip out anything that isn't a digit
	pos = strippedNumber.find_first_not_of(L"0123456789");
	while( pos!=wstring::npos) {
		strippedNumber = strippedNumber.replace(pos,1,L"");
		pos = strippedNumber.find_first_not_of(L"0123456789");
	}
	objContact.PutAttribute( strippedAttrName, CLdapValue(strippedNumber.c_str()));
}

void CEntry::EraseAttribute( LPCOLESTR Name )
{
	map<wstring,const CAttribute*>::const_iterator i = m_Attributes.find( wstring(Name) );
	if( i != m_Attributes.end() )
	{
		const CAttribute * attr = i->second;
		m_Attributes.erase(Name);
		delete attr;
	}
}

void CEntry::FlushAllAttributes()
{
	map<wstring,const CAttribute*>::const_iterator i;
	for( i = m_Attributes.begin(); i!= m_Attributes.end(); ++i)
		delete i->second;
	m_Attributes.erase( m_Attributes.begin(), m_Attributes.end());
}

void CEntry::DebugBox()
{
	wostringstream os;

	wstring out = os.str();
	out += L'\n';

	map<wstring,const CAttribute*>::const_iterator i;
	for( i = m_Attributes.begin(); i!= m_Attributes.end(); ++i)
	{
		const CAttribute* attr = i->second; 
		out += i->first + L'=' + attr->GetValueStr() + L'\n';
	}

	wstring title(L"Entry:");
	title += GetEntryName();

	MessageBoxW(NULL, out.c_str(), title.c_str(), MB_SETFOREGROUND);
}

void CEntry::LdapDebugBox(CLdapObject& obj)
{
	wstring s; 
	s += L"DN=";
	s += obj.GetDN();
	s += L"\n\n";
	for( CLdapObject::AttributeIterator iter = obj.AttributeBegin(); iter != obj.AttributeEnd(); iter++ )
	{
		const CLdapAttribute& attr = iter.Attribute();
		s += L"Attribute \"" + iter.Name() + L"\"\n";
		int nValues = attr.ValueCount();
		if( nValues == 1 ) {
			const CLdapValue& val = attr[ 0 ];
			s += L"	\"";
			if( val.GetSyntaxCategory() == 2 )
				s += val.GetUnicodeString();
			s += L"\"\n";
		} else {
			for( int nVal = 0; nVal < nValues; nVal++ ) {
				const CLdapValue& val = attr[ nVal ];
				s += L"	\"";
				if( val.GetSyntaxCategory() == 2 )
					s += val.GetUnicodeString();
				s += L"\"\n";
			}
		}
	}
	MessageBoxW(NULL, s.c_str(), L"LDAP entry", MB_SETFOREGROUND);
}

// Convert those odd LDAP string syntaxes into something more normal
const wstring CEntry::FixLdapRecvStringSyntax( LPCOLESTR StrW, LdapSyntax Syntax )
{
	wstring Str(StrW);

	switch(Syntax) {
		case SyntaxPostalAddress:
			// $ characters must be replaced by newlines !
			string::size_type s;
			for( s = Str.find_first_of(L"$"); s != string::npos; s = Str.find_first_of(L"$" ,s) )
			{
				Str.erase( s, 1 );
				Str.insert( s, 1, L'\n' );
			}
			break;
	}

	return Str;
}

void CEntry::Recv( CLdapSession& sess, LPCOLESTR BaseDN )
{
	
	CLdapSearch objSearch( sess );
	wstring szContactDN = GetIndex() + L"," + BaseDN;
	
	// Contact ID is a unique index so there can at most one object returned
	CLdapObject objContact;
	try
	{
		objSearch.Search( szContactDN.c_str(), LDAP_SCOPE_BASE );
		if( !objSearch.GetNextObject( objContact ) )
			throw CError( L"Server search for entry failed");
	} catch (CLdapException&  )
	{
		throw CError( L"Server search for entry failed");
	}

	FlushAllAttributes();

	for( int i=0; i<MaxLdapAttribMap(); i++) 
	{
		// Check we want to touch this LDAP attribute
		if( ! m_AttribMap[i].szLdapName )
			continue;

		// Get the attribute which is optional in the LDAP entry
		// Unpack will clear the fields where CAttributes don't exist
		CLdapAttribute objLdapAttr;
		if( ! objContact.GetAttribute( m_AttribMap[i].szLdapName, objLdapAttr ) )
			continue;

		for( int ValueNum = 0; ValueNum < objLdapAttr.ValueCount(); ValueNum++ )
		{

			if( ! m_AttribMap[i].szOutlookName[ValueNum] ) 
				continue;

			CLdapValue objValue((CLdapValue)objLdapAttr[ValueNum]);

			// Convert the attribute value 
			CComVariant Value;
			switch( m_AttribMap[i].iOutlookType ) {
			case VT_BSTR:
				CHECKTRUE( COIDSyntaxMap::GetEntryForSyntax( m_AttribMap[i].iLdapType )->catSyntax == CategoryString );
				Value = FixLdapRecvStringSyntax((LPCOLESTR)objValue, m_AttribMap[i].iLdapType).c_str();
				// Check for dumy LDAP attribute value
				if( wstring(Value.bstrVal) == L" " ) continue;
				break;
			case VT_DATE:
				CHECKTRUE( COIDSyntaxMap::GetEntryForSyntax( m_AttribMap[i].iLdapType )->catSyntax == CategoryString );
				Value = LdapTimeToGMTVariantTime( wstring((LPCOLESTR)objValue) );
				CHECKHR(Value.ChangeType(VT_DATE));
				break;
			case VT_I4:
				CHECKTRUE( COIDSyntaxMap::GetEntryForSyntax( m_AttribMap[i].iLdapType )->catSyntax == CategoryString );
				Value = (int)objValue;
				CHECKHR(Value.ChangeType(VT_I4));
				break;
			default:
				CHECKTRUE(false); // Need to write a conversion handler
			}
			CAttribute* pAttrib = new CAttribute(Value);
			PutAttribute( m_AttribMap[i].szOutlookName[ValueNum], pAttrib );
		}
	}
}

// Ensure that we conform to those odd LDAP string syntaxes 
const wstring CEntry::FixLdapSendStringSyntax( LPCOLESTR StrW, LdapSyntax Syntax )
{
	wstring Str(StrW);

	switch(Syntax) {
		case SyntaxPostalAddress:
			// Newlines must be replaced by $ characters !
			string::size_type s;
			for( s = Str.find_first_of(L"\n\r"); s != string::npos; s = Str.find_first_of(L"\n\r" ,s) )
			{
				Str.erase( s, (Str.find_first_of(L"\n\r" ,s+1) == (s+1))?2:1 );
				Str.insert( s, 1, L'$' );
			}
			break;
	}

	return Str;
}

void CEntry::Send( CLdapSession& sess, LPCOLESTR BaseDN, const bool Create )
{

	wstring szContactDN = GetIndex() + L"," + BaseDN;
	CLdapObject objContact( szContactDN.c_str() );
	CLdapError e;

	try {
		CLdapSearch objSearch( sess );
		objSearch.Search( szContactDN.c_str(), LDAP_SCOPE_BASE );
		bool exists = objSearch.GetNextObject( objContact );
		if (exists) {
			e = objContact.Delete( sess );
			if( !e.Success() ) {
				wstring msg(L"Failed while deleting entry from server : ");
				msg += e.Description();
				throw CError( msg );
			}
		}
	} catch (CLdapException&  ) {
		throw CError( L"Server search for entry failed");
	}

	for( int i=0; i<MaxLdapAttribMap(); i++) 
	{
		for( int ValueNum = 0; ValueNum<MAX_LDAP_VALUES; ValueNum++ ) 
		{
			// Check we want to put this attribute in the LDAP entry
			if( ! m_AttribMap[i].szLdapName || ! m_AttribMap[i].szOutlookName[ValueNum] )
				continue;
			objContact.DelAttribute( m_AttribMap[i].szLdapName );
		}
	}

	CLdapAttribute objClass;
	GetLdapObjectClass( objClass );
	if(!objContact.ExistsAttribute(L"objectClass"))
		objContact.PutAttribute( L"objectClass", objClass );
	// Don't include the <attrname>= part of the index if there is one!
	if(!objContact.ExistsAttribute(L"guid")) {
		wstring idx(GetIndex());
		wstring::size_type pos = idx.find(L'=');
		if(pos==string::npos )
			objContact.PutAttribute( L"guid", GetIndex().c_str() );
		else {
			wstring idxVal(idx,pos+1);
			objContact.PutAttribute( L"guid", idxVal.c_str() );
		}
	}

	for( int i=0; i<MaxLdapAttribMap(); i++) 
	{
		for( int ValueNum = 0; ValueNum<MAX_LDAP_VALUES; ValueNum++ ) 
		{
			// Check we want to put this attribute in the LDAP entry
			if( ! m_AttribMap[i].szLdapName || ! m_AttribMap[i].szOutlookName[ValueNum] )
				continue;

			// A missing CAttribute implies a missing/deleted LDAP attribute
			const CAttribute* pAttr;
			if( !FindAttribute(m_AttribMap[i].szOutlookName[ValueNum], pAttr) ) {
				continue;
			}

			CComVariant Value;
			pAttr->GetValue( Value);

			// Convert the attribute value to the required LDAP type
			CComBSTR ValueStr;
			CLdapValue LdapVal;
			switch( COIDSyntaxMap::GetEntryForSyntax( m_AttribMap[i].iLdapType )->catSyntax ) {
			case CategoryString:
				CHECKTRUE( Value.vt == m_AttribMap[i].iOutlookType );
				switch( m_AttribMap[i].iOutlookType ) {
				case VT_BSTR:
					ValueStr = FixLdapSendStringSyntax( Value.bstrVal, m_AttribMap[i].iLdapType ).c_str();
					if( ValueStr.Length() > 0 ) 
						LdapVal = CLdapValue( ValueStr.m_str, m_AttribMap[i].iLdapType);
					break;
				case VT_DATE:
					ValueStr = GMTVariantTimeToLdapTime( Value.date ).c_str();
					LdapVal = CLdapValue( ValueStr.m_str, m_AttribMap[i].iLdapType);
					break;
				case VT_I4:
					LdapVal = CLdapValue((unsigned int)Value.lVal);
					break;
				default:
					CHECKTRUE(false); // Need to write conversion handler
				}
				break;
			case CategoryBinary:
			default:
				CHECKTRUE(false); // Need to write a conversion handler
			}

			if( LdapVal.Size() > 0 ) {
				CLdapAttribute oldAttr;
				if( objContact.GetAttribute( m_AttribMap[i].szLdapName, oldAttr ) )
				{
					oldAttr += LdapVal; // Add new value to existing attribute
					objContact.PutAttribute( m_AttribMap[i].szLdapName, oldAttr );
				}
				else
					objContact.PutAttribute( m_AttribMap[i].szLdapName, LdapVal );
			}
		}
	}

	// Compulsory attributes are set to dummy values if they aren't present
	CLdapAttribute attr;
	if( !objContact.GetAttribute( L"cn", attr) )
		objContact.PutAttribute( L"cn", CLdapValue(L" "));
	if( !objContact.GetAttribute( L"sn", attr) )
		objContact.PutAttribute( L"sn", CLdapValue(L" "));
	
	// Create the stripped phone number attrs
	stripPhoneNumber(objContact,L"otherTelephoneNumber",L"softStrippedNumber");
	stripPhoneNumber(objContact,L"homePhone",L"homeStrippedNumber");
	stripPhoneNumber(objContact,L"mobile",L"mobileStrippedNumber");
	stripPhoneNumber(objContact,L"telephoneNumber",L"workStrippedNumber");

	//LdapDebugBox(objContact);

	e = objContact.Add( sess );
	if( !e.Success() ) {
		wstring msg(L"Failed while sending entry to server : ");
		msg += e.Description();
		throw CError( msg );
	}
}

void CEntry::SetIndexStr( LPCOLESTR szIndex )
{
	m_sIndex = szIndex; 
}

void CEntry::SetOlIndexStr( LPCOLESTR szIndex )
{
	m_sOlIndex = szIndex; 
}

void CEntry::SetTimeStampDate( const DATE& date )
{
	m_GMTTimeStamp = date; 
}

void CEntry::SetLdapTimeStamp( const DATE& date)
{
	m_GMTLdapTimeStamp = date;
}

CEntry * CEntry::Make( const Outlook::OlDefaultFolders EntryType, int OutlookVersion )
{
	CEntry * pEntry = NULL;
	switch( EntryType) {
	case Outlook::olFolderContacts:
		pEntry = new CContact();
		break;
	}

	if( !pEntry ) return NULL;

	pEntry->m_OutlookVersion = OutlookVersion;

	return pEntry;
}

LPCOLESTR CEntry::GetEntryClass() const
{
	return m_EntryClass;
}

const DATE CEntry::LdapTimeToGMTVariantTime(const wstring &in)
{
	SYSTEMTIME GMTTime; 

	// Ldap time format is "YYYYMMDDhhmmssZ" where Z is the time zone Z=zulu
	wchar_t Zone;
	if( (swscanf( in.c_str(), L"%4hd%2hd%2hd%2hd%2hd%2hd%lc", 
			&GMTTime.wYear, &GMTTime.wMonth, &GMTTime.wDay
			, &GMTTime.wHour, &GMTTime.wMinute, &GMTTime.wSecond, &Zone )) != 7)
		throw CError( wstring(L"Unable to convert from LDAP date format \"") + in + wstring(L"\" (expecting YYYYMMDDhhmmssZ format)"));

	// OpenLDAP is sensible enough to always store its times in Zulu/GMT/UTC time
	if( Zone != L'Z' )
	{
		wchar_t ZoneStr[2] = { Zone, 0 };
		throw CError( wstring(L"Unexpected LDAP timezone \'") + wstring(ZoneStr) + wstring(L"\'"));
	}

	// " MS Dates are represented as double-precision numbers, 
	// where midnight, January 1, 1900 is 2.0, January 2, 1900
	// is 3.0, and so on. The value is passed in date. 
	// This is the same numbering system used by most spreadsheet programs,
	// although some specify incorrectly that February 29, 1900 existed, 
	// and thus set January 1, 1900 to 1.0. " !!
	// SystemTime always in GMT
	DATE out;
	CHECKTRUE( SystemTimeToVariantTime( &GMTTime, &out) == TRUE);
	return out;
}

const wstring CEntry::GMTVariantTimeToLdapTime(const DATE &in)
{
	SYSTEMTIME GMTTime; 
	CHECKTRUE( VariantTimeToSystemTime( in, &GMTTime) == TRUE);
	
	wchar_t buf[32];
	swprintf( buf, L"%04hu%02hu%02hu%02hu%02hu%02huZ", GMTTime.wYear
				, GMTTime.wMonth, GMTTime.wDay, GMTTime.wHour
				, GMTTime.wMinute, GMTTime.wSecond );
	return wstring(buf);
}

// This converts a variant time into a human readable string
// Note that the string is displayed in local time
const wstring CEntry::GMTVariantTimeToStr( const DATE& in)
{
	SYSTEMTIME Time;
	CHECKTRUE( VariantTimeToSystemTime( GMTVariantTimeToVariantTime(in), &Time) == TRUE);

	wstring ZoneName( L"unknown time zone" );
	TIME_ZONE_INFORMATION LocalTimeZone;
	switch ( GetTimeZoneInformation( &LocalTimeZone ) ) {
		case TIME_ZONE_ID_STANDARD: 
			ZoneName = LocalTimeZone.StandardName;
			break;
		case TIME_ZONE_ID_DAYLIGHT:
			ZoneName = LocalTimeZone.StandardName;
			ZoneName += L" daylight savings";
			break;
	}

	wchar_t buf[256];
	swprintf( buf, L"%02hu/%02hu/%04hu %02hu:%02hu:%02hu [%s]", Time.wDay
				, Time.wMonth, Time.wYear, Time.wHour
				, Time.wMinute, Time.wSecond, ZoneName.c_str() );

	return wstring(buf);
}

// Ghastly fn to add a specified number of minutes to a variant time
const DATE CEntry::VariantTimeAddSeconds( const DATE& in, int Secs )
{
	SYSTEMTIME Time;
	CHECKTRUE( VariantTimeToSystemTime( in, &Time) == TRUE);
	
	// This fn is unfortunately not supported on some versions of Win2000
	//CHECKTRUE( TzSpecificLocalTimeToSystemTime( &LocalTimeZone, &LocalTime, &ZuluTime ) != 0 );
	// So we have to do a tedious and error-prone conversion ourselves
	FILETIME FileTime;
	CHECKTRUE( SystemTimeToFileTime( &Time, &FileTime ) != 0 );
	
	// Filetime is a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601.
	LONGLONG* pFileTime64 = (LONGLONG *)&FileTime;

	// Now convert seconds into 100ns intervals
	LONGLONG Adjustment64 = (LONGLONG)Secs;
	Adjustment64 = Adjustment64 * 1000 * 1000 * 10;

	// Now add those seconds
	*pFileTime64 = *pFileTime64 + Adjustment64;

	// Now convert filetime back to a systemtime
	SYSTEMTIME Time2;
	CHECKTRUE( FileTimeToSystemTime( &FileTime, &Time2 ) != 0 );

	// Now convert systemtime to a Variant Time
	DATE Time3;
	CHECKTRUE( SystemTimeToVariantTime( &Time2, &Time3) == TRUE);

	return Time3;
}

// Convert time specified in local time zone to GMT
const DATE CEntry::VariantTimeToGMTVariantTime( const DATE& in)
{
	TIME_ZONE_INFORMATION LocalTimeZone;
	if( GetTimeZoneInformation( &LocalTimeZone ) == TIME_ZONE_ID_DAYLIGHT )
		return VariantTimeAddSeconds( in, ((LocalTimeZone.Bias + LocalTimeZone.DaylightBias) * 60) );
	return VariantTimeAddSeconds( in, (LocalTimeZone.Bias * 60) );
}

// Convert time specified in GMT to local timezone
const DATE CEntry::GMTVariantTimeToVariantTime( const DATE& in)
{
	TIME_ZONE_INFORMATION LocalTimeZone;
	if( GetTimeZoneInformation( &LocalTimeZone ) == TIME_ZONE_ID_DAYLIGHT )
		return VariantTimeAddSeconds( in, -((LocalTimeZone.Bias + LocalTimeZone.DaylightBias) * 60) );
	return VariantTimeAddSeconds( in, -(LocalTimeZone.Bias * 60) );
}
