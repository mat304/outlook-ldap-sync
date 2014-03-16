// Contact.cpp: implementation of the CContact class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
using namespace std;


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CContact::CContact() 
	: CEntry( ContactMap, L"IPM.Contact", L"Contact" ) 
{
}

// Map Outlook contact field names to Ldap contact attribute names
// Adds extra outlook fields while trying to maintain backwards compatibility with standard
// inetOrgPerson, Person & OrganizationalPerson schemas 
// All attributes must use standard syntaxes from RFC2252

CContact::AttribMapping CContact::ContactMap[] = 
{
	// Personal details
	{ VT_BSTR, { L"FileAs", NULL, NULL},		SyntaxDirectoryString,	L"display-name" /* OlContact */ },
	{ VT_BSTR, { L"FullName", NULL, NULL },		SyntaxDirectoryString,	L"cn" /* top */ },
	{ VT_BSTR, { L"Title",	NULL, NULL},		SyntaxDirectoryString,	L"personTitle" /* OlContact */ },
	{ VT_BSTR, { L"FirstName", NULL, NULL},		SyntaxDirectoryString,	L"givenName" /* inetOrgPerson */ },
	{ VT_BSTR, { L"MiddleName", NULL, NULL},	SyntaxDirectoryString,	L"initials" /* inetOrgPerson */ },
	{ VT_BSTR, { L"LastName", NULL, NULL},		SyntaxDirectoryString,	L"sn" /* inetOrgPerson */ },
	{ VT_BSTR, { L"Suffix",	NULL, NULL},		SyntaxDirectoryString,	L"personSuffix" /* OlContact */ },
	{ VT_BSTR, { L"Initials", NULL, NULL},		SyntaxDirectoryString,	L"personInitials" /* OlContact */ },
	{ VT_BSTR, { L"JobTitle", NULL, NULL},		SyntaxDirectoryString,	L"title" /* inetOrgPerson */ },
	{ VT_BSTR, { L"CompanyName", NULL, NULL},	SyntaxDirectoryString,	L"o" /* inetOrgPerson */ },
	{ VT_I4,   { L"Sensitivity", NULL, NULL},	SyntaxInteger,			L"sensitivity" /* OlContact */ },
	{ VT_I4,   { L"Importance", NULL, NULL},	SyntaxInteger,			L"importance" /* OlContact */ },

	// Phone numbers
	{ VT_BSTR, { L"AssistantTelephoneNumber", NULL, NULL},						SyntaxTelephoneNumber,			L"assistantTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"BusinessTelephoneNumber", L"Business2TelephoneNumber", NULL},SyntaxTelephoneNumber,			L"telephoneNumber" /* Person */ },
	{ VT_BSTR, { L"BusinessFaxNumber", NULL, NULL},								SyntaxTelephoneNumber,			L"facsimileTelephoneNumber" /* organizationalPerson */ },
	{ VT_BSTR, { L"CallbackTelephoneNumber", NULL, NULL},						SyntaxTelephoneNumber,			L"callbackTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"CarTelephoneNumber", NULL, NULL},							SyntaxTelephoneNumber,			L"carTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"CompanyMainTelephoneNumber", NULL, NULL},					SyntaxTelephoneNumber,			L"companyMainTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"HomeTelephoneNumber", L"Home2TelephoneNumber", NULL},		SyntaxTelephoneNumber,			L"homePhone" /* inetOrgPerson */ },
	{ VT_BSTR, { L"HomeFaxNumber", NULL, NULL},									SyntaxTelephoneNumber,			L"otherFacsimileTelephoneNumber" /* ?? */ },
	{ VT_BSTR, { L"ISDNNumber", NULL, NULL},									SyntaxTelephoneNumber,			L"internationaliSDNNumber" /* organizationalPerson */ },
	{ VT_BSTR, { L"MobileTelephoneNumber", NULL, NULL},							SyntaxTelephoneNumber,			L"mobile" /* inetOrgPerson */ },
	{ VT_BSTR, { L"OtherTelephoneNumber", NULL, NULL},							SyntaxTelephoneNumber,			L"otherTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"OtherFaxNumber", NULL, NULL},								SyntaxTelephoneNumber,			L"otherFaxNumber" /* OlContact */ },
	{ VT_BSTR, { L"PagerNumber", NULL, NULL},									SyntaxTelephoneNumber,			L"pager" /* inetOrgPerson */ },
	{ VT_BSTR, { L"PrimaryTelephoneNumber", NULL, NULL},						SyntaxTelephoneNumber,			L"primaryTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"RadioTelephoneNumber", NULL, NULL},							SyntaxTelephoneNumber,			L"radioTelephoneNumber" /* OlContact */ },
	{ VT_BSTR, { L"TelexNumber", NULL, NULL},									SyntaxTelephoneNumber,			L"telexNumber" /* organizationalPerson */ },
	{ VT_BSTR, { L"TTYTDDTelephoneNumber", NULL, NULL},							SyntaxTelephoneNumber,			L"TTYTDDTelephoneNumber" /* OlContact */ },

	// Addresses
	{ VT_BSTR, { L"BusinessAddress", NULL, NULL},			SyntaxPostalAddress,	L"postalAddress" /* organizationalPerson */ },
	{ VT_BSTR, { L"BusinessAddressStreet", NULL, NULL},		SyntaxDirectoryString,		L"street" /* organizationalPerson */ },
	{ VT_BSTR, { L"BusinessAddressCity", NULL, NULL},		SyntaxDirectoryString,		L"l" /* organizationalPerson */ },
	{ VT_BSTR, { L"BusinessAddressState", NULL, NULL},		SyntaxDirectoryString,		L"st" /* organizationalPerson */ },
	{ VT_BSTR, { L"BusinessAddressCountry", NULL, NULL},	SyntaxDirectoryString,		L"c" /* ?? */ },
	{ VT_BSTR, { L"BusinessAddressPostalCode", NULL, NULL},	SyntaxDirectoryString,		L"postalCode" /* organizationalPerson */ },
	{ VT_BSTR, { L"HomeAddress", NULL, NULL},				SyntaxPostalAddress,	L"homePostalAddress" /* inetOrgPerson */ },
	{ VT_BSTR, { L"HomeAddressStreet", NULL, NULL},			SyntaxDirectoryString,		L"homeAddressStreet" /* OlContact */ },
	{ VT_BSTR, { L"HomeAddressCity", NULL, NULL},			SyntaxDirectoryString,		L"homeAddressCity" /* OlContact */ },
	{ VT_BSTR, { L"HomeAddressState", NULL, NULL},			SyntaxDirectoryString,		L"homeAddressState" /* OlContact */ },
	{ VT_BSTR, { L"HomeAddressCountry", NULL, NULL},		SyntaxDirectoryString,		L"homeAddressCountry" /* OlContact */ },
	{ VT_BSTR, { L"HomeAddressPostalCode", NULL, NULL},		SyntaxDirectoryString,		L"homeAddressPostalCode" /* OlContact */ },
	{ VT_BSTR, { L"OtherAddress", NULL, NULL},				SyntaxPostalAddress,	L"registeredAddress" /* inetOrgPerson */ },
	{ VT_BSTR, { L"OtherAddressStreet", NULL, NULL},		SyntaxDirectoryString,		L"otherAddressStreet" /* OlContact */ },
	{ VT_BSTR, { L"OtherAddressCity", NULL, NULL},			SyntaxDirectoryString,		L"otherAddressCity" /* OlContact */ },
	{ VT_BSTR, { L"OtherAddressState", NULL, NULL},			SyntaxDirectoryString,		L"otherAddressState" /* OlContact */ },
	{ VT_BSTR, { L"OtherAddressCountry", NULL, NULL},		SyntaxDirectoryString,		L"otherAddressCountry" /* OlContact */ },
	{ VT_BSTR, { L"OtherAddressPostalCode", NULL, NULL},	SyntaxDirectoryString,		L"otherAddressPostalCode" /* OlContact */ },
	{ VT_I4,   { L"SelectedMailingAddress", NULL, NULL},	SyntaxInteger,			L"mailingAddressType" /* OlContact */ },

	// Email addresses
	{ VT_BSTR, { L"Email1Address", L"Email2Address", L"Email3Address"},				SyntaxIA5String, L"mail" /* inetOrgPerson */ },
	{ VT_BSTR, { L"Email1DisplayName", L"Email2DisplayName", L"Email3DisplayName"}, SyntaxIA5String, L"mailDisplayName" /* OlContact */ },
	
	// Webpage
	{ VT_BSTR, { L"WebPage", NULL, NULL},				SyntaxDirectoryString, L"URL" /* OlContact */ },
	{ VT_BSTR, { L"IMAddress", NULL, NULL},				SyntaxDirectoryString, L"IMAddress" /* OlContact */ },

	// Categories
	{ VT_BSTR, { L"Categories", NULL, NULL},			SyntaxDirectoryString, L"categories" /* OlContact */ },

	// Office
	{ VT_BSTR, { L"Department", NULL, NULL},			SyntaxDirectoryString,	L"ou" /* organizationalPerson */ },
	{ VT_BSTR, { L"OfficeLocation", NULL, NULL},		SyntaxDirectoryString,	L"physicalDeliveryOfficeName" /* ?? */ },
	{ VT_BSTR, { L"Profession", NULL, NULL},			SyntaxDirectoryString,		L"profession" /* OlContact */ },
	{ VT_BSTR, { L"ManagerName", NULL, NULL},			SyntaxDirectoryString,		L"managerName" /* OlContact */ },
	{ VT_BSTR, { L"AssistantName", NULL, NULL},			SyntaxDirectoryString,		L"assistantName" /* OlContact */ },

	// Personal 
	{ VT_BSTR, { L"NickName", NULL, NULL},				SyntaxDirectoryString,		L"displayName" /* OlContact */ },
	{ VT_BSTR, { L"Spouse", NULL, NULL},				SyntaxDirectoryString,		L"spouse" /* OlContact */ },
	{ VT_DATE, { L"Birthday", NULL, NULL},				SyntaxGeneralizedTime,	L"birthday" /* OlContact */ },
	{ VT_DATE, { L"Anniversary", NULL, NULL},			SyntaxGeneralizedTime,	L"anniversary" /* OlContact */ },

	// NetMeeting
	{ VT_BSTR, { L"NetMeetingServer", NULL, NULL},		SyntaxDirectoryString,	L"conferenceInformation" /* OlContact */ },
	{ VT_BSTR, { NULL/*L"NetMeetingAlias"*/, NULL, NULL},		SyntaxDirectoryString,	L"netMeetingAlias" /* OlContact */ },

	// Internet free busy
	{ VT_BSTR, { L"InternetFreeBusyAddress", NULL, NULL},	SyntaxDirectoryString,	L"internetFreeBusyAddress" /* OlContact */ },

	// Text notes
	{ VT_BSTR, { L"Body", NULL, NULL},					SyntaxDirectoryString,	L"comment"/* OlContact */ },
};
int CContact::MaxLdapAttribMap()
{
	return sizeof(ContactMap)/sizeof(AttribMapping);
}

// These ones are read from MAPI to avoid setting off the Outlook
// anti-virus security warning
CContact::MapiMapping CContact::ContactMapiMap[] = 
{
	{ L"Email1Address", CMAPIMessage::MP_EMAIL1_ADDRESS },
	{ L"Email2Address", CMAPIMessage::MP_EMAIL2_ADDRESS },
	{ L"Email3Address", CMAPIMessage::MP_EMAIL3_ADDRESS },
	{ L"Email1DisplayName", CMAPIMessage::MP_EMAIL1_DISPLAY_NAME },
	{ L"Email2DisplayName", CMAPIMessage::MP_EMAIL2_DISPLAY_NAME },
	{ L"Email3DisplayName", CMAPIMessage::MP_EMAIL3_DISPLAY_NAME },
	{ L"NetMeetingAlias", CMAPIMessage::MP_NETMEETING_ALIAS },
	{ L"IMAddress", CMAPIMessage::MP_IMADDRESS },
	{ L"Body", CMAPIMessage::MP_BODY }
};
int CContact::MaxMapiMapping() const 
{ 
	return sizeof(ContactMapiMap)/sizeof(MapiMapping); 
}

#ifdef MSO9_COMPATIBLE

// Truely horrible stuff that isn't required for MSO>=10
// Trying to keep horribleness in one place
#define ContactBSTRValue( A, B ) \
	if(Name==L"FullName") hr=spContact->##A##_FullName(##B##); \
	else if(Name==L"Title") hr=spContact->##A##_Title(##B##); \
	else if(Name==L"FirstName") hr=spContact->##A##_FirstName(##B##); \
	else if(Name==L"MiddleName") hr=spContact->##A##_MiddleName(##B##); \
	else if(Name==L"LastName") hr=spContact->##A##_LastName(##B##); \
	else if(Name==L"Suffix") hr=spContact->##A##_Suffix(##B##); \
	else if(Name==L"Initials") hr=spContact->##A##_Initials(##B##); \
	else if(Name==L"JobTitle") hr=spContact->##A##_JobTitle(##B##); \
	else if(Name==L"CompanyName") hr=spContact->##A##_CompanyName(##B##); \
	else if(Name==L"FileAs") hr=spContact->##A##_FileAs(##B##); \
	else if(Name==L"AssistantTelephoneNumber") hr=spContact->##A##_AssistantTelephoneNumber(##B##); \
	else if(Name==L"BusinessTelephoneNumber") hr=spContact->##A##_BusinessTelephoneNumber(##B##); \
	else if(Name==L"Business2TelephoneNumber") hr=spContact->##A##_Business2TelephoneNumber(##B##); \
	else if(Name==L"BusinessFaxNumber") hr=spContact->##A##_BusinessFaxNumber(##B##); \
	else if(Name==L"CallbackTelephoneNumber") hr=spContact->##A##_CallbackTelephoneNumber(##B##); \
	else if(Name==L"CarTelephoneNumber") hr=spContact->##A##_CarTelephoneNumber(##B##); \
	else if(Name==L"CompanyMainTelephoneNumber") hr=spContact->##A##_CompanyMainTelephoneNumber(##B##); \
	else if(Name==L"HomeTelephoneNumber") hr=spContact->##A##_HomeTelephoneNumber(##B##); \
	else if(Name==L"Home2TelephoneNumber") hr=spContact->##A##_Home2TelephoneNumber(##B##); \
	else if(Name==L"HomeFaxNumber") hr=spContact->##A##_HomeFaxNumber(##B##); \
	else if(Name==L"ISDNNumber") hr=spContact->##A##_ISDNNumber(##B##); \
	else if(Name==L"MobileTelephoneNumber") hr=spContact->##A##_MobileTelephoneNumber(##B##); \
	else if(Name==L"OtherTelephoneNumber") hr=spContact->##A##_OtherTelephoneNumber(##B##); \
	else if(Name==L"OtherFaxNumber") hr=spContact->##A##_OtherFaxNumber(##B##); \
	else if(Name==L"PagerNumber") hr=spContact->##A##_PagerNumber(##B##); \
	else if(Name==L"PrimaryTelephoneNumber") hr=spContact->##A##_PrimaryTelephoneNumber(##B##); \
	else if(Name==L"RadioTelephoneNumber") hr=spContact->##A##_RadioTelephoneNumber(##B##); \
	else if(Name==L"TelexNumber") hr=spContact->##A##_TelexNumber(##B##); \
	else if(Name==L"TTYTDDTelephoneNumber") hr=spContact->##A##_TTYTDDTelephoneNumber(##B##); \
	else if(Name==L"BusinessAddress") hr=spContact->##A##_BusinessAddress(##B##); \
	else if(Name==L"BusinessAddressStreet") hr=spContact->##A##_BusinessAddressStreet(##B##); \
	else if(Name==L"BusinessAddressCity") hr=spContact->##A##_BusinessAddressCity(##B##); \
	else if(Name==L"BusinessAddressState") hr=spContact->##A##_BusinessAddressState(##B##); \
	else if(Name==L"BusinessAddressCountry") hr=spContact->##A##_BusinessAddressCountry(##B##); \
	else if(Name==L"BusinessAddressPostalCode") hr=spContact->##A##_BusinessAddressPostalCode(##B##); \
	else if(Name==L"HomeAddress") hr=spContact->##A##_HomeAddress(##B##); \
	else if(Name==L"HomeAddressStreet") hr=spContact->##A##_HomeAddressStreet(##B##); \
	else if(Name==L"HomeAddressCity") hr=spContact->##A##_HomeAddressCity(##B##); \
	else if(Name==L"HomeAddressState") hr=spContact->##A##_HomeAddressState(##B##); \
	else if(Name==L"HomeAddressCountry") hr=spContact->##A##_HomeAddressCountry(##B##); \
	else if(Name==L"HomeAddressPostalCode") hr=spContact->##A##_HomeAddressPostalCode(##B##); \
	else if(Name==L"OtherAddress") hr=spContact->##A##_OtherAddress(##B##); \
	else if(Name==L"OtherAddressStreet") hr=spContact->##A##_OtherAddressStreet(##B##); \
	else if(Name==L"OtherAddressCity") hr=spContact->##A##_OtherAddressCity(##B##); \
	else if(Name==L"OtherAddressState") hr=spContact->##A##_OtherAddressState(##B##); \
	else if(Name==L"OtherAddressCountry") hr=spContact->##A##_OtherAddressCountry(##B##); \
	else if(Name==L"OtherAddressPostalCode") hr=spContact->##A##_OtherAddressPostalCode(##B##); \
	else if(Name==L"Email1Address") hr=spContact->##A##_Email1Address(##B##); \
	else if(Name==L"Email2Address") hr=spContact->##A##_Email2Address(##B##); \
	else if(Name==L"Email3Address") hr=spContact->##A##_Email3Address(##B##); \
	else if(Name==L"WebPage") hr=spContact->##A##_WebPage(##B##); \
	else if(Name==L"Categories") hr=spContact->##A##_Categories(##B##); \
	else if(Name==L"Department") hr=spContact->##A##_Department(##B##); \
	else if(Name==L"OfficeLocation") hr=spContact->##A##_OfficeLocation(##B##); \
	else if(Name==L"Profession") hr=spContact->##A##_Profession(##B##); \
	else if(Name==L"ManagerName") hr=spContact->##A##_ManagerName(##B##); \
	else if(Name==L"AssistantName") hr=spContact->##A##_AssistantName(##B##); \
	else if(Name==L"NickName") hr=spContact->##A##_NickName(##B##); \
	else if(Name==L"Spouse") hr=spContact->##A##_Spouse(##B##); \
	else if(Name==L"NetMeetingServer") hr=spContact->##A##_NetMeetingServer(##B##); \
	else if(Name==L"NetMeetingAlias") hr=spContact->##A##_NetMeetingAlias(##B##); \
	else if(Name==L"InternetFreeBusyAddress") hr=spContact->##A##_InternetFreeBusyAddress(##B##); \
	else if(Name==L"Body") hr=spContact->##A##_Body(##B##); \
	else hr=-1; // ie FAILED

#define ContactDATEValue( A, B ) \
	if(Name==L"Birthday") hr=spContact->##A##_Birthday(##B##); \
	else if(Name==L"Anniversary") hr=spContact->##A##_Anniversary(##B##); \
	else hr=-1; // ie FAILED

HRESULT CContact::oldGetContactValue( CComQIPtr<Outlook::_ContactItem>& spContact, const wstring& Name, CComVariant& Value)
{
	HRESULT hr;
	CComBSTR bstr = NULL;
	switch(Value.vt) {
		case VT_BSTR:
			ContactBSTRValue( get, &bstr )
			if(Name==L"IMAddress") hr=(m_OutlookVersion>=10)?spContact->get_IMAddress(&bstr):0;
			else if(Name==L"Email1DisplayName") hr=(m_OutlookVersion>=10)?spContact->get_Email1DisplayName(&bstr):0;
			else if(Name==L"Email2DisplayName") hr=(m_OutlookVersion>=10)?spContact->get_Email2DisplayName(&bstr):0;
			else if(Name==L"Email3DisplayName") hr=(m_OutlookVersion>=10)?spContact->get_Email3DisplayName(&bstr):0;
			if(bstr==NULL) Value=L"";
			else Value = bstr;
			break;
		case VT_I4:
			if(Name==L"SelectedMailingAddress") {
				Outlook::OlMailingAddress MailingAddressType;
				hr=spContact->get_SelectedMailingAddress( &MailingAddressType );
				if( SUCCEEDED(hr) ) Value = MailingAddressType;
				else Value = 0; // No selected mailing address
			} else if(Name==L"Sensitivity") {
				Outlook::OlSensitivity Sen;
				hr=spContact->get_Sensitivity(&Sen);
				if( SUCCEEDED(hr) ) Value = Sen;
				else Value = 0;
			} else if(Name==L"Importance") {
				Outlook::OlImportance Importance;
				hr=spContact->get_Importance(&Importance);
				if( SUCCEEDED(hr) ) Value = Importance;
				else Value = 0;
			}
			break;
		case VT_DATE:
			ContactDATEValue( get, &Value.date )
			break;
		default:
			hr = -1; // ie FAILED
	}
	return hr;
}

HRESULT CContact::oldPutContactValue( CComQIPtr<Outlook::_ContactItem>& spContact, const wstring& Name, CComVariant& Value)
{
	HRESULT hr;
	switch(Value.vt) {
		case VT_BSTR:
			ContactBSTRValue( put, Value.bstrVal )
			if(Name==L"IMAddress") hr=(m_OutlookVersion>=10)?spContact->put_IMAddress(Value.bstrVal):0;
			else if(Name==L"Email1DisplayName") hr=(m_OutlookVersion>=10)?spContact->put_Email1DisplayName(Value.bstrVal):0;
			else if(Name==L"Email2DisplayName") hr=(m_OutlookVersion>=10)?spContact->put_Email2DisplayName(Value.bstrVal):0;
			else if(Name==L"Email3DisplayName") hr=(m_OutlookVersion>=10)?spContact->put_Email3DisplayName(Value.bstrVal):0;
			break;
		case VT_I4:
			if(Name==L"SelectedMailingAddress")
				hr=spContact->put_SelectedMailingAddress( (Outlook::OlMailingAddress)Value.lVal );
			else if(Name==L"Sensitivity")
				hr=spContact->put_Sensitivity((Outlook::OlSensitivity)Value.lVal);
			else if(Name==L"Importance") 
				hr=spContact->put_Importance((Outlook::OlImportance)Value.lVal);
			break;
		case VT_DATE:
			ContactDATEValue( put, Value.date )
			break;
		default:
			hr = -1; // ie FAILED
	}
	return hr;
}

#endif

void CContact::Pack( IDispatch * pDisp, CMAPIFolder * pMAPIFolder ) 
{
	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);

#ifndef MSO9_COMPATIBLE
	CComPtr<Outlook::ItemProperties> Properties;
	spContact->get_ItemProperties( &Properties );
#endif

	FlushAllAttributes();

	// Get MAPI interface to message so we can use unrestricted MAPI calls
	// to get around those annoying Outlook security restrictions
	auto_ptr<CMAPIMessage> MAPIMsg (new CMAPIMessage(pMAPIFolder));
	CMAPIEntryID MAPIEntryID;
	
	// Outlook index must be set by ScanLocal
	CHECKTRUE( GetOlIndex() != L"" );
	MAPIEntryID.Set(GetOlIndex());

	CHECKHR( MAPIMsg->OpenMessage( MAPIEntryID));

//	MAPIMsg->DebugDisplayAllProps();

	for( int i=0; i<MaxLdapAttribMap(); i++) 
	{
		for( int ValueNum = 0; ValueNum < MAX_LDAP_VALUES; ValueNum++ )
		{
			if( !m_AttribMap[i].szOutlookName[ValueNum] ) 
				continue;
			
			const wstring PropertyName(m_AttribMap[i].szOutlookName[ValueNum]);
			CComVariant Value;

			// See if this attr needs to be read from MAPI or from Outlook
			const MapiMapping * MapiAttrPtr = NULL;
			for( int j = 0; j<MaxMapiMapping(); j++ )
				if( ContactMapiMap[j].szOutlookName == PropertyName )
				{
					MapiAttrPtr = &ContactMapiMap[j];
					break;
				}

			if( MapiAttrPtr )
			{
				// Reading field directly from MAPI interface
				// .. to avoid security restrictions

				// Outlook < 10 doesn't support email display names 
				if( (m_OutlookVersion<10) && ( (PropertyName==L"Email1DisplayName")
					|| (PropertyName==L"Email2DisplayName")
					|| (PropertyName==L"Email3DisplayName") ) )
					continue;

				if( FAILED( MAPIMsg->GetProperty( MapiAttrPtr->eMAPIProp, Value ) ) )
					throw CError( L"Failed to MAPI read Outlook property: " + PropertyName );
			} 
			else 
			{
				// Reading field from Outlook

	#ifdef MSO9_COMPATIBLE
				CHECKHR(Value.ChangeType(m_AttribMap[i].iOutlookType));
				if( FAILED( oldGetContactValue( spContact, PropertyName, Value)))
	#else
				CComPtr<Outlook::ItemProperty> Property;
				if( FAILED( Properties->Item( CComVariant( PropertyName.c_str() ), &Property )) )
					throw CError( L"Failed to find Outlook property: " + PropertyName );

				if( FAILED( Property->get_Value( &Value ) ) )
	#endif
					throw CError( L"Failed to read Outlook property: " + PropertyName );

				// Empty/Clear field implies non-existant LDAP attribute
				if( Value.vt == VT_EMPTY ) 
					continue;

				if( Value.vt != m_AttribMap[i].iOutlookType )
					throw CError( L"Got unexpected type for Outlook property " + PropertyName );

				// Special date value denotes that date is not set
				if( (Value.vt == VT_DATE) && (Value.date == 949998.0 ) )
					continue;

				// All internal times stored in GMT
				// All outlook times stored in local timezone
				// Exceptions to the rule are birthday and anniversary
				// which are independant of timezone
				if( (Value.vt == VT_DATE) 
					&& !(( PropertyName == L"Birthday" )
						|| ( PropertyName == L"Anniversary" )) ) 
				{
					Value = VariantTimeToGMTVariantTime( Value.date );
					CHECKHR(Value.ChangeType(VT_DATE));
				}
			}

			CAttribute * attr = new CAttribute(Value);
			PutAttribute( PropertyName.c_str(), attr );

		}
	}
	
#if 0
	// Now check this contact's attachments
	CComPtr<Outlook::Attachments> Attachments;
	long NoofAttachments;
	CHECKHR(spContact->get_Attachments(&Attachments));
	CHECKHR(Attachments->get_Count(&NoofAttachments));
	for( int i = 1; i <= NoofAttachments; ++i )
	{
		// Get the MAPI IAttach interface for this attachment
		CComPtr<IUnknown> pIAttach;
		CComPtr<Outlook::Attachment> Attachment;
		CHECKHR(Attachments->Item(CComVariant(i),&Attachment));
		CHECKHR(Attachment->get_MAPIOBJECT(&pIAttach));
		auto_ptr<CMAPIAttachment> pMAPIAttachment (new CMAPIAttachment(pIAttach));

		long RenderingPosition = pMAPIAttachment->GetRenderingPosition();
		CMAPIAttachment::t_AttachMethod AttachMethod = pMAPIAttachment->GetAttachMethod();
		if( pMAPIAttachment->IsFile() )	pMAPIAttachment->GetFile((LPTSTR)"c:\\jm");
	}
#endif
}

bool CContact::Validate()
{
	const CAttribute * pAttr;

	// FileAs is compulsory in Outlook
	// Evolution doesn't create FileAs so we sometimes need to generate a default value
	if( !FindAttribute( L"FileAs", pAttr ) ) {

		wstring FirstName(L""), LastName(L""), FileAs;

		if( FindAttribute( L"FirstName", pAttr ) )
			FirstName = pAttr->GetValueStr();
		if( FindAttribute( L"LastName", pAttr ) )
			LastName = pAttr->GetValueStr();

		if( (FirstName != L"") && (LastName == L"") )
			FileAs = FirstName;
		else if ( (FirstName == L"") && (LastName != L"") )
			FileAs = LastName;
		else if ( (FirstName != L"") && (LastName != L"")  )
			FileAs = LastName + L", " + FirstName;
		else 
			FileAs = GetOlIndex();

		CComVariant ComValue((LPCOLESTR)FileAs.c_str());
		pAttr = new CAttribute(ComValue);
		PutAttribute( L"FileAs", pAttr );
	}

	return true;
}

void CContact::UnPack( IDispatch * pDisp, const bool Create )
{

	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);

#ifndef MSO9_COMPATIBLE
	CComPtr<Outlook::ItemProperties> Properties;
	spContact->get_ItemProperties( &Properties );
#endif

	for( int i=0; i<MaxLdapAttribMap(); i++) 
	{
		for( int ValueNum = 0; ValueNum<MAX_LDAP_VALUES; ValueNum++ )
		{

			// Check we want to use this Outlook field
			if( ! m_AttribMap[i].szOutlookName[ValueNum] )
				continue;

			const wstring PropertyName = m_AttribMap[i].szOutlookName[ValueNum];

#ifndef MSO9_COMPATIBLE
			CComPtr<Outlook::ItemProperty> Property;
			if( FAILED( Properties->Item( CComVariant( PropertyName.c_str() ), &Property )) )
				throw CError( L"Failed to find Outlook property: " + PropertyName );
#endif

			const CAttribute * pValue; 
			CComVariant var; 
			// Non-existant LDAP attribute implies empty/cleared Outlook field
			if( ! FindAttribute( PropertyName.c_str(), pValue ) )
			{
				// "Not set" value depends on type 
				var.ChangeType(m_AttribMap[i].iOutlookType);
				switch(var.vt ) {
				case VT_BSTR: var = L""; break;
				case VT_DATE: var.date = 949998.0; break;
				}
			}
			else
				pValue->GetValue(var); 

			// All internal times stored in GMT
			// All outlook times stored in local timezone
			if( (var.vt == VT_DATE)
				&& !(( PropertyName == L"Birthday" )
						|| ( PropertyName == L"Anniversary" )) )
			{	
				var = GMTVariantTimeToVariantTime( var.date );
				CHECKHR(var.ChangeType(VT_DATE));
			}

#ifdef MSO9_COMPATIBLE
			if( FAILED( oldPutContactValue( spContact, PropertyName, var)))
#else
			if( FAILED( Property->put_Value(var) ) )
#endif
				throw CError( wstring(L"Failed to write Outlook property: ") + wstring(PropertyName));
		}
	}

	// Annoyingly the newly added entry has a different (read-only) unique EntryID
	// We must save the LDAP EntryID so that we can match entries in future syncronizations
	if( Create ) 
		CHECKHR(spContact->put_User1( CComBSTR(GetIndex().c_str())));

	// Remember to save the changes
	if( FAILED(spContact->Save()))
		throw CError(L"Outlook failed to save this contact");

	if( Create )
		SetIndex(pDisp);
}

bool CContact::CheckOlPtr( IDispatch * pDisp )
{
	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);
	return (spContact!=NULL);
}

// Given an Outlook pointer to a contact object set our entry's index accordingly
void CContact::SetIndex( IDispatch * pDisp )
{
	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);

	// Always store the outlook entry index
	// It is essential to use this for any outlook calls
	CComBSTR EntryID = NULL;
	CHECKHR( spContact->get_EntryID( &EntryID));
	SetOlIndexStr( EntryID.m_str );

	// Entries that we have created ourselves (from LDAP entries)
	// will have (another!) index set in the User1 field. This index will match
	// the LDAP entry index and should be used for everything except outlook calls.
	// (See comment in Unpack() above)
	CComBSTR bstr;
	wstring User1;
	CHECKHR(spContact->get_User1( &bstr));
	if(bstr) User1 = wstring(bstr.m_str);
	else User1 = wstring(L"");

	if( User1 != L"" )
	{
		if(User1.find(wstring(L"guid="))==string::npos)
			User1 = L"guid="+User1;
		SetIndexStr( User1.c_str() );
		return;
	}

	// Build an index from the Outlook EntryID field
	wstring idx(EntryID.m_str);
	idx = L"guid="+idx;
	SetIndexStr( idx.c_str() );
}

void CContact::SetTimeStamp( IDispatch * pDisp )
{
	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);

	DATE olDate;
	CHECKHR( spContact->get_LastModificationTime( &olDate ));
	SetTimeStampDate(VariantTimeToGMTVariantTime(olDate));
}

void CContact::SetEntryName( IDispatch* pDisp)
{
	CComQIPtr<Outlook::_ContactItem> spContact(pDisp);

	CComBSTR Value;
	CHECKHR( spContact->get_FullName( &Value ) );
	if( Value == NULL ) Value = L"";
	SetEntryNameStr( Value.m_str);
}

void CContact::GetLdapObjectClass(CLdapAttribute &rv)
{
	rv += CLdapValue( L"person" ); 
	rv += CLdapValue( L"organizationalPerson" );
	rv += CLdapValue( L"inetOrgPerson" );
	rv += CLdapValue( L"OlEntry" );
	rv += CLdapValue( L"OlContact" );
}

CLdapString CContact::GetLdapFilter()
{
	CLdapString szContactID = GetIndex().c_str();
	CLdapString szFilter(L"(&(objectClass=inetOrgPerson)(cn=");
	szFilter += szContactID + L")";
	return szFilter;
}

