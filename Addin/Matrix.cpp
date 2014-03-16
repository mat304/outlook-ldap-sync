#include "stdafx.h"
#include "matrix.h"

CMatrix::CMatrix(void)
{
}

CMatrix::~CMatrix(void)
{
}

const int CMatrix::Added = 0;
const int CMatrix::Modified = 1;
const int CMatrix::Deleted = 2;
const int CMatrix::UnModified = 3;
const int CMatrix::NonExistant = 4;
const int CMatrix::MaxIndexType = 4;

LPCOLESTR CMatrix::ChangeTypeStr[] = 
{	L"Added", 
	L"Modified", 
	L"Deleted", 
	L"UnModified", 
	L"NonExistant" 
};

CMatrix::Rule_t CMatrix::RuleBook[5][5] = {
			    /*LDAP Add*/	/* LDAP Mod*/	/*LDAP Del*/	/*LDAP UMd*/	/*LDAP NE*/
/*Local Add*/ { RTCMR,			RTCMR,			RTCOC,			RTN,			RTCOC },
/*Local Mod*/ { RTCMR,			RTCMR,			RTCOC,			RTCO,			RTCOC },
/*Local Del*/ { RTDL,			RTDL,			RTN,			RTDL,			RTN },
/*Local UMd*/ { RTN,			RTCL,			RTDO,			RTN,			RTCOC },
/*Local NE*/  { RTCLC,			RTCLC,			RTN,			RTCLC,			RTN  }
};
	
LPCOLESTR CMatrix::RuleStr[] = 
{ 
/*RTN*/		L"make no changes",
/*RTDO*/	L"delete Outlook entry",
/*RTDL*/	L"delete server entry",
/*RTCO*/	L"modify server entry", /* LDAP modify /*
/*RTCOC*/	L"create server entry", /* LDAP create /*
/*RTCL*/	L"modify Outlook entry", /* Outlook modify /*
/*RTCLC*/	L"create Outlook entry", /* Outlook create /*
/*RTCMR*/	L"copy most recent version"
};

void CMatrix::Intersection(   const set<wstring>* LocalAdded
							, const set<wstring>* LocalModified
							, const set<wstring>* LocalDeleted
							, const set<wstring>* LocalUnModified
							, const set<wstring>* LdapAdded
							, const set<wstring>* LdapModified
							, const set<wstring>* LdapDeleted
							, const set<wstring>* LdapUnModified
							, IRule& RuleFn )

{
	const set<wstring>* LocalIndex[] = { LocalAdded, LocalModified, LocalDeleted, LocalUnModified };
	const set<wstring>* LdapIndex[] = { LdapAdded, LdapModified, LdapDeleted, LdapUnModified };

	set<wstring>::const_iterator i;
	set<wstring> Intersection, Difference;

	// For each local entry row
	set<wstring> ColumnUnion[4];
	set<wstring> RowUnion;
	for( int LocalIndexType = 0; LocalIndexType < MaxIndexType; LocalIndexType++ )
	{
		RowUnion.erase( RowUnion.begin(), RowUnion.end() );

		// For each ldap entry column
		for( int LdapIndexType = 0; LdapIndexType < MaxIndexType; LdapIndexType++ )
		{

			Intersection.erase( Intersection.begin(), Intersection.end());
			set_intersection( LocalIndex[LocalIndexType]->begin()
							, LocalIndex[LocalIndexType]->end()
							, LdapIndex[LdapIndexType]->begin()
							, LdapIndex[LdapIndexType]->end()
							, inserter(Intersection, Intersection.begin()) );

			// Apply the relevant rule to syncronize these entries
			for( i = Intersection.begin(); i != Intersection.end(); ++i )
			{
				RuleFn.Do( RuleBook[LocalIndexType][LdapIndexType]
							, LocalIndexType, LdapIndexType, *i );
				RowUnion.insert(*i);
				ColumnUnion[LdapIndexType].insert(*i);
			}
		}

		// Last Column is the one where the entry doesn't exist on Ldap
		// Difference returns all entries in row that haven't already been dealt with
		Difference.erase( Difference.begin(), Difference.end());
		set_difference( LocalIndex[LocalIndexType]->begin()
					  , LocalIndex[LocalIndexType]->end()
					  , RowUnion.begin()
					  , RowUnion.end()
					  , inserter(Difference, Difference.begin()) );

		for( i = Difference.begin(); i != Difference.end(); ++i )
			RuleFn.Do( RuleBook[LocalIndexType][NonExistant]
						, LocalIndexType, NonExistant, *i );
	}
		
	// Do the columns in the last row 
	for( int LdapIndexType = 0; LdapIndexType < MaxIndexType; LdapIndexType++ )
	{
		// Difference returns all entries in column that haven't already been dealt with
		Difference.erase( Difference.begin(), Difference.end());
		set_difference(   LdapIndex[LdapIndexType]->begin()
						, LdapIndex[LdapIndexType]->end()
						, ColumnUnion[LdapIndexType].begin()
						, ColumnUnion[LdapIndexType].end()
						, inserter(Difference, Difference.begin()) );

		for( i = Difference.begin(); i != Difference.end(); ++i )
			RuleFn.Do( RuleBook[NonExistant][LdapIndexType]
						, NonExistant, LdapIndexType, *i );
	}
}

void CMatrix::CorrectIndex( const Rule_t Rule, const wstring Value, const int LdapOp, const int OlOp, set<wstring>& LdapIndex, set<wstring>& OlIndex )
{
	switch(Rule) {
		case RTDO:
		case RTDL:
			// Value has been deleted from both Outlook & LDAP
			// make sure it no longer appears in either index
			LdapIndex.erase(Value);
			OlIndex.erase(Value);
			break;
		case RTCO:
		case RTCL:
		case RTCMR:
		case RTCOC:
		case RTCLC:
			// Value now exists at both Outlook and LDAP ends
			// Make sure it appears in both indexes
			LdapIndex.insert(Value);
			OlIndex.insert(Value);
			break;
		case RTN:
			// Bring indexes up to date
			if( LdapOp == Added ) LdapIndex.insert(Value);
			if( LdapOp == Deleted ) LdapIndex.erase(Value);
			if( OlOp == Added ) OlIndex.insert(Value);
			if( OlOp == Deleted ) OlIndex.erase(Value);
			break;
	}
}

bool CMatrix::WasModified( const int LdapOp, const int OlOp )
{
	return ( (LdapOp==Modified) || (OlOp==Modified) 
			 || (LdapOp==Added) || (OlOp==Added) );
}

