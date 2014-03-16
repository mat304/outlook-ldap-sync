#pragma once

class CMatrix
{
public:
	
	typedef enum { RTN=0, RTDO, RTDL, RTCO, RTCOC, RTCL, RTCLC, RTCMR } Rule_t;

private:
	static const int Added /*= 0*/;
	static const int Modified /*= 1*/;
	static const int Deleted /*= 2*/;
	static const int UnModified /*= 3*/;
	static const int NonExistant /*= 4*/;
	static const int MaxIndexType /*= 4*/;
	static Rule_t RuleBook[5][5];

public:

	CMatrix(void);
	~CMatrix(void);

	static LPCOLESTR ChangeTypeStr[];
	static LPCOLESTR RuleStr[];

	class IRule
	{
	public:
		virtual void Do( const CMatrix::Rule_t Rule, const int LocalOp, const int LdapOp, const wstring Value ) = 0;
	};

	void Intersection(	  const set<wstring>* LocalAdded
						, const set<wstring>* LocalModified
						, const set<wstring>* LocalDeleted
						, const set<wstring>* LocalUnModified
						, const set<wstring>* LdapAdded
						, const set<wstring>* LdapModified
						, const set<wstring>* LdapDeleted
						, const set<wstring>* LdapUnModified
						, IRule& RuleFn );

	static void CorrectIndex( const Rule_t Rule, const wstring Value, const int LdapOp, const int OlOp, set<wstring>& LdapIndex, set<wstring>& OlIndex);
	static bool WasModified( const int LdapOp, const int OlOp );
};
