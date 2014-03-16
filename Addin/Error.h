#if !defined(AFX_ERROR_H__41167BB1_6D8C_4030_8EB5_2E458924D6B6__INCLUDED_)
#define AFX_ERROR_H__41167BB1_6D8C_4030_8EB5_2E458924D6B6__INCLUDED_

class CError  
{
	const wstring m_Desc;
public:
	CError( const wstring& desc );
	virtual ~CError();
	const wstring GetDescription();
};

#endif // !defined(AFX_ERROR_H__41167BB1_6D8C_4030_8EB5_2E458924D6B6__INCLUDED_)
