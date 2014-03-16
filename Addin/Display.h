#if !defined(AFX_IDISPLAY_H__963C1766_CC84_4A17_B81D_3915AE2F31B8__INCLUDED_)
#define AFX_IDISPLAY_H__963C1766_CC84_4A17_B81D_3915AE2F31B8__INCLUDED_

class IDisplay  
{
public:
	IDisplay() {}
	virtual ~IDisplay() {}
	virtual IDisplay& operator<<( LPCOLESTR s) = 0;
	virtual IDisplay& operator<<( const wstring& s) = 0;
};

class CDisplay : public IDisplay  
{
	IDisplay* m_RealDisplay;
	CCPLWinCriticalSection m_Crit;
public:
	CDisplay() : m_RealDisplay(NULL) {}
	virtual ~CDisplay() {}
	virtual IDisplay& operator<<( LPCOLESTR s);
	virtual IDisplay& operator<<( const wstring& s); 
	void RegisterDisplay(IDisplay* pDisp);
	void UnRegisterDisplay();
};

#endif // !defined(AFX_IDISPLAY_H__963C1766_CC84_4A17_B81D_3915AE2F31B8__INCLUDED_)
