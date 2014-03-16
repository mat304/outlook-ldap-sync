#pragma once

// Singleton

class CLogger;
class CLogger :
	public IDisplay
{
	ofstream fs;
	static CLogger* Singleton;

	CLogger(const wstring LogFile);
	void LogString(LPCOLESTR s);

public:

	~CLogger(void) {}

	static CLogger* Make( const wstring LogFile);

	// IDisplay functions
	virtual IDisplay& operator<<( LPCOLESTR s) { LogString(s); return *this; }
	virtual IDisplay& operator<<( const wstring& s) { LogString(s.c_str()); return *this; }
};
