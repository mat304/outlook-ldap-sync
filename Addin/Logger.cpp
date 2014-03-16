#include "stdafx.h"
#include "logger.h"

CLogger::CLogger(const wstring LogFile)
{
	USES_CONVERSION;
	fs.open(W2CA(LogFile.c_str()),(ios::app|ios::out));
}

CLogger* CLogger::Singleton = NULL;
CLogger* CLogger::Make( const wstring LogFile)
{
	if(!Singleton) Singleton = new CLogger(LogFile);
	return Singleton;
}

void CLogger::LogString(LPCOLESTR s)
{
	USES_CONVERSION;
	if( fs.is_open() && !fs.bad() )	
	{
		fs << W2CA(s);
		fs.flush();
	}
}

