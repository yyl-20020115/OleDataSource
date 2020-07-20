#pragma once
#include "COleStreamFile.h"

class CMyOleStreamFile : public COleStreamFile
{
public:
	CMyOleStreamFile(const TCHAR* sp);
	virtual ~CMyOleStreamFile();
public:
	virtual UINT Read(void* lpBuf, UINT nCount);

private:
	TCHAR* m_srcPath;
};

