#include "CMyOleStreamFile.h"
#include <tchar.h>
CMyOleStreamFile::CMyOleStreamFile(const TCHAR* sp)
	:m_srcPath(_tcsdup(sp))
{
}

CMyOleStreamFile::~CMyOleStreamFile()
{
	if (this->m_srcPath != nullptr) {
		free(this->m_srcPath);
	}
	this->m_srcPath = nullptr;
}

UINT CMyOleStreamFile::Read(void* lpBuf, UINT nCount)
{
	return 0;
}
