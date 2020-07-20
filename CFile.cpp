#include "CFile.h"


const HANDLE CFile::hFileNull = INVALID_HANDLE_VALUE;

CFile::CFile()
	: m_bCloseOnDelete(FALSE)
	, m_strFileName(nullptr)
	, m_hFile(INVALID_HANDLE_VALUE)
{
}



