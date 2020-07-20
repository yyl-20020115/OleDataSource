#include "CSharedFile.h"

CSharedFile::CSharedFile(UINT nAllocFlags, UINT nGrowBytes)
	: CMemFile(nGrowBytes)
	, m_nAllocFlags(nAllocFlags)
	, m_bAllowGrow(TRUE)
	, m_hGlobalMemory(INVALID_HANDLE_VALUE)
{
}

CSharedFile::~CSharedFile()
{
	if (m_lpBuffer != nullptr) {
		Close();        // call appropriate Close/Free
	}
	this->m_lpBuffer = nullptr;
}

BOOL CSharedFile::SetHandle(HGLOBAL hGlobalMemory, BOOL bAllowGrow)
{
	if (hGlobalMemory == NULL)
	{
		return FALSE;
	}

	this->m_hGlobalMemory = hGlobalMemory;
	this->m_lpBuffer 
		= (BYTE*)::GlobalLock(this->m_hGlobalMemory);
	this->m_nBufferSize 
		= this->m_nFileSize 
		= ::GlobalSize(this->m_hGlobalMemory);
	this->m_bAllowGrow = bAllowGrow;
	return TRUE;
}

BYTE* CSharedFile::Alloc(SIZE_T nBytes)
{
	this->m_hGlobalMemory 
		= ::GlobalAlloc(m_nAllocFlags, nBytes);
	if (this->m_hGlobalMemory == nullptr)
		return nullptr;
	else
		return (BYTE*)::GlobalLock(this->m_hGlobalMemory);
}

BYTE* CSharedFile::Realloc(BYTE*, SIZE_T nBytes)
{
	if (!this->m_bAllowGrow)
		return NULL;

	::GlobalUnlock(this->m_hGlobalMemory);
	HGLOBAL hNew = ::GlobalReAlloc(this->m_hGlobalMemory, nBytes, m_nAllocFlags);
	if (hNew == nullptr)
		return nullptr;
	this->m_hGlobalMemory = hNew;
	return (BYTE*)::GlobalLock(this->m_hGlobalMemory);
}

void CSharedFile::Free(BYTE*)
{
	::GlobalUnlock(this->m_hGlobalMemory);
	::GlobalFree(this->m_hGlobalMemory);
}

HGLOBAL CSharedFile::Detach()
{
	HGLOBAL hMem = this->m_hGlobalMemory;

	::GlobalUnlock(this->m_hGlobalMemory);
	this->m_hGlobalMemory = nullptr; // detach from global handle

	// re-initialize the CMemFile parts too
	this->m_lpBuffer = nullptr;
	this->m_nBufferSize = 0;

	return hMem;
}
