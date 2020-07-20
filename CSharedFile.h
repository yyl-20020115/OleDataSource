#pragma once
#include "CMemFile.h"


class CSharedFile : public CMemFile
{
public:
	// Constructors
	explicit CSharedFile(UINT nAllocFlags = GMEM_MOVEABLE,
		UINT nGrowBytes = 4096);

	// Attributes
	HGLOBAL Detach();
	BOOL SetHandle(HGLOBAL hGlobalMemory, BOOL bAllowGrow = TRUE);

	// Implementation
public:
	virtual ~CSharedFile();
protected:
	virtual BYTE* Alloc(SIZE_T nBytes);
	virtual BYTE* Realloc(BYTE* lpMem, SIZE_T nBytes);
	virtual void Free(BYTE* lpMem);

	UINT m_nAllocFlags;
	HGLOBAL m_hGlobalMemory;
	BOOL m_bAllowGrow;
};

