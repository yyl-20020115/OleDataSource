#pragma once
#include "CFile.h"
// Memory based file implementation

class CMemFile : public CFile
{

public:
	// Constructors
	explicit CMemFile(UINT nGrowBytes = 1024);
	CMemFile(BYTE* lpBuffer, UINT nBufferSize, UINT nGrowBytes = 0);

	// Operations
	void Attach(BYTE* lpBuffer, UINT nBufferSize, UINT nGrowBytes = 0);
	BYTE* Detach();

	// Advanced Overridables
protected:
	virtual BYTE* Alloc(SIZE_T nBytes);
	virtual BYTE* Realloc(BYTE* lpMem, SIZE_T nBytes);
	virtual BYTE* Memcpy(BYTE* lpMemTarget, const BYTE* lpMemSource, SIZE_T nBytes);
	virtual void Free(BYTE* lpMem);
	virtual void GrowFile(SIZE_T dwNewLen);

	// Implementation
protected:
	SIZE_T m_nGrowBytes;
	SIZE_T m_nPosition;
	SIZE_T m_nBufferSize;
	SIZE_T m_nFileSize;
	BYTE* m_lpBuffer;
	BOOL m_bAutoDelete;

public:
	virtual ~CMemFile();

	virtual ULONGLONG GetPosition() const;
	//BOOL GetStatus(CFileStatus& rStatus) const;
	virtual ULONGLONG Seek(LONGLONG lOff, UINT nFrom);
	virtual void SetLength(ULONGLONG dwNewLen);
	virtual UINT Read(void* lpBuf, UINT nCount);
	virtual void Write(const void* lpBuf, UINT nCount);
	virtual void Abort();
	virtual void Flush();
	virtual void Close();
	virtual UINT GetBufferPtr(UINT nCommand, UINT nCount = 0,
		void** ppBufStart = NULL, void** ppBufMax = NULL);
	virtual ULONGLONG GetLength() const;

};

