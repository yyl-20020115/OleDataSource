#include "CMemFile.h"

CMemFile::CMemFile(UINT nGrowBytes)
	: m_nGrowBytes(nGrowBytes)
	, m_nPosition()
	, m_nBufferSize()
	, m_nFileSize()
	, m_lpBuffer()
	, m_bAutoDelete(TRUE)
{
}

CMemFile::CMemFile(BYTE* lpBuffer, UINT nBufferSize, UINT nGrowBytes)
	: m_nGrowBytes(nGrowBytes)
	, m_nPosition()
	, m_nBufferSize(nBufferSize)
	, m_nFileSize(nGrowBytes == 0 ? nBufferSize : 0)
	, m_lpBuffer(lpBuffer)
	, m_bAutoDelete(FALSE)
{
}

void CMemFile::Attach(BYTE* lpBuffer, UINT nBufferSize, UINT nGrowBytes)
{
	this->m_nGrowBytes = nGrowBytes;
	this->m_nPosition = 0;
	this->m_nBufferSize = nBufferSize;
	this->m_nFileSize = nGrowBytes == 0 ? nBufferSize : 0;
	this->m_lpBuffer = lpBuffer;
	this->m_bAutoDelete = FALSE;
}

BYTE* CMemFile::Detach()
{
	BYTE* lpBuffer = this->m_lpBuffer;
	this->m_lpBuffer = nullptr;
	this->m_nFileSize = 0;
	this->m_nBufferSize = 0;
	this->m_nPosition = 0;

	return lpBuffer;
}

CMemFile::~CMemFile()
{
	// Close should have already been called, but we check anyway
	if (this->m_lpBuffer!=nullptr)
	{
		this->Close();
	}
	this->m_nGrowBytes = 0;
	this->m_nPosition = 0;
	this->m_nBufferSize = 0;
	this->m_nFileSize = 0;
}

BYTE* CMemFile::Alloc(SIZE_T nBytes)
{
	return (BYTE*)malloc(nBytes);
}

BYTE* CMemFile::Realloc(BYTE* lpMem, SIZE_T nBytes)
{
	return (BYTE*)realloc(lpMem, nBytes);
}

BYTE* CMemFile::Memcpy(BYTE* lpMemTarget, const BYTE* lpMemSource,
	SIZE_T nBytes)
{
	memcpy_s(lpMemTarget, nBytes, lpMemSource, nBytes);
	return lpMemTarget;
}

void CMemFile::Free(BYTE* lpMem)
{
	free(lpMem);
}

ULONGLONG CMemFile::GetPosition() const
{
	return this->m_nPosition;
}

void CMemFile::GrowFile(SIZE_T dwNewLen)
{
	if (dwNewLen > this->m_nBufferSize)
	{
		// grow the buffer
		SIZE_T dwNewBufferSize = this->m_nBufferSize;

		// watch out for buffers which cannot be grown!
		// determine new buffer size
		while (dwNewBufferSize < dwNewLen)
			dwNewBufferSize += this->m_nGrowBytes;

		// allocate new buffer
		BYTE* lpNew = nullptr;
		if (m_lpBuffer == NULL)
			lpNew = Alloc(dwNewBufferSize);
		else
			lpNew = Realloc(this->m_lpBuffer, dwNewBufferSize);

		this->m_lpBuffer = lpNew;
		this->m_nBufferSize = dwNewBufferSize;
	}
}

ULONGLONG CMemFile::GetLength() const
{
	return this->m_nFileSize;
}

void CMemFile::SetLength(ULONGLONG dwNewLen)
{
	if (dwNewLen > this->m_nBufferSize)
		this->GrowFile((SIZE_T)dwNewLen);

	if (dwNewLen < this->m_nPosition)
		this->m_nPosition = (SIZE_T)dwNewLen;

	this->m_nFileSize = (SIZE_T)dwNewLen;
}

UINT CMemFile::Read(void* lpBuf, UINT nCount)
{
	if (nCount == 0)
		return 0;
	if (lpBuf == NULL)
	{
	}

	if (this->m_nPosition > this->m_nFileSize)
		return 0;

	UINT nRead = 0;
	if (this->m_nPosition + nCount > m_nFileSize || this->m_nPosition + nCount < m_nPosition)
		nRead = (UINT)(this->m_nFileSize - this->m_nPosition);
	else
		nRead = nCount;

	Memcpy((BYTE*)lpBuf, (BYTE*)m_lpBuffer + m_nPosition, nRead);
	this->m_nPosition += nRead;

	return nRead;
}

void CMemFile::Write(const void* lpBuf, UINT nCount)
{
	if (nCount == 0)
		return;

	if (lpBuf == NULL)
	{
	}
	//If we have no room for nCount, it must be an overflow
	if (this->m_nPosition + nCount < this->m_nPosition)
	{
	}

	if (this->m_nPosition + nCount > m_nBufferSize)
		this->GrowFile(this->m_nPosition + nCount);


	Memcpy((BYTE*)m_lpBuffer + m_nPosition, (BYTE*)lpBuf, nCount);

	this->m_nPosition += nCount;

	if (this->m_nPosition > this->m_nFileSize)
		this->m_nFileSize = this->m_nPosition;

}

ULONGLONG CMemFile::Seek(LONGLONG lOff, UINT nFrom)
{
	LONGLONG lNewPos = this->m_nPosition;

	if (nFrom == begin)
		lNewPos = lOff;
	else if (nFrom == current)
		lNewPos += lOff;
	else if (nFrom == end) {
		if (lOff > 0)
		{
		}
		lNewPos = this->m_nFileSize + lOff;
	}
	else
		return this->m_nPosition;

	if (static_cast<DWORD>(lNewPos) > m_nFileSize)
		this->GrowFile((SIZE_T)lNewPos);

	this->m_nPosition = (SIZE_T)lNewPos;

	return this->m_nPosition;
}

void CMemFile::Flush()
{
}

void CMemFile::Close()
{
	this->m_nGrowBytes = 0;
	this->m_nPosition = 0;
	this->m_nBufferSize = 0;
	this->m_nFileSize = 0;
	if (this->m_lpBuffer!=nullptr && m_bAutoDelete)
		this->Free(this->m_lpBuffer);
	this->m_lpBuffer = nullptr;
}

void CMemFile::Abort()
{
	this->Close();
}
UINT CMemFile::GetBufferPtr(UINT nCommand, UINT nCount,
	void** ppBufStart, void** ppBufMax)
{
	if (nCommand == bufferCheck)
	{
		// only allow direct buffering if we're 
		// growable
		if (m_nGrowBytes > 0)
			return bufferDirect;
		else
			return 0;
	}

	if (nCommand == bufferCommit)
	{
		// commit buffer
		m_nPosition += nCount;
		if (m_nPosition > m_nFileSize)
			m_nFileSize = m_nPosition;
		return 0;
	}

	if (ppBufStart == NULL || ppBufMax == NULL)
	{
		return 0;
	}

	// when storing, grow file as necessary to satisfy buffer request
	if (nCommand == bufferWrite)
	{
		if (this->m_nPosition + nCount < this->m_nPosition 
			|| this->m_nPosition + nCount < nCount)
		{
		}
		if (this->m_nPosition + nCount > this->m_nBufferSize)
		{
			this->GrowFile(this->m_nPosition + nCount);
		}
	}

	// store buffer max and min
	*ppBufStart = m_lpBuffer + m_nPosition;

	// end of buffer depends on whether you are reading or writing
	if (nCommand == bufferWrite)
		*ppBufMax = this->m_lpBuffer + min(m_nBufferSize, m_nPosition + nCount);
	else
	{
		if (nCount == (UINT)-1)
			nCount = UINT(m_nBufferSize - this->m_nPosition);
		*ppBufMax = this->m_lpBuffer + min(this->m_nFileSize, this->m_nPosition + nCount);
		m_nPosition += LPBYTE(*ppBufMax) - LPBYTE(*ppBufStart);
	}

	// return number of bytes in returned buffer space (may be <= nCount)
	return ULONG(LPBYTE(*ppBufMax) - LPBYTE(*ppBufStart));
}
