#include "COleStreamFile.h"

COleStreamFile::COleStreamFile(LPSTREAM lpStream)
	: m_lpStream(lpStream)
	, m_strStorageName(nullptr)
{
	if (this->m_lpStream != nullptr)
	{
	}
}
COleStreamFile::~COleStreamFile()
{
	if (this->m_lpStream != nullptr && this->m_bCloseOnDelete)
	{
		this->Close();
	}
}
LPSTREAM COleStreamFile::Detach()
{
	LPSTREAM lpStream = m_lpStream;
	m_lpStream = NULL;  // detach and transfer ownership of m_lpStream
	return lpStream;
}
void COleStreamFile::Attach(LPSTREAM lpStream)
{
	if (lpStream == NULL)
	{
	}

	this->m_lpStream = lpStream;
	this->m_bCloseOnDelete = FALSE;
}
HRESULT _AfxReadFromStream(LPSTREAM pStream, void* lpBuf, UINT nCount, DWORD& nRead)
{
	if (nCount == 0)
	{
		nRead = 0;
		return S_OK;
	}

	if (pStream == NULL || lpBuf == NULL)
	{
		return E_INVALIDARG;
	}

	// read from the stream
	SCODE sc = pStream->Read(lpBuf, nCount, &nRead);
	return ResultFromScode(sc);
}

ULONGLONG COleStreamFile::GetPosition() const
{
	ULARGE_INTEGER liPosition = { 0 };
	LARGE_INTEGER liZero = { 0 }; 
	
	SCODE sc = m_lpStream->Seek(liZero, STREAM_SEEK_CUR, &liPosition);
	if (sc != S_OK)
	{
	}
	return liPosition.QuadPart;
}
BOOL COleStreamFile::OpenStream(LPSTORAGE lpStorage, LPCTSTR lpszStreamName,
	DWORD nOpenFlags)
{
	if (lpStorage == NULL || lpszStreamName == NULL)
	{
		return FALSE;
	}

	SCODE sc = lpStorage->OpenStream(lpszStreamName, NULL, nOpenFlags, 0, &m_lpStream);

	return !FAILED(sc);
}

BOOL COleStreamFile::CreateStream(LPSTORAGE lpStorage, LPCTSTR lpszStreamName,
	DWORD nOpenFlags)
{
	if (lpStorage == NULL || lpszStreamName == NULL)
	{
		return FALSE;
	}
	SCODE sc = lpStorage->CreateStream(lpszStreamName, nOpenFlags, 0, 0, &m_lpStream);

	return !FAILED(sc);
}

BOOL COleStreamFile::CreateMemoryStream()
{
	SCODE sc = CreateStreamOnHGlobal(NULL, TRUE, &m_lpStream);

	return !FAILED(sc);
}
ULONGLONG COleStreamFile::Seek(LONGLONG lOff, UINT nFrom)
{
	ULARGE_INTEGER liNewPosition;
	LARGE_INTEGER liOff;
	liOff.QuadPart = lOff;
	SCODE sc = m_lpStream->Seek(liOff, nFrom, &liNewPosition);

	return liNewPosition.QuadPart;
}

void COleStreamFile::SetLength(ULONGLONG dwNewLen)
{
	ULARGE_INTEGER liNewLen;
	liNewLen.QuadPart = dwNewLen;
	SCODE sc = m_lpStream->SetSize(liNewLen);
}

ULONGLONG COleStreamFile::GetLength() const
{
	// get status of the stream
	STATSTG statstg;
	SCODE sc = m_lpStream->Stat(&statstg, STATFLAG_NONAME);
	// map to CFileStatus struct
	return statstg.cbSize.QuadPart;
}

UINT COleStreamFile::Read(void* lpBuf, UINT nCount)
{
	DWORD dwBytesRead;
	HRESULT hr = _AfxReadFromStream(m_lpStream, lpBuf, nCount, dwBytesRead);

	// always return number of bytes read
	return (UINT)dwBytesRead;
}

void COleStreamFile::Write(const void* lpBuf, UINT nCount)
{
	if (nCount == 0)
		return;

	// write to the stream
	DWORD dwBytesWritten;
	SCODE sc = m_lpStream->Write(lpBuf, nCount, &dwBytesWritten);

	// if no error, all bytes should have been written
}

void COleStreamFile::LockRange(ULONGLONG dwPos, ULONGLONG dwCount)
{
	// convert parameters to long integers
	ULARGE_INTEGER liPos;
	liPos.QuadPart = dwPos;
	ULARGE_INTEGER liCount;
	liCount.QuadPart = dwCount;

	// then lock the region
	SCODE sc = m_lpStream->LockRegion(liPos, liCount, LOCK_EXCLUSIVE);
}

void COleStreamFile::UnlockRange(ULONGLONG dwPos, ULONGLONG dwCount)
{
	// convert parameters to long integers
	ULARGE_INTEGER liPos;
	liPos.QuadPart = dwPos;
	ULARGE_INTEGER liCount;
	liCount.QuadPart = dwCount;

	// then lock the region
	SCODE sc = m_lpStream->UnlockRegion(liPos, liCount, LOCK_EXCLUSIVE);
}

void COleStreamFile::Abort()
{
	if (m_lpStream != NULL)
	{
		m_lpStream->Revert();
		m_lpStream->Release();
		m_lpStream = nullptr;
	}
	delete m_strStorageName;
}

void COleStreamFile::Flush()
{
	SCODE sc = m_lpStream->Commit(0);
}

void COleStreamFile::Close()
{
	if (this->m_lpStream != NULL)
	{
		// commit the stream via Flush (which can be overriden)
		this->Flush();
		this->m_lpStream->Release();
		this->m_lpStream = nullptr;
	}
	if (this->m_strStorageName != nullptr) {
		delete[] this->m_strStorageName;
	}
	this->m_strStorageName = nullptr;
}




