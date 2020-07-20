#pragma once
#include "CFile.h"

class COleStreamFile : public CFile
{
private:
	using CFile::Open;

	// Constructors and Destructors
public:
	explicit COleStreamFile(LPSTREAM lpStream = NULL);
	
	BOOL CreateMemoryStream();

	// Operations
		// Note: OpenStream and CreateStream can accept eith STGM_ flags or
		//  CFile::OpenFlags bits since common values are guaranteed to have
		//  the same semantics.
	BOOL OpenStream(LPSTORAGE lpStorage, LPCTSTR lpszStreamName,
		DWORD nOpenFlags = modeReadWrite | shareExclusive
	);
	BOOL CreateStream(LPSTORAGE lpStorage, LPCTSTR lpszStreamName,
		DWORD nOpenFlags = modeReadWrite | shareExclusive | modeCreate);


	// attach & detach can be used when Open/Create functions aren't adequate
	void Attach(LPSTREAM lpStream);
	LPSTREAM Detach();

	IStream* GetStream() const { return this->m_lpStream; }
	// Returns the current stream

// Implementation
public:
	LPSTREAM m_lpStream;

	virtual ~COleStreamFile();

	// attributes for implementation
	//BOOL GetStatus(CFileStatus& rStatus) const;
	virtual ULONGLONG GetPosition() const;

	virtual const TCHAR* GetStorageName() const {
		return this->m_strStorageName;
	}

	// overrides for implementation
	virtual ULONGLONG Seek(LONGLONG lOff, UINT nFrom);
	virtual void SetLength(ULONGLONG dwNewLen);
	virtual ULONGLONG GetLength() const;
	virtual UINT Read(void* lpBuf, UINT nCount);
	virtual void Write(const void* lpBuf, UINT nCount);
	virtual void LockRange(ULONGLONG dwPos, ULONGLONG dwCount);
	virtual void UnlockRange(ULONGLONG dwPos, ULONGLONG dwCount);
	virtual void Abort();
	virtual void Flush();
	virtual void Close();

protected:

	TCHAR* m_strStorageName;
};
