#pragma once
#include <Windows.h>

class CFile
{
public:
	// Flag values
	enum OpenFlags {
		modeRead = (int)0x00000,
		modeWrite = (int)0x00001,
		modeReadWrite = (int)0x00002,
		shareCompat = (int)0x00000,
		shareExclusive = (int)0x00010,
		shareDenyWrite = (int)0x00020,
		shareDenyRead = (int)0x00030,
		shareDenyNone = (int)0x00040,
		modeNoInherit = (int)0x00080,
#ifdef _UNICODE
		typeUnicode = (int)0x00400, // used in derived classes (e.g. CStdioFile) only
#endif
		modeCreate = (int)0x01000,
		modeNoTruncate = (int)0x02000,
		typeText = (int)0x04000, // used in derived classes (e.g. CStdioFile) only
		typeBinary = (int)0x08000, // used in derived classes (e.g. CStdioFile) only
		osNoBuffer = (int)0x10000,
		osWriteThrough = (int)0x20000,
		osRandomAccess = (int)0x40000,
		osSequentialScan = (int)0x80000,
	};

	enum Attribute {
		normal = 0x00,                // note: not same as FILE_ATTRIBUTE_NORMAL
		readOnly = FILE_ATTRIBUTE_READONLY,
		hidden = FILE_ATTRIBUTE_HIDDEN,
		system = FILE_ATTRIBUTE_SYSTEM,
		volume = 0x08,
		directory = FILE_ATTRIBUTE_DIRECTORY,
		archive = FILE_ATTRIBUTE_ARCHIVE,
		device = FILE_ATTRIBUTE_DEVICE,
		temporary = FILE_ATTRIBUTE_TEMPORARY,
		sparse = FILE_ATTRIBUTE_SPARSE_FILE,
		reparsePt = FILE_ATTRIBUTE_REPARSE_POINT,
		compressed = FILE_ATTRIBUTE_COMPRESSED,
		offline = FILE_ATTRIBUTE_OFFLINE,
		notIndexed = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED,
		encrypted = FILE_ATTRIBUTE_ENCRYPTED
	};

	enum SeekPosition { begin = 0x0, current = 0x1, end = 0x2 };

	static const HANDLE hFileNull;

	// Constructors
	CFile();

	// Attributes
	HANDLE m_hFile;
	operator HANDLE() const { return m_hFile; };

	virtual ULONGLONG GetPosition() { return 0; }
	//BOOL GetStatus(CFileStatus& rStatus) const;
	virtual const TCHAR* GetFileName() const { return nullptr; }
	virtual const TCHAR* GetFileTitle() const { return nullptr; }
	virtual const TCHAR* GetFilePath() const { return nullptr; }
	virtual void SetFilePath(LPCTSTR lpszNewName) {}

	
	/// <summary>
	/// Open is designed for use with the default CFile constructor</summary>
	/// <returns>
	/// TRUE if succeeds; otherwise FALSE.</returns>
	/// <param name="lpszFileName">A string that is the path to the desired file. The path can be relative or absolute.</param>
	/// <param name="nOpenFlags">Sharing and access mode. Specifies the action to take when opening the file. You can combine options listed below by using the bitwise-OR (|) operator. One access permission and one share option are required; the modeCreate and modeNoInherit modes are optional.</param>
	/// <param name="pTM">Pointer to CAtlTransactionManager object</param>
	/// <param name="pError">A pointer to an existing file-exception object that will receive the status of a failed operation</param>
	virtual BOOL Open(LPCTSTR lpszFileName, UINT nOpenFlags) {
		return FALSE;
	};

	virtual ULONGLONG SeekToEnd() { return 0; }
	virtual void SeekToBegin() {}

	// Overridables

	virtual ULONGLONG Seek(LONGLONG lOff, UINT nFrom) { return 0LL; }
	virtual void SetLength(ULONGLONG dwNewLen) {}
	virtual ULONGLONG GetLength() const {
		return 0LL;
	}

	virtual UINT Read(void* lpBuf, UINT nCount) { return 0; }
	virtual void Write(const void* lpBuf, UINT nCount) {}

	virtual void LockRange(ULONGLONG dwPos, ULONGLONG dwCount) {}
	virtual void UnlockRange(ULONGLONG dwPos, ULONGLONG dwCount) {}

	virtual void Abort() {}
	virtual void Flush() {}
	virtual void Close() {}

	// Implementation
public:
	virtual ~CFile() {}

	enum BufferCommand { bufferRead, bufferWrite, bufferCommit, bufferCheck };
	enum BufferFlags
	{
		bufferDirect = 0x01,
		bufferBlocking = 0x02
	};
	virtual UINT GetBufferPtr(UINT nCommand, UINT nCount = 0,
		void** ppBufStart = NULL, void** ppBufMax = NULL) {		
		return 0;
	}

protected:

	BOOL m_bCloseOnDelete;
	TCHAR* m_strFileName;
};
