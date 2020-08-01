#include "CMyOleDataSource.h"
#include <ShlObj.h>
#include <tchar.h>
#include <stdio.h>

void Log(const TCHAR* message) {
	FILE* f = _tfopen(_T("log.txt"), _T("a+"));
	if (f != nullptr) {
		_ftprintf(f, _T("%s\n"), message);
		fclose(f);
	}
}

long GetFileLength(const TCHAR* path)
{
	int length = 0;
	FILE* f=_tfopen(path, _T("rb"));
	if (f != nullptr) {
		fseek(f, 0, SEEK_END);
		length = ftell(f);
		fclose(f);
	}
	return length;
}
void PutFilesIntoClipboard(const TCHAR** paths, ULONGLONG* sizes, int count) {

}

HGLOBAL PutIntoClipboard(const TCHAR* path)
{
	CMyOleDataSource* pDataSrc = new CMyOleDataSource();
	UINT uFileCount = 1;
	UINT uBuffSize = sizeof(FILEGROUPDESCRIPTOR) +
		(uFileCount - 1) * sizeof(FILEDESCRIPTOR);
	HGLOBAL hFileDescriptor = GlobalAlloc(
		GHND | GMEM_SHARE, uBuffSize);
	if (hFileDescriptor)
	{
		FILEGROUPDESCRIPTOR* pGroupDescriptor =
			(FILEGROUPDESCRIPTOR*)GlobalLock(hFileDescriptor);
		if (pGroupDescriptor)
		{
			FILEDESCRIPTOR* pFileDescriptorArray =
				(FILEDESCRIPTOR*)((LPBYTE)pGroupDescriptor + sizeof(UINT));
			pGroupDescriptor->cItems = uFileCount;
			int index = 0;
			{
			
				ZeroMemory(&pFileDescriptorArray[index],
					sizeof(FILEDESCRIPTOR));
				const TCHAR* pname = _tcsrchr(path, '\\');
				//get the last backslash
				if (pname == nullptr) {
					pname = path;
				}
				else {
					pname++;
				}
				lstrcpy(pFileDescriptorArray[index].cFileName, pname);

				pDataSrc->m_sPaths = new TCHAR * [1];
				pDataSrc->m_nFiles = 1;
				pDataSrc->m_sPaths[0] = new TCHAR[MAX_PATH];
				lstrcpy(pDataSrc->m_sPaths[0], path);

				pFileDescriptorArray[index].dwFlags = FD_FILESIZE | FD_ATTRIBUTES;
				pFileDescriptorArray[index].nFileSizeLow = GetFileLength(path);
				pFileDescriptorArray[index].nFileSizeHigh = 0;
				pFileDescriptorArray[index].dwFileAttributes 
					= FILE_ATTRIBUTE_NORMAL;
				//FILE_ATTRIBUTE_DIRECTORY: for directory
			}
		}
		else
		{
			GlobalFree(hFileDescriptor);
		}
	}
	GlobalUnlock(hFileDescriptor);

	FORMATETC etcDescriptor = {
		(CLIPFORMAT)RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR),
		NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	pDataSrc->CacheGlobalData(RegisterClipboardFormat(
		CFSTR_FILEDESCRIPTOR), hFileDescriptor, &etcDescriptor);

	FORMATETC etcContents = {
		(CLIPFORMAT) (CFSTR_FILECONTENTS),
		NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	pDataSrc->DelayRenderFileData(RegisterClipboardFormat(
		CFSTR_FILECONTENTS), &etcContents);
	
	//SetClipboard will release pDataSrc
	pDataSrc->SetClipboard();

	return hFileDescriptor;
}

void SetClipboardWithFile() {
	
	TCHAR wd[MAX_PATH] = { 0 };
	GetCurrentDirectory(sizeof(wd), wd);
	size_t len = _tcslen(wd);
	if (len > 0) {
		if (wd[len - 1] != _T('\\')) {
			lstrcat(wd, _T("\\"));
		}
	}
	lstrcat(wd, _T("test.txt"));
	if ((_taccess(wd, 0)) != -1){
		PutIntoClipboard(wd);
	}
}
