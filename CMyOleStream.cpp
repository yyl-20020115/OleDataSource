#include "CMyOleStream.h"
#include<tchar.h>
CMyOleStream::CMyOleStream(const TCHAR* sp)
	: m_srcPath(_tcsdup(sp))
	, m_pFile(nullptr)
{
	this->m_pFile = _tfopen(this->m_srcPath, _T("rb"));
	if (this->m_pFile != nullptr) {
		_fseeki64(this->m_pFile, 0, SEEK_END);
		this->m_Length.QuadPart = _ftelli64(this->m_pFile);
	}
}

CMyOleStream::~CMyOleStream()
{
	if (this->m_srcPath != nullptr) {
		free(this->m_srcPath);
	}
	this->m_srcPath = nullptr;
	if (this->m_pFile != nullptr) {
		fclose(this->m_pFile);
		this->m_pFile = nullptr;
	}
}

HRESULT __stdcall CMyOleStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
	*pcbRead = 0;
	if (this->m_pFile != nullptr) {
		//NOTICE: read bytes from source file and save data to buffer pointed by pv
		*pcbRead = fread(pv, 1, cb, this->m_pFile);

		return S_OK;
	}
	return E_FAIL;
}

HRESULT __stdcall CMyOleStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
{
	if (this->m_pFile != nullptr) {

		_fseeki64(this->m_pFile, dlibMove.QuadPart, dwOrigin);
		this->m_Pos.QuadPart = plibNewPosition->QuadPart = _ftelli64(this->m_pFile);
		return S_OK;
	}
	return E_FAIL;
}
