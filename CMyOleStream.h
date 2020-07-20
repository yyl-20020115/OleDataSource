#pragma once
#include "COleStream.h"
#include<stdio.h>

class CMyOleStream :
    public COleStream
{
public:
	CMyOleStream(const TCHAR* sp);
	virtual ~CMyOleStream();
public:
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Read(
        _Out_writes_bytes_to_(cb, *pcbRead)  void* pv,
        _In_  ULONG cb,
        _Out_opt_  ULONG* pcbRead);
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Seek(
        /* [in] */ LARGE_INTEGER dlibMove,
        /* [in] */ DWORD dwOrigin,
        _Out_opt_  ULARGE_INTEGER* plibNewPosition);

private:
	TCHAR* m_srcPath;
    FILE* m_pFile;
};

