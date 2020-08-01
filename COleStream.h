#pragma once

#include <objidl.h>

class COleStream 
    : public IStream
{
public:

	// IUnknown members
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG   __stdcall AddRef(void);
	ULONG   __stdcall Release(void);

public:
	//ISequentialStream
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Read(
        /* [annotation] */
        _Out_writes_bytes_to_(cb, *pcbRead)  void* pv,
        /* [annotation][in] */
        _In_  ULONG cb,
        /* [annotation] */
        _Out_opt_  ULONG* pcbRead);

    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Write(
        /* [annotation] */
        _In_reads_bytes_(cb)  const void* pv,
        /* [annotation][in] */
        _In_  ULONG cb,
        /* [annotation] */
        _Out_opt_  ULONG* pcbWritten);
public:

    //IStream
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Seek(
        /* [in] */ LARGE_INTEGER dlibMove,
        /* [in] */ DWORD dwOrigin,
        _Out_opt_  ULARGE_INTEGER* plibNewPosition);

    virtual HRESULT STDMETHODCALLTYPE SetSize(
        /* [in] */ ULARGE_INTEGER libNewSize);

    virtual /* [local] */ HRESULT STDMETHODCALLTYPE CopyTo(
        /* [annotation][unique][in] */
        _In_  IStream* pstm,
        /* [in] */ ULARGE_INTEGER cb,
        /* [annotation] */
        _Out_opt_  ULARGE_INTEGER* pcbRead,
        /* [annotation] */
        _Out_opt_  ULARGE_INTEGER* pcbWritten);

    virtual HRESULT STDMETHODCALLTYPE Commit(
        /* [in] */ DWORD grfCommitFlags);

    virtual HRESULT STDMETHODCALLTYPE Revert(void);

    virtual HRESULT STDMETHODCALLTYPE LockRegion(
        /* [in] */ ULARGE_INTEGER libOffset,
        /* [in] */ ULARGE_INTEGER cb,
        /* [in] */ DWORD dwLockType);

    virtual HRESULT STDMETHODCALLTYPE UnlockRegion(
        /* [in] */ ULARGE_INTEGER libOffset,
        /* [in] */ ULARGE_INTEGER cb,
        /* [in] */ DWORD dwLockType);

    virtual HRESULT STDMETHODCALLTYPE Stat(
        /* [out] */ __RPC__out STATSTG* pstatstg,
        /* [in] */ DWORD grfStatFlag);

    virtual HRESULT STDMETHODCALLTYPE Clone(
        /* [out] */ __RPC__deref_out_opt IStream** ppstm);
public:
    COleStream();
    virtual ~COleStream();
public:
    virtual void OnFinalRelease();

protected:
    LONG m_lRefCount;
    ULARGE_INTEGER m_Pos;
    ULARGE_INTEGER m_Length;
};

