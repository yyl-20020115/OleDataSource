#include "COleStream.h"

COleStream::COleStream()
	: m_lRefCount(0)
	, m_Pos()
	, m_Length()
{
	this->m_lRefCount = 1;
}

COleStream::~COleStream()
{
}

void COleStream::OnFinalRelease()
{
	delete this;
}

HRESULT __stdcall COleStream::QueryInterface(REFIID iid, void** ppvObject)
{
	if (iid == IID_IStream)
	{
		this->AddRef();
		*ppvObject = static_cast<IStream*>(this);
		return S_OK;
	}
	else if (iid == IID_ISequentialStream)
	{
		this->AddRef();
		*ppvObject = static_cast<ISequentialStream*>(this);
		return S_OK;
	}
	else if (iid == IID_IUnknown)
	{
		this->AddRef();
		*ppvObject = static_cast<IUnknown*>(this);
		return S_OK;
	}
	return (DWORD)E_NOINTERFACE;
}

ULONG __stdcall COleStream::AddRef(void)
{
	return InterlockedIncrement(&this->m_lRefCount);
}

ULONG __stdcall COleStream::Release(void)
{
	if (this->m_lRefCount == 0) {
		return 0;
	}
	ULONG c = InterlockedDecrement(&this->m_lRefCount);
	if (c == 0) {
		this->OnFinalRelease();
	}
	return c;
}
HRESULT __stdcall COleStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
	*pcbRead = 0;
	return S_OK;
}

HRESULT __stdcall COleStream::Write(const void* pv, ULONG cb, ULONG* pcbWritten)
{
	*pcbWritten = cb;
	return S_OK;
}

HRESULT __stdcall COleStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
{
	switch (dwOrigin) {
	case STREAM_SEEK_SET:
		this->m_Pos.QuadPart = dlibMove.QuadPart;
		break;
	case STREAM_SEEK_CUR:
		this->m_Pos.QuadPart += dlibMove.QuadPart;
		break;
	case STREAM_SEEK_END:
		this->m_Pos.QuadPart = m_Length.QuadPart - dlibMove.QuadPart;
		break;
	}
	plibNewPosition->QuadPart = this->m_Pos.QuadPart;
	return S_OK;
}

HRESULT __stdcall COleStream::SetSize(ULARGE_INTEGER libNewSize)
{
	this->m_Length.QuadPart = libNewSize.QuadPart;
	return S_OK;
}

HRESULT __stdcall COleStream::CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten)
{
	*pcbRead = *pcbWritten = cb;

	return S_OK;
}

HRESULT __stdcall COleStream::Commit(DWORD grfCommitFlags)
{
	return S_OK;
}

HRESULT __stdcall COleStream::Revert(void)
{
	return S_OK;
}

HRESULT __stdcall COleStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return S_OK;
}

HRESULT __stdcall COleStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return S_OK;
}

HRESULT __stdcall COleStream::Stat(STATSTG* pstatstg, DWORD grfStatFlag)
{
	*pstatstg = STATSTG();
	return S_OK;
}

HRESULT __stdcall COleStream::Clone(IStream** ppstm)
{
	*ppstm = nullptr;

	return S_OK;
}
