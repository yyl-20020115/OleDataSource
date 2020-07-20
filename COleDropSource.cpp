#include "COleDropSource.h"

HRESULT __stdcall COleDropSource::QueryInterface(REFIID iid, void** ppvObject)
{
	//Only provide myself
	if (iid == IID_IDropSource)
	{
		this->AddRef();
		*ppvObject = static_cast<IDropSource*>(this);
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

ULONG __stdcall COleDropSource::AddRef(void)
{
	return InterlockedIncrement(&this->m_lRefCount);
}

ULONG __stdcall COleDropSource::Release(void)
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

void COleDropSource::OnFinalRelease()
{
	delete this;
}

COleDropSource::COleDropSource()
	:m_lRefCount(0)
{
	this->m_lRefCount = 1;
}

COleDropSource::~COleDropSource()
{
}

HRESULT __stdcall COleDropSource::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
	return S_OK;
}

HRESULT __stdcall COleDropSource::GiveFeedback(DWORD dwEffect)
{
	return S_OK;
}
