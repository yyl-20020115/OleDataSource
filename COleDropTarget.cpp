#include "COleDropTarget.h"

HRESULT __stdcall COleDropTarget::QueryInterface(REFIID iid, void** ppvObject)
{
	//Only provide myself
	if (iid == IID_IDropTarget)
	{
		this->AddRef();
		*ppvObject = static_cast<IDropTarget*>(this);
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

ULONG __stdcall COleDropTarget::AddRef(void)
{
	return InterlockedIncrement(&this->m_lRefCount);
}

ULONG __stdcall COleDropTarget::Release(void)
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

void COleDropTarget::OnFinalRelease()
{
	delete this;
}

COleDropTarget::COleDropTarget()
	: m_lRefCount(0)
{
	this->m_lRefCount = 1;
}

COleDropTarget::~COleDropTarget()
{
}

HRESULT __stdcall COleDropTarget::DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	return S_OK;
}

HRESULT __stdcall COleDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	return S_OK;
}

HRESULT __stdcall COleDropTarget::DragLeave(void)
{
	return S_OK;
}

HRESULT __stdcall COleDropTarget::Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
	*pdwEffect = DROPEFFECT_NONE;
	return S_OK;
}
