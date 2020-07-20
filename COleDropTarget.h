#pragma once
#include <Windows.h>
#include "COleDropSource.h"
class COleDataObject;
class COleDropTarget : IDropTarget
{
public:
	// IUnknown members
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG   __stdcall AddRef(void);
	ULONG   __stdcall Release(void);

public:
	virtual void OnFinalRelease();
public:
	COleDropTarget();
	virtual ~COleDropTarget();
public:

// Overridables
    virtual HRESULT STDMETHODCALLTYPE DragEnter(
        /* [unique][in] */ __RPC__in_opt IDataObject* pDataObj,
        /* [in] */ DWORD grfKeyState,
        /* [in] */ POINTL pt,
        /* [out][in] */ __RPC__inout DWORD* pdwEffect);

    virtual HRESULT STDMETHODCALLTYPE DragOver(
        /* [in] */ DWORD grfKeyState,
        /* [in] */ POINTL pt,
        /* [out][in] */ __RPC__inout DWORD* pdwEffect);

    virtual HRESULT STDMETHODCALLTYPE DragLeave(void);

    virtual HRESULT STDMETHODCALLTYPE Drop(
        /* [unique][in] */ __RPC__in_opt IDataObject* pDataObj,
        /* [in] */ DWORD grfKeyState,
        /* [in] */ POINTL pt,
        /* [out][in] */ __RPC__inout DWORD* pdwEffect);

protected:
	LONG m_lRefCount;
};

