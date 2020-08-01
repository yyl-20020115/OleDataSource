#pragma once
#include "COleDropTarget.h"
class CMyOleDropTarget :
    public COleDropTarget
{
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
};

