#include "CMyOleDropTarget.h"
#include "framework.h"


//NOTICE:
//PLEASE READ THIS ARTICLE:
//http://www.cppblog.com/andxie99/archive/2007/03/05/19228.html

HRESULT __stdcall CMyOleDropTarget::DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    Log(_T("Drag Enter"));
    *pdwEffect = DROPEFFECT_MOVE;
    return S_OK;
}

HRESULT __stdcall CMyOleDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    Log(_T("Drag Over"));
    *pdwEffect = DROPEFFECT_MOVE;
    return S_OK;
}

HRESULT __stdcall CMyOleDropTarget::DragLeave(void)
{
    Log(_T("Drag Leave"));
    return S_OK;
}

HRESULT __stdcall CMyOleDropTarget::Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    Log(_T("Drag Drop"));

    *pdwEffect = DROPEFFECT_MOVE;
    return S_OK;
}
