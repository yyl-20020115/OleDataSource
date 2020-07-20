#pragma once
#include <Windows.h>
#include <ObjIdl.h>

class CEnumFormatEtc : public IEnumFORMATETC
{
public:

    // IUnknown members
    HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject);
    ULONG   __stdcall AddRef(void);
    ULONG   __stdcall Release(void);
public:
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Next(
        /* [in] */ ULONG celt,
        /* [annotation] */
        _Out_writes_to_(celt, *pceltFetched)  FORMATETC* rgelt,
        /* [annotation] */
        _Out_opt_  ULONG* pceltFetched);

    virtual HRESULT STDMETHODCALLTYPE Skip(
        /* [in] */ ULONG celt);

    virtual HRESULT STDMETHODCALLTYPE Reset(void);

    virtual HRESULT STDMETHODCALLTYPE Clone(
        /* [out] */ __RPC__deref_out_opt IEnumFORMATETC** ppenum);
public:

    CEnumFormatEtc();
    virtual ~CEnumFormatEtc();
public:
    void AddFormat(const FORMATETC* lpFormatEtc);
    virtual BOOL OnNext(void* pv);
    virtual BOOL OnSkip();
    virtual void OnReset();
    virtual CEnumFormatEtc* OnClone();

public:
    virtual void OnFinalRelease();

private:
    IEnumFORMATETC* m_pClonedFrom;
    size_t m_nSizeElem;     // size of each item in the array

    LONG m_lRefCount;
    UINT m_nMaxSize;
    UINT m_nSize;       // total number of items in m_pvEnum
    BYTE* m_pvEnum;     // pointer data to enumerate
    UINT m_nCurPos;     // current position in m_pvEnum
    BOOL m_bNeedFree;   // free on release?

};

DVTARGETDEVICE* ODSOleCopyTargetDevice(DVTARGETDEVICE* ptdSrc);


