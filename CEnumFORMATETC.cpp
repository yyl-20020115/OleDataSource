#include "CEnumFORMATETC.h"
DVTARGETDEVICE* ODSOleCopyTargetDevice(DVTARGETDEVICE* ptdSrc)
{
    if (ptdSrc == NULL)
        return NULL;

    DVTARGETDEVICE* ptdDest =
        (DVTARGETDEVICE*)CoTaskMemAlloc(ptdSrc->tdSize);
    if (ptdDest == NULL)
        return NULL;

    memcpy_s(ptdDest, (size_t)ptdSrc->tdSize,
        ptdSrc, (size_t)ptdSrc->tdSize);
    return ptdDest;
}
HRESULT __stdcall CEnumFormatEtc::QueryInterface(REFIID iid, void** ppvObject)
{
    //Only provide myself
    if (iid == IID_IEnumFORMATETC)
    {
        this->AddRef();
        *ppvObject = static_cast<IEnumFORMATETC*>(this);
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

ULONG __stdcall CEnumFormatEtc::AddRef(void)
{
    return InterlockedIncrement(&this->m_lRefCount);
}

ULONG __stdcall CEnumFormatEtc::Release(void)
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

HRESULT __stdcall CEnumFormatEtc::Next(ULONG celt, FORMATETC* reelt, ULONG* pceltFetched)
{
    if (pceltFetched != NULL)
        *pceltFetched = 0;

    BYTE* pchCur = (BYTE*)reelt;

    ULONG celtT = celt;
    SCODE sc = E_UNEXPECTED;
    try
    {
        while (celtT != 0 && this->OnNext((void*)pchCur))
        {
            pchCur += this->m_nSizeElem;
            --celtT;
        }
        if (pceltFetched != NULL)
            *pceltFetched = celt - celtT;
        sc = celtT == 0 ? S_OK : S_FALSE;
    }
    catch (...) {

    }
    return sc;
}

HRESULT __stdcall CEnumFormatEtc::Skip(ULONG celt)
{
    ULONG celtT = celt;
    SCODE sc = E_UNEXPECTED;
    try
    {
        while (celtT != 0 && this->OnSkip())
            --celtT;
        sc = celtT == 0 ? S_OK : S_FALSE;
    }
    catch (...) {

    }
    return celtT != 0 ? S_FALSE : S_OK;
}

HRESULT __stdcall CEnumFormatEtc::Reset(void)
{
    this->OnReset();

    return S_OK;
}

HRESULT __stdcall CEnumFormatEtc::Clone(IEnumFORMATETC** ppenm)
{
    *ppenm = NULL;

    SCODE sc = E_UNEXPECTED;
    try
    {
        CEnumFormatEtc* clone = this->OnClone();

        // we use an extra reference to keep the original object alive
        //  (the extra reference is removed in the clone's destructor)
        if (this->m_pClonedFrom != NULL)
            clone->m_pClonedFrom = this->m_pClonedFrom;
        else
            clone->m_pClonedFrom = this;
        clone->m_pClonedFrom->AddRef();
        *ppenm = clone;

        sc = S_OK;
    }
    catch (...) {

    }

    return sc;
}

CEnumFormatEtc::CEnumFormatEtc()
    : m_lRefCount(0)
    , m_nSizeElem(0)
    , m_pvEnum(nullptr)
    , m_nSize(0)
    , m_nCurPos(0)
    , m_bNeedFree(TRUE)
    , m_pClonedFrom(nullptr)
    , m_nMaxSize(0)
{
    this->m_lRefCount = 1;
    this->m_nSizeElem = sizeof(FORMATETC);
}

CEnumFormatEtc::~CEnumFormatEtc()
{
    if (this->m_pClonedFrom != nullptr) {
        this->m_pClonedFrom->Release();
    }
    else  if (this->m_pClonedFrom == nullptr) {

        // release all of the pointers to DVTARGETDEVICE
        LPFORMATETC lpFormatEtc = (LPFORMATETC)m_pvEnum;
        for (UINT nIndex = 0; nIndex < m_nSize; nIndex++)
            CoTaskMemFree(lpFormatEtc[nIndex].ptd);

    }
    this->m_pClonedFrom = nullptr;
    if (this->m_bNeedFree && this->m_pvEnum!=nullptr) {
        delete this->m_pvEnum;
    }
    this->m_pvEnum = nullptr;
}

void CEnumFormatEtc::AddFormat(const FORMATETC* lpFormatEtc)
{
    if (m_nSize == m_nMaxSize)
    {
        // not enough space for new item -- allocate more
        FORMATETC* pListNew = new FORMATETC[m_nSize + 10];
        m_nMaxSize += 10;
        memcpy_s(pListNew, (m_nSize + 10) * sizeof(FORMATETC),
            m_pvEnum, m_nSize * sizeof(FORMATETC));
        delete m_pvEnum;
        m_pvEnum = (BYTE*)pListNew;
    }

    // add this item to the list
    FORMATETC* pFormat = &((FORMATETC*)m_pvEnum)[m_nSize];
    pFormat->cfFormat = lpFormatEtc->cfFormat;
    pFormat->ptd = lpFormatEtc->ptd;
    // Note: ownership of lpFormatEtc->ptd is transfered with this call.
    pFormat->dwAspect = lpFormatEtc->dwAspect;
    pFormat->lindex = lpFormatEtc->lindex;
    pFormat->tymed = lpFormatEtc->tymed;
    ++m_nSize;
}

BOOL CEnumFormatEtc::OnNext(void* pv)
{
    if (m_nCurPos >= m_nSize)
        return FALSE;

    memcpy_s(pv, m_nSizeElem,
        &m_pvEnum[m_nCurPos * m_nSizeElem], m_nSizeElem);

    ++m_nCurPos;
    // any outgoing formatEtc may require the DVTARGETDEVICE to
//  be copied (the caller has responsibility to free it)
    LPFORMATETC lpFormatEtc = (LPFORMATETC)pv;
    if (lpFormatEtc->ptd != NULL)
    {
        lpFormatEtc->ptd = ODSOleCopyTargetDevice(lpFormatEtc->ptd);
        if (lpFormatEtc->ptd == NULL)
            return FALSE;
    }
    // otherwise, copying worked...
    return TRUE;
}

BOOL CEnumFormatEtc::OnSkip()
{
    if (m_nCurPos >= m_nSize)
        return FALSE;

    return ++m_nCurPos < m_nSize;
}

void CEnumFormatEtc::OnReset()
{
    m_nCurPos = 0;
}

CEnumFormatEtc* CEnumFormatEtc::OnClone()
{
    // set up an exact copy of this object
    //  (derivatives may have to replace this code)
    CEnumFormatEtc* pClone = new CEnumFormatEtc();
    pClone->m_bNeedFree = FALSE;
     // clones should never free themselves
    pClone->m_nCurPos = m_nCurPos;

    return pClone;
}

void CEnumFormatEtc::OnFinalRelease()
{
    delete this;
}
