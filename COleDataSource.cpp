#include "COleDataSource.h"
#include "CEnumFormatEtc.h"
#include "CSharedFile.h"
#include "COleStreamFile.h"

// Helper for creating default FORMATETC from cfFormat
LPFORMATETC FillFormatEtcHelper(
    LPFORMATETC lpFormatEtc, CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtcFill)
{
    //ENSURE(lpFormatEtcFill != nullptr);
    if (lpFormatEtc == nullptr && cfFormat != 0)
    {
        lpFormatEtc = lpFormatEtcFill;
        lpFormatEtc->cfFormat = cfFormat;
        lpFormatEtc->ptd = nullptr;
        lpFormatEtc->dwAspect = DVASPECT_CONTENT;
        lpFormatEtc->lindex = -1;
        lpFormatEtc->tymed = (DWORD)-1;
    }
    return lpFormatEtc;
}

void ODSOleCopyFormatEtc(LPFORMATETC petcDest, LPFORMATETC petcSrc)
{

    petcDest->cfFormat = petcSrc->cfFormat;
    petcDest->ptd = ODSOleCopyTargetDevice(petcSrc->ptd);
    petcDest->dwAspect = petcSrc->dwAspect;
    petcDest->lindex = petcSrc->lindex;
    petcDest->tymed = petcSrc->tymed;
}
HGLOBAL ODSCopyGlobalMemory(HGLOBAL hDest, HGLOBAL hSource)
{
    // make sure we have suitable hDest
    ULONG_PTR nSize = ::GlobalSize(hSource);
    if (hDest == NULL)
    {
        hDest = ::GlobalAlloc(GMEM_SHARE | GMEM_MOVEABLE, nSize);
        if (hDest == NULL)
            return NULL;
    }
    else if (nSize > ::GlobalSize(hDest))
    {
        // hDest is not large enough
        return NULL;
    }

    // copy the bits
    LPVOID lpSource = ::GlobalLock(hSource);
    LPVOID lpDest = ::GlobalLock(hDest);
    memcpy_s(lpDest, ::GlobalSize(hDest), lpSource, (ULONG)nSize);
    ::GlobalUnlock(hDest);
    ::GlobalUnlock(hSource);

    // success -- return hDest
    return hDest;
}
BOOL ODSCopyStgMedium(
    CLIPFORMAT cfFormat, LPSTGMEDIUM lpDest, LPSTGMEDIUM lpSource)
{
    if (lpDest->tymed == TYMED_NULL)
    {
        switch (lpSource->tymed)
        {
        case TYMED_ENHMF:
        case TYMED_HGLOBAL:
            lpDest->tymed = lpSource->tymed;
            lpDest->hGlobal = NULL;
            break;  // fall through to CopyGlobalMemory case

        case TYMED_ISTREAM:
            lpDest->pstm = lpSource->pstm;
            lpDest->pstm->AddRef();
            lpDest->tymed = TYMED_ISTREAM;
            return TRUE;

        case TYMED_ISTORAGE:
            lpDest->pstg = lpSource->pstg;
            lpDest->pstg->AddRef();
            lpDest->tymed = TYMED_ISTORAGE;
            return TRUE;

        case TYMED_MFPICT:
        {
            // copy LPMETAFILEPICT struct + embedded HMETAFILE
            HGLOBAL hDest = ODSCopyGlobalMemory(NULL, lpSource->hGlobal);
            if (hDest == NULL)
                return FALSE;
            LPMETAFILEPICT lpPict = (LPMETAFILEPICT)::GlobalLock(hDest);
            lpPict->hMF = ::CopyMetaFile(lpPict->hMF, NULL);
            if (lpPict->hMF == NULL)
            {
                ::GlobalUnlock(hDest);
                ::GlobalFree(hDest);
                return FALSE;
            }
            ::GlobalUnlock(hDest);

            // fill STGMEDIUM struct
            lpDest->hGlobal = hDest;
            lpDest->tymed = TYMED_MFPICT;
        }
        return TRUE;

        case TYMED_GDI:
            lpDest->tymed = TYMED_GDI;
            lpDest->hGlobal = NULL;
            break;

        case TYMED_FILE:
        {
            lpDest->tymed = TYMED_FILE;
            if (lpSource->lpszFileName == NULL)
            {
                //AfxThrowInvalidArgException();
            }
            UINT cbSrc = static_cast<UINT>(wcslen(lpSource->lpszFileName));
            LPOLESTR szFileName = 
                (LPOLESTR)::CoTaskMemAlloc((cbSrc + 1)* sizeof(OLECHAR));
            lpDest->lpszFileName = szFileName;
            if (szFileName == NULL)
                return FALSE;

            memcpy_s(szFileName, (cbSrc + 1) * sizeof(OLECHAR),
                lpSource->lpszFileName, (cbSrc + 1) * sizeof(OLECHAR));
            return TRUE;
        }

        // unable to create + copy other TYMEDs
        default:
            return FALSE;
        }
    }

    switch (lpSource->tymed)
    {
    case TYMED_HGLOBAL:
    {
        HGLOBAL hDest = ODSCopyGlobalMemory(lpDest->hGlobal,
            lpSource->hGlobal);
        if (hDest == NULL)
            return FALSE;

        lpDest->hGlobal = hDest;
    }
    return TRUE;

    case TYMED_ISTREAM:
    {

        // get the size of the source stream
        STATSTG stat;
        if (lpSource->pstm->Stat(&stat, STATFLAG_NONAME) != S_OK)
        {
            // unable to get size of source stream
            return FALSE;
        }

        // always seek to zero before copy
        LARGE_INTEGER zero = { 0, 0 };
        lpDest->pstm->Seek(zero, STREAM_SEEK_SET, NULL);
        lpSource->pstm->Seek(zero, STREAM_SEEK_SET, NULL);

        // copy source to destination
        if (lpSource->pstm->CopyTo(lpDest->pstm, stat.cbSize,
            NULL, NULL) != NULL)
        {
            // copy from source to dest failed
            return FALSE;
        }

        // always seek to zero after copy
        lpDest->pstm->Seek(zero, STREAM_SEEK_SET, NULL);
        lpSource->pstm->Seek(zero, STREAM_SEEK_SET, NULL);
    }
    return TRUE;

    case TYMED_ISTORAGE:
    {
        // just copy source to destination
        if (lpSource->pstg->CopyTo(0, NULL, NULL, lpDest->pstg) != S_OK)
            return FALSE;
    }
    return TRUE;

    case TYMED_FILE:
    {
        //CString strSource(lpSource->lpszFileName);
        //CString strDest(lpDest->lpszFileName);
        //return CopyFile(lpSource->lpszFileName ? strSource.GetString() : NULL,
        //    lpDest->lpszFileName ? strDest.GetString() : NULL, FALSE);
        return CopyFile(lpSource->lpszFileName ,
            lpDest->lpszFileName, FALSE);
    }


    case TYMED_ENHMF:
    case TYMED_GDI:
    {
        // with TYMED_GDI cannot copy into existing HANDLE
        if (lpDest->hGlobal != NULL)
            return FALSE;

        // otherwise, use OleDuplicateData for the copy
        lpDest->hGlobal = OleDuplicateData(lpSource->hGlobal, cfFormat, 0);
        if (lpDest->hGlobal == NULL)
            return FALSE;
    }
    return TRUE;

    // other TYMEDs cannot be copied
    default:
        return FALSE;
    }
}




COleDataSource* COleDataSource::CurrentClipboardSource = nullptr;

HRESULT __stdcall COleDataSource::QueryInterface(REFIID iid, void** ppvObject)
{
    if (iid == IID_IDataObject)
    {
        this->AddRef();
        *ppvObject = static_cast<IDataObject*>(this);
        return S_OK;
    }
    else if (iid == IID_IUnknown)
    {
        this->AddRef();
        *ppvObject = static_cast<IUnknown*>(this);
        return S_OK;
    }
    //Only provide myself
    return (DWORD)E_NOINTERFACE;
}

ULONG __stdcall COleDataSource::AddRef(void)
{
    return InterlockedIncrement(&this->m_lRefCount);
}

ULONG __stdcall COleDataSource::Release(void)
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
BOOL COleDataSource::FlushClipboard()
{
    if (GetClipboardOwner() != nullptr)
    {
        // active clipboard source and it is on the clipboard - flush it
        ::OleFlushClipboard();

        // shouldn't be clipboard owner any more...
    }
    return (GetClipboardOwner() == nullptr);
}

COleDataSource* COleDataSource::GetClipboardOwner()
{
    COleDataSource* ccp = CurrentClipboardSource;
    if (ccp == nullptr)
        return nullptr;    // can't own the clipboard if pClipboardSource isn't set

    if (::OleIsCurrentClipboard(ccp) != S_OK)
    {
        return nullptr;    // don't own the clipboard anymore
    }

    // return current clipboard source
    return ccp;
}

COleDataSource::COleDataSource()
    : m_lRefCount()
    , m_pDataCache(0)
    , m_nMaxSize(0)
    , m_nSize(0)
    , m_nGrowBy(10)
{
    this->m_lRefCount = 1;
    this->m_pDataCache = nullptr;
    this->m_nMaxSize = 0;
    this->m_nSize = 0;
    this->m_nGrowBy = 10;
}

COleDataSource::~COleDataSource()
{
    // clear clipboard source if this object was on the clipboard
    COleDataSource* ccp = CurrentClipboardSource;
    if (this == ccp)
        CurrentClipboardSource = nullptr;

    // free the clipboard data cache
    this->Empty();
}

void COleDataSource::OnFinalRelease()
{
    delete this;
}

ODS_DATACACHE_ENTRY* COleDataSource::Lookup(LPFORMATETC lpFormatEtc, DATADIR nDataDir) const
{
    ODS_DATACACHE_ENTRY* pLast = nullptr;

    // look for suitable match to lpFormatEtc in cache
    for (UINT nIndex = 0; nIndex < m_nSize; nIndex++)
    {
        // get entry from cache at nIndex
        ODS_DATACACHE_ENTRY* pCache = &m_pDataCache[nIndex];
        FORMATETC* pCacheFormat = &pCache->m_formatEtc;

        // check for match
        // we threat lindex == 0 and lindex == -1 as equivalent
        if (pCacheFormat->cfFormat == lpFormatEtc->cfFormat &&
            (pCacheFormat->tymed & lpFormatEtc->tymed) != 0 &&
            (pCacheFormat->dwAspect == DVASPECT_THUMBNAIL ||
                pCacheFormat->dwAspect == DVASPECT_ICON ||
                pCache->m_stgMedium.tymed == TYMED_NULL ||
                pCacheFormat->lindex == lpFormatEtc->lindex ||
                (pCacheFormat->lindex == 0 && lpFormatEtc->lindex == -1) ||
                (pCacheFormat->lindex == -1 && lpFormatEtc->lindex == 0)) &&
            pCacheFormat->dwAspect == lpFormatEtc->dwAspect &&
            pCache->m_nDataDir == nDataDir)
        {
            // for backward compatibility we match even if we never
            // find an exact match for the DVTARGETDEVICE
            pLast = pCache;
            DVTARGETDEVICE* ptd1 = pCacheFormat->ptd;
            DVTARGETDEVICE* ptd2 = lpFormatEtc->ptd;

            if (ptd1 == nullptr && ptd2 == nullptr)
            {
                // both target devices are nullptr (exact match), so return it
                break;
            }
            if (ptd1 != nullptr && ptd2 != nullptr &&
                ptd1->tdSize == ptd2->tdSize &&
                memcmp(ptd1, ptd2, ptd1->tdSize) == 0)
            {
                // exact match, so return it
                break;
            }
            // continue looking for better match
        }
    }

    return pLast;    // not found
}

ODS_DATACACHE_ENTRY* COleDataSource::GetCacheEntry(LPFORMATETC lpFormatEtc, DATADIR nDataDir)
{
    ODS_DATACACHE_ENTRY* pEntry = Lookup(lpFormatEtc, nDataDir);
    if (pEntry != nullptr)
    {
        // cleanup current entry and return it
        CoTaskMemFree(pEntry->m_formatEtc.ptd);
        ::ReleaseStgMedium(&pEntry->m_stgMedium);
    }
    else
    {
        // allocate space for item at m_nSize (at least room for 1 item)
        if (m_pDataCache == nullptr || m_nSize == m_nMaxSize)
        {
            //ASSERT(m_nGrowBy != 0);
            ODS_DATACACHE_ENTRY* pCache = new ODS_DATACACHE_ENTRY[m_nMaxSize + m_nGrowBy];
            m_nMaxSize += m_nGrowBy;
            if (m_pDataCache != nullptr)
            {
                memcpy_s(pCache, (m_nMaxSize + m_nGrowBy) * sizeof(ODS_DATACACHE_ENTRY),
                    m_pDataCache, m_nSize * sizeof(ODS_DATACACHE_ENTRY));
                delete[] m_pDataCache;
            }
            m_pDataCache = pCache;
        }
        //ASSERT(m_pDataCache != nullptr);
        //ASSERT(m_nMaxSize != 0);

        pEntry = &m_pDataCache[m_nSize++];
    }

    // fill the cache entry with the format and data direction and return it
    pEntry->m_nDataDir = nDataDir;
    pEntry->m_formatEtc = *lpFormatEtc;
    return pEntry;
}

void COleDataSource::Empty()
{
    if (m_pDataCache != nullptr && m_nMaxSize != 0 && m_nSize != 0)
    {
        // release all of the STGMEDIUMs and FORMATETCs
        for (UINT nIndex = 0; nIndex < m_nSize; nIndex++)
        {
            CoTaskMemFree(m_pDataCache[nIndex].m_formatEtc.ptd);
            ::ReleaseStgMedium(&m_pDataCache[nIndex].m_stgMedium);
        }

        // delete the cache
        delete[] m_pDataCache;
        m_pDataCache = nullptr;
        m_nMaxSize = 0;
        m_nSize = 0;
    }
}

BOOL COleDataSource::SetClipboard()
{
    // attempt OLE set clipboard operation
    LPDATAOBJECT lpDataObject = this;
    SCODE sc = ::OleSetClipboard(lpDataObject);
    if (sc != S_OK)
        return FALSE;

    // success - set as current clipboard source
    CurrentClipboardSource = this;
    BOOL ret = TRUE;
    if (::OleIsCurrentClipboard(lpDataObject) != S_OK)
    {
        ret = FALSE;
    }
    this->Release();
    return ret;
}

void COleDataSource::CacheGlobalData(CLIPFORMAT cfFormat, HGLOBAL hGlobal, LPFORMATETC lpFormatEtc)
{
    // fill in FORMATETC struct
    FORMATETC formatEtc = { 0 };
    lpFormatEtc = FillFormatEtcHelper(lpFormatEtc, cfFormat, &formatEtc);
    lpFormatEtc->tymed = TYMED_HGLOBAL;
    // add it to the cache
    ODS_DATACACHE_ENTRY* pEntry = GetCacheEntry(lpFormatEtc, DATADIR_GET);
    pEntry->m_stgMedium.tymed = TYMED_HGLOBAL;
    pEntry->m_stgMedium.hGlobal = hGlobal;
    pEntry->m_stgMedium.pUnkForRelease = nullptr;
}

void COleDataSource::DelayRenderFileData(CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc)
{
    // fill in FORMATETC struct
    FORMATETC formatEtc = { 0 };
    lpFormatEtc = FillFormatEtcHelper(lpFormatEtc, cfFormat, &formatEtc);
    lpFormatEtc->tymed = TYMED_ISTREAM | TYMED_HGLOBAL;

    // add it to the cache
    ODS_DATACACHE_ENTRY* pEntry = GetCacheEntry(lpFormatEtc, DATADIR_GET);
    pEntry->m_stgMedium.tymed = TYMED_NULL;
    pEntry->m_stgMedium.hGlobal = nullptr;
    pEntry->m_stgMedium.pUnkForRelease = nullptr;
}

DROPEFFECT COleDataSource::DoDragDrop(DWORD dwEffects, LPCRECT lpRectStartDrag, COleDropSource* pDropSource)
{
    // use standard drop source implementation if one not provided
    COleDropSource dropSource;
    if (pDropSource == nullptr)
        pDropSource = &dropSource;

    // call global OLE api to do the drag drop
    DWORD dwResultEffect = DROPEFFECT_NONE;
    ::DoDragDrop(this, pDropSource, dwEffects, &dwResultEffect);
    return dwResultEffect;
}

void COleDataSource::CacheData(CLIPFORMAT cfFormat, LPSTGMEDIUM lpStgMedium, LPFORMATETC lpFormatEtc)
{
    // fill in FORMATETC struct
    FORMATETC formatEtc = { 0 };
    lpFormatEtc = FillFormatEtcHelper(lpFormatEtc, cfFormat, &formatEtc);
    lpFormatEtc->tymed = lpStgMedium->tymed;

    // add it to the cache
    ODS_DATACACHE_ENTRY* pEntry = GetCacheEntry(lpFormatEtc, DATADIR_GET);
    pEntry->m_stgMedium = *lpStgMedium;
}

void COleDataSource::DelayRenderData(CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc)
{
    FORMATETC formatEtc = { 0 };
    if (lpFormatEtc == nullptr)
    {
        lpFormatEtc = FillFormatEtcHelper(lpFormatEtc, cfFormat, &formatEtc);
        lpFormatEtc->tymed = TYMED_HGLOBAL;
    }
    if (cfFormat != 0) {
        lpFormatEtc->cfFormat = cfFormat;
    }
    // add it to the cache
    ODS_DATACACHE_ENTRY* pEntry = GetCacheEntry(lpFormatEtc, DATADIR_GET);
    memset(&pEntry->m_stgMedium, 0, sizeof pEntry->m_stgMedium);
}

void COleDataSource::DelaySetData(CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc)
{
    FORMATETC formatEtc = { 0 };
    lpFormatEtc = FillFormatEtcHelper(lpFormatEtc, cfFormat, &formatEtc);

    // add it to the cache
    ODS_DATACACHE_ENTRY* pEntry = GetCacheEntry(lpFormatEtc, DATADIR_SET);
    pEntry->m_stgMedium.tymed = TYMED_NULL;
    pEntry->m_stgMedium.hGlobal = nullptr;
    pEntry->m_stgMedium.pUnkForRelease = nullptr;

}

BOOL COleDataSource::OnRenderGlobalData(LPFORMATETC lpFormatEtc, HGLOBAL* phGlobal)
{
    return FALSE;
}

BOOL COleDataSource::OnRenderFileData(LPFORMATETC lpFormatEtc, CSharedFile* file)
{
    return FALSE;
}

BOOL COleDataSource::OnRenderFileData(LPFORMATETC lpFormatEtc, COleStreamFile* file)
{
    return FALSE;
}

BOOL COleDataSource::OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
    //NOTICE: the following code is commented out
    //NOTICE: because we don't use the logic when we use IStream logic
#if 0
    if (lpFormatEtc->tymed & TYMED_HGLOBAL)
    {
        HGLOBAL hGlobal = lpStgMedium->hGlobal;
        if (OnRenderGlobalData(lpFormatEtc, &hGlobal))
        {
            lpStgMedium->tymed = TYMED_HGLOBAL;
            lpStgMedium->hGlobal = hGlobal;
            return TRUE;
        }

        CSharedFile file;
        if (lpStgMedium->tymed == TYMED_HGLOBAL)
        {
            file.SetHandle(lpStgMedium->hGlobal,FALSE);
        }
        if (OnRenderFileData(lpFormatEtc, &file))
        {
            lpStgMedium->tymed = TYMED_HGLOBAL;
            lpStgMedium->hGlobal = file.Detach();
            return TRUE;
        }
        if (lpStgMedium->tymed == TYMED_HGLOBAL)
        {
            file.Detach();
        }
    }
    // attempt TYMED_ISTREAM format
    if (lpFormatEtc->tymed & TYMED_ISTREAM)
    {
        COleStreamFile file;
        if (lpStgMedium->tymed == TYMED_ISTREAM)
        {
            file.Attach(lpStgMedium->pstm);
        }
        else
        {
            if (!file.CreateMemoryStream()) {
                return FALSE;
            }
        }
        // get data into the stream
        if (OnRenderFileData(lpFormatEtc, &file))
        {
            lpStgMedium->tymed = TYMED_ISTREAM;
            lpStgMedium->pstm = file.Detach();
            return TRUE;
        }
        if (lpStgMedium->tymed == TYMED_ISTREAM)
        {
            file.Detach();
        }
    }
#endif

    return FALSE;   // default does nothing
}

BOOL COleDataSource::OnSetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium, BOOL bRelease)
{
    return FALSE;   // default does nothing
}

HRESULT __stdcall COleDataSource::GetData(FORMATETC* lpFormatEtc, STGMEDIUM* lpStgMedium)
{
    if (lpFormatEtc == NULL || lpStgMedium == NULL)
    {
        return E_INVALIDARG;
    }
    ODS_DATACACHE_ENTRY* pCache = this->Lookup(lpFormatEtc, DATADIR_GET);
    if (pCache == NULL)
    {
        return DATA_E_FORMATETC;
    }
    // use cache if entry is not delay render
    memset(lpStgMedium, 0, sizeof(STGMEDIUM));
    if (pCache->m_stgMedium.tymed != TYMED_NULL)
    {
        // Copy the cached medium into the lpStgMedium provided by caller.
        if (!ODSCopyStgMedium(lpFormatEtc->cfFormat, lpStgMedium, &pCache->m_stgMedium))
        {
            return DATA_E_FORMATETC;
        }

        // format was supported for copying
        return S_OK;
    }

    SCODE sc = DATA_E_FORMATETC;
    try
    {
        // attempt LPSTGMEDIUM based delay render
        if (this->OnRenderData(lpFormatEtc, lpStgMedium))
        {
            sc = S_OK;
        }
    }
    catch(...)
    {
        sc = E_FAIL;
    }

    return sc;
}

HRESULT __stdcall COleDataSource::GetDataHere(FORMATETC* lpFormatEtc, STGMEDIUM* lpStgMedium)
{
    if (lpFormatEtc == NULL || lpStgMedium == NULL)
    {
        return E_INVALIDARG;
    }

    // these two must be the same
    lpFormatEtc->tymed = lpStgMedium->tymed;    // but just in case...

    // attempt to find match in the cache
    ODS_DATACACHE_ENTRY* pCache = this->Lookup(lpFormatEtc, DATADIR_GET);
    if (pCache == NULL)
    {
        return DATA_E_FORMATETC;
    }

    // handle cached medium and copy
    if (pCache->m_stgMedium.tymed != TYMED_NULL)
    {
        // found a cached format -- copy it to dest medium
        if (!ODSCopyStgMedium(lpFormatEtc->cfFormat, lpStgMedium, &pCache->m_stgMedium))
        {
            return DATA_E_FORMATETC;
        }

        // format was supported for copying
        return S_OK;
    }

    SCODE sc = DATA_E_FORMATETC;
    try
    {
        // attempt LPSTGMEDIUM based delay render
        if (this->OnRenderData(lpFormatEtc, lpStgMedium))
        {
            sc = S_OK;
        }
    }
    catch (...)
    {
        sc = E_FAIL;
    }

    return sc;
}

HRESULT __stdcall COleDataSource::QueryGetData(FORMATETC* lpFormatEtc)
{
    if (lpFormatEtc == NULL)
    {
        return E_INVALIDARG;
    }

    // attempt to find match in the cache
    ODS_DATACACHE_ENTRY* pCache = this->Lookup(lpFormatEtc, DATADIR_GET);
    if (pCache == NULL)
    {
        return DATA_E_FORMATETC;
    }

    // it was found in the cache or can be rendered -- success
    return S_OK;
}

HRESULT __stdcall COleDataSource::GetCanonicalFormatEtc(FORMATETC* pformatectIn, FORMATETC* pformatetcOut)
{
    return DATA_S_SAMEFORMATETC;
}

HRESULT __stdcall COleDataSource::SetData(FORMATETC* lpFormatEtc, STGMEDIUM* lpStgMedium, BOOL bRelease)
{
    if (lpFormatEtc == NULL || lpStgMedium == NULL)
    {
        return E_INVALIDARG;
    }

    // attempt to find match in the cache
    ODS_DATACACHE_ENTRY* pCache = this->Lookup(lpFormatEtc, DATADIR_SET);
    if (pCache == NULL)
    {
        return DATA_E_FORMATETC;
    }


    SCODE sc = E_UNEXPECTED;
    try
    {
        // attempt LPSTGMEDIUM based SetData
        if (this->OnSetData(lpFormatEtc, lpStgMedium, bRelease))
        {
            sc = S_OK;
        }
    }
    catch (...)
    {
        sc = E_FAIL;
    }

    return sc;
}

HRESULT __stdcall COleDataSource::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc)
{
    if (ppenumFormatEtc == NULL)
    {
        return E_POINTER;
    }

    *ppenumFormatEtc = NULL;

    CEnumFormatEtc* pFormatList = NULL;
    SCODE sc = E_OUTOFMEMORY;
    try
    {
        // generate a format list from the cache
        pFormatList = new CEnumFormatEtc;
        for (UINT nIndex = 0; nIndex < this->m_nSize; nIndex++)
        {
            ODS_DATACACHE_ENTRY* pCache = &this->m_pDataCache[nIndex];
            if ((DWORD)pCache->m_nDataDir & dwDirection)
            {
                // entry should be enumerated -- add it to the list
                FORMATETC formatEtc;
                ODSOleCopyFormatEtc(&formatEtc, &pCache->m_formatEtc);
                pFormatList->AddFormat(&formatEtc);
            }
        }

        // give it away to OLE (ref count is already 1)
        *ppenumFormatEtc = (LPENUMFORMATETC)pFormatList;
        sc = S_OK;
    }
    catch (...) {

    }
    return sc;
}


HRESULT __stdcall COleDataSource::DAdvise(FORMATETC* pFormatEtc, DWORD advf, IAdviseSink*, DWORD*)
{
    return OLE_E_ADVISENOTSUPPORTED;
}

HRESULT __stdcall COleDataSource::DUnadvise(DWORD dwConnection)
{
    return OLE_E_ADVISENOTSUPPORTED;
}

HRESULT __stdcall COleDataSource::EnumDAdvise(IEnumSTATDATA** ppEnumAdvise)
{
    ppEnumAdvise = nullptr;
    return OLE_E_ADVISENOTSUPPORTED;
}
