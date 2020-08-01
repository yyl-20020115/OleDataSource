#pragma once
#include <windows.h>
#include "COleDropSource.h"
#include "CFile.h"
#include "CSharedFile.h"
#include "COleStreamFile.h"

struct ODS_DATACACHE_ENTRY
{
	FORMATETC m_formatEtc;
	STGMEDIUM m_stgMedium;
	DATADIR m_nDataDir;
};

class COleDataSource 
	: public IDataObject
{
public:
	static COleDataSource* CurrentClipboardSource;

public:

	// IUnknown members
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG   __stdcall AddRef(void);
	ULONG   __stdcall Release(void);

public:
	static BOOL PASCAL FlushClipboard();
	static COleDataSource* PASCAL GetClipboardOwner();

public:
	COleDataSource();
	virtual ~COleDataSource();
	virtual void OnFinalRelease();

public:
	void Empty();
	BOOL SetClipboard();
	void CacheGlobalData(CLIPFORMAT cfFormat, HGLOBAL hGlobal,
		LPFORMATETC lpFormatEtc = NULL);    // for HGLOBAL based data
	void DelayRenderFileData(CLIPFORMAT cfFormat,
		LPFORMATETC lpFormatEtc = NULL);    // for CFile* based delayed render

	DROPEFFECT DoDragDrop(
		DWORD dwEffects = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK,
		LPCRECT lpRectStartDrag = NULL,
		COleDropSource* pDropSource = NULL);

	void CacheData(CLIPFORMAT cfFormat, LPSTGMEDIUM lpStgMedium,
		LPFORMATETC lpFormatEtc = NULL);    // for LPSTGMEDIUM based data

	void DelayRenderData(CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc = NULL);

	void DelaySetData(CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc = NULL);

public:
	// Overidables
	virtual BOOL OnRenderGlobalData(LPFORMATETC lpFormatEtc, HGLOBAL* phGlobal);
	virtual BOOL OnRenderFileData(LPFORMATETC lpFormatEtc, CSharedFile* file);
	virtual BOOL OnRenderFileData(LPFORMATETC lpFormatEtc, COleStreamFile* file);
	virtual BOOL OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium);
	virtual BOOL OnSetData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium,
		BOOL bRelease);

public:
	// IDataObject members
	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetData(
		/* [annotation][unique][in] */
		_In_  FORMATETC* pformatetcIn,
		/* [annotation][out] */
		_Out_  STGMEDIUM* pmedium);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDataHere(
		/* [annotation][unique][in] */
		_In_  FORMATETC* pformatetc,
		/* [annotation][out][in] */
		_Inout_  STGMEDIUM* pmedium);

	virtual HRESULT STDMETHODCALLTYPE QueryGetData(
		/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc);

	virtual HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(
		/* [unique][in] */ __RPC__in_opt FORMATETC* pformatectIn,
		/* [out] */ __RPC__out FORMATETC* pformatetcOut);

	virtual /* [local] */ HRESULT STDMETHODCALLTYPE SetData(
		/* [annotation][unique][in] */
		_In_  FORMATETC* pformatetc,
		/* [annotation][unique][in] */
		_In_  STGMEDIUM* pmedium,
		/* [in] */ BOOL fRelease);

	virtual HRESULT STDMETHODCALLTYPE EnumFormatEtc(
		/* [in] */ DWORD dwDirection,
		/* [out] */ __RPC__deref_out_opt IEnumFORMATETC** ppenumFormatEtc);

	virtual HRESULT STDMETHODCALLTYPE DAdvise(
		/* [in] */ __RPC__in FORMATETC* pformatetc,
		/* [in] */ DWORD advf,
		/* [unique][in] */ __RPC__in_opt IAdviseSink* pAdvSink,
		/* [out] */ __RPC__out DWORD* pdwConnection);

	virtual HRESULT STDMETHODCALLTYPE DUnadvise(
		/* [in] */ DWORD dwConnection);

	virtual HRESULT STDMETHODCALLTYPE EnumDAdvise(
		/* [out] */ __RPC__deref_out_opt IEnumSTATDATA** ppenumAdvise);

private:

	LONG m_lRefCount;

	ODS_DATACACHE_ENTRY* m_pDataCache;  // data cache itself
	UINT m_nMaxSize;    // current allocated size
	UINT m_nSize;       // current size of the cache
	UINT m_nGrowBy;     // number of cache elements to grow by for new allocs

	ODS_DATACACHE_ENTRY* Lookup(
		LPFORMATETC lpFormatEtc, DATADIR nDataDir) const;
	ODS_DATACACHE_ENTRY* GetCacheEntry(
		LPFORMATETC lpFormatEtc, DATADIR nDataDir);

};


