#pragma once
#include <Windows.h>

typedef DWORD DROPEFFECT;

class COleDropSource : IDropSource
{
public:

	// IUnknown members
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject);
	ULONG   __stdcall AddRef(void);
	ULONG   __stdcall Release(void);
public:
	virtual void OnFinalRelease();

public:
	COleDropSource();
	virtual ~COleDropSource();

public:
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(
		/* [annotation][in] */
		_In_  BOOL fEscapePressed,
		/* [annotation][in] */
		_In_  DWORD grfKeyState);

	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(
		/* [annotation][in] */
		_In_  DWORD dwEffect);

private:
	LONG m_lRefCount;
public:

	friend class COleDataSource;

};

