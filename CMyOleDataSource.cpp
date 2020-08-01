#include <ShlObj.h>
#include <stdio.h>
#include <tchar.h>
#include "CMyOleDataSource.h"
#include "CMyOleStream.h"

CMyOleDataSource::CMyOleDataSource()
	:m_nFiles(0)
	,m_sPaths()
{
}

CMyOleDataSource::~CMyOleDataSource()
{
	if (this->m_sPaths != nullptr) {
		for (int i = 0; i < this->m_nFiles; i++) {
			delete[] this->m_sPaths[i];
		}
		delete[] this->m_sPaths;
		this->m_sPaths = nullptr;
	}
	this->m_nFiles = 0;
}

BOOL CMyOleDataSource::OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium)
{
	BOOL ret = COleDataSource::OnRenderData(lpFormatEtc, lpStgMedium);
	if (!ret) {
		if (lpFormatEtc->cfFormat ==
			RegisterClipboardFormat(CFSTR_FILECONTENTS) 
			&& lpFormatEtc->lindex >= 0)
		{
			TCHAR* src = this->m_sPaths[lpFormatEtc->lindex];

			CMyOleStream* cmf = new CMyOleStream(src);
			
			lpStgMedium->tymed = TYMED_ISTREAM;
		 	
			lpStgMedium->pstm = cmf;

			lpStgMedium->pUnkForRelease = cmf;
			//ref-ed twice, so use one more AddRef
			cmf->AddRef();

			ret = TRUE;
		}
	}
	return ret;
}
 