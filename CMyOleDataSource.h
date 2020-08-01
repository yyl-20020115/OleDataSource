#pragma once
#include "COleDataSource.h"

class CMyOleDataSource 
	: public COleDataSource
{
public:
	CMyOleDataSource();
	virtual ~CMyOleDataSource();
public:
	virtual BOOL OnRenderData(LPFORMATETC lpFormatEtc, LPSTGMEDIUM lpStgMedium);

public:

	TCHAR** m_sPaths;
	int m_nFiles;
};

