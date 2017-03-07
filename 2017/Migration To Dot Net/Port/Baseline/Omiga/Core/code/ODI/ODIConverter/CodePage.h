///////////////////////////////////////////////////////////////////////////////
//	FILE:			CodePage.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		28/09/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_CODEPAGE_H__B9149005_FCF9_4828_96CA_A14BE0182978__INCLUDED_)
#define AFX_CODEPAGE_H__B9149005_FCF9_4828_96CA_A14BE0182978__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CCodePage
{
public:
	enum ECodePage
	{
		// Some example code pages.
		// See Control Panel, Regional Options, General, Advanced for additional Ebcdic code page ids.
		cpEbcdicUSCanada			= 37,
		cpEbcdicInternational		= 500,
		cpEbcdicUSCanadaEuro		= 1140,
		cpEbcdicUKEuro				= 1146,
		cpEbcdicFranceEuro			= 1147,
		cpEbcdicInternationalEuro	= 1148,
		cpEbcdicUK					= 20285,
		cpEbcdicFrance				= 20297
	};

public:
	CCodePage(
		LPCWSTR szCountry = L"", 
		LPCWSTR szLanguage = L"", 
		bool bEuro = false, 
		int nMSCodePage = 0, 
		int nIBMCCSID = 0, 
		bool bDefault = false,
		LPCWSTR szTranslationTable = L"");
	virtual ~CCodePage();
	void SetCountry(LPCWSTR szCountry) { m_bstrCountry = szCountry; }
	_bstr_t GetCountry() const { return m_bstrCountry; }
	void SetLanguage(LPCWSTR szLanguage) { m_bstrLanguage = szLanguage; }
	_bstr_t GetLanguage() const { return m_bstrLanguage; }
	void SetEuro(bool bEuro) { m_bEuro = bEuro; }
	bool GetEuro() const { return m_bEuro; }
	void SetMSCodePage(const int nMSCodePage) { m_nMSCodePage = nMSCodePage; }
	int GetMSCodePage() const { return m_nMSCodePage; }
	void SetIBMCCSID(const int nIBMCCSID) { m_nIBMCCSID = nIBMCCSID; }
	int GetIBMCCSID() const { return m_nIBMCCSID; }
	void SetDefault(bool bDefault) { m_bDefault = bDefault; }
	bool GetDefault() const { return m_bDefault; }
	void SetTranslationTable(LPCWSTR szTranslationTable);
	unsigned short* GetTranslationTableSend() const { return m_pTranslationTableSend; }
	unsigned short* GetTranslationTableRecv() const { return m_pTranslationTableRecv; }
	int GetTranslationTableSize() const { return m_nTranslationTableSize; }

	// For all these conversion functions, the caller is responsible for ensuring that the
	// output parameter is big enough for the converted data, as follows:
	//	Source		Target		Target size (in bytes)
	//	wchar_t		wchar_t		wcslen(Source) + 2
	//	wchar_t		char		(wcslen(Source) / 2) + 1
	//	char		wchar_t		(strlen(Source) * 2) + 2
	//	char		char		strlen(Source) + 1
	// Assumes lpWide is null terminated, and lpEbcdic can accomodate converted Unicode string. 
	int SendW2MBC(LPCWSTR lpWide, LPSTR lpMBC);
	// Does not assume lpWide is null terminated. lpEbcdic must accomodate converted Unicode string. 
	int SendW2MBC(LPCWSTR lpWide, int nMaxSizeWide, LPSTR lpMBC, int nMaxSizeMBC);
	// Assumes lpEbcdic is null terminated, and lpWide can accomodate converted EBCDIC string.
	int RecvMBC2W(LPCSTR lpMBC, LPWSTR lpWide);
	// Does not assume lpEbcdic is null terminated. lpWide must accomodate converted EBCDIC string. 
	int RecvMBC2W(LPCSTR lpMBC, int nMaxSizeMBC, LPWSTR lpWide, int nMaxSizeWide);

	static int GetCodePageTable(int nCodePage, unsigned short* CodePageTable, int nMaxSizeCodePageTable);
	void Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode);

private:
	int CalculateTranslationTableSize(LPCWSTR szTranslationTable) const;
	void InitTranslationTables();
	int Translate(
		const BYTE* pbySrc, 
		const int nSrcLen,
		BYTE* pbyTgt, 
		const int nTgtLen,
		const unsigned short* TransTable, 
		const int nMaxSizeTransTable,
		const BOOL bSrcW = TRUE,
		const BOOL bTgtW = TRUE);

	_bstr_t				m_bstrCountry;				// Country code as used by Optimus.
	_bstr_t				m_bstrLanguage;				// Language code as used by Optimus.
	bool				m_bEuro;					// If true, then code page supports Euro symbol.
	int					m_nMSCodePage;				// Microsoft code page number.
	int					m_nIBMCCSID;				// IBM CCS ID.
	bool				m_bDefault;					// If true, then this is the default code page.
	_bstr_t				m_bstrTranslationTable;		// Translation table text.
	unsigned short*		m_pTranslationTableSend;	// Translation table for sent data.
	unsigned short*		m_pTranslationTableRecv;	// Translation table for received data.
	int					m_nTranslationTableSize;	// Size of translation tables.
};

#endif // !defined(AFX_CODEPAGE_H__B9149005_FCF9_4828_96CA_A14BE0182978__INCLUDED_)
