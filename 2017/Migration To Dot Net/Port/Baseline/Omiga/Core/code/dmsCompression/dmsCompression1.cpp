///////////////////////////////////////////////////////////////////////////////
//	FILE:			DMSCOMPRESSION1.CPP
//	DESCRIPTION: 	
//
//	SYSTEM:	    	DMS
//	COPYRIGHT:		(c) 2002, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	LD		23/04/02	First version
//	LD		14/05/02	Set decompression growby size to decompressed size * 3
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DmsCompression.h"
#include "dmsCompression1.h"
#include "Zipper.h"
#include "SafeArrayAccessUnaccessData.h"

#if defined(MSXML4)
// no_function_mapping is required to prevent "LNK2001: unresolved external symbol" errors with MSXML4.
// no_function_mapping is an undocumented attribute, that removes the "#pragma implementation_key" 
// (another undocumented attribute) from the tli file. no_function_mapping is not documented in MSDN, 
// but you can find references to it by searching Google groups.
#import "msxml4.dll" rename_namespace("MSXML") no_function_mapping
#ifdef FREETHREADEDDOMDOCUMENT
#define DOMDocument FreeThreadedDOMDocument40
#else
#define DOMDocument DOMDocument40
#endif
#else
#import "msxml3.dll" rename_namespace("MSXML")
#ifdef FREETHREADEDDOMDOCUMENT
#define DOMDocument FreeThreadedDOMDocument
#else
#define DOMDocument DOMDocument
#endif
#endif


/////////////////////////////////////////////////////////////////////////////
// CdmsCompression1
CdmsCompression1::CdmsCompression1() :
	m_bstrCompressionMethod(L"COMPAPI")
{
}

STDMETHODIMP CdmsCompression1::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IdmsCompression1
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
#ifdef VC6
		if (::ATL::InlineIsEqualGUID(*arr[i],riid))
#else
		if (InlineIsEqualGUID(*arr[i],riid))
#endif
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CdmsCompression1::get_CompressionMethod(BSTR* pbstrCompressionMethod)
{
	*pbstrCompressionMethod = m_bstrCompressionMethod.copy();
	return S_OK;
}

STDMETHODIMP CdmsCompression1::put_CompressionMethod(BSTR bstrCompressionMethod)
{
	m_bstrCompressionMethod = bstrCompressionMethod;
	return S_OK;
}

STDMETHODIMP CdmsCompression1::BSTRCompressToSafeArray(BSTR bstrIn, VARIANT *pvarOut)
{
	HRESULT hr = E_FAIL;
	VariantInit(pvarOut);

	try
	{
		// extract source
		long lSrcMaxLen = ::SysStringByteLen(bstrIn);
		char* pSrcBuffer = (LPSTR)bstrIn;

		// zip
		CAutoBuffer AutoBuffer;
		AutoBuffer.SetGrowBy(lSrcMaxLen);

		CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

		if (pZipper->ZipBufferToBuffer(&AutoBuffer, pSrcBuffer, lSrcMaxLen, NULL))
		{
			// return data
			char* pTgtBuffer = AutoBuffer.GetData();
			ULONG uLength = AutoBuffer.GetDataSize();

			//Set the type to an array of unsigned chars (OLE SAFEARRAY)
			pvarOut->vt = VT_ARRAY | VT_UI1;
			SAFEARRAYBOUND  rgsabound[1];
			rgsabound[0].cElements = uLength;
			rgsabound[0].lLbound = 0;
			pvarOut->parray = SafeArrayCreate(VT_UI1,1,rgsabound);
			if(pvarOut->parray != NULL)
			{	
				void* pArrayData = NULL;
				CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
				SafeArrayAccessUnaccessData.Access(pvarOut->parray, &pArrayData);
				memcpy(pArrayData, pTgtBuffer, uLength);
				hr = S_OK;
			}
			else
			{
				Error(L"Unable to create safe array");
			}
		}
		else
		{
			Error(L"Unable to zip data");
		}

		delete pZipper;
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}


STDMETHODIMP CdmsCompression1::SafeArrayDecompressToBSTR(VARIANT varIn, BSTR *pbstrOut)
{
	HRESULT hr = E_FAIL;
	*pbstrOut = NULL;

	enum
	{
		VALUETYPE_SAFEARRAYBYTES,
		VALUETYPE_BSTR,
		VALUETYPE_NOTSUPPORTED
	} eValueType = VALUETYPE_NOTSUPPORTED;

	try
	{
		// extract source
		long lSrcMaxLen = 0;
		char* pSrcBuffer = NULL;
		LONG lLBound = 0;
		LONG lUBound = 0;
		CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
		if (V_VT(&varIn) == (VT_ARRAY | VT_UI1) && // array of bytes
			SafeArrayGetDim(V_ARRAY(&varIn)) == 1 && // one dimensional 
			SUCCEEDED(SafeArrayGetLBound(V_ARRAY(&varIn), 1, &lLBound)) &&
			SUCCEEDED(SafeArrayGetUBound(V_ARRAY(&varIn), 1, &lUBound)) &&
			SUCCEEDED(SafeArrayAccessUnaccessData.Access(V_ARRAY(&varIn), reinterpret_cast<void**>(&pSrcBuffer))))
		{
			eValueType = VALUETYPE_SAFEARRAYBYTES;
			lSrcMaxLen = lUBound - lLBound + 1;
		}
		else if (V_VT(&varIn) == VT_BSTR)
		{
			lSrcMaxLen = ::SysStringByteLen(V_BSTR(&varIn));
			pSrcBuffer = reinterpret_cast<char*>(V_BSTR(&varIn));
			eValueType = VALUETYPE_BSTR;
		}
		else
		{
			eValueType = VALUETYPE_NOTSUPPORTED;
		}

		// pSrcBuffer is a pointer to a buffer 
		// lSrcMaxLen set
		if (eValueType != VALUETYPE_NOTSUPPORTED)
		{
			// unzip
			CAutoBuffer AutoBuffer;
			AutoBuffer.SetGrowBy(lSrcMaxLen * 3); // assume 3 * compression

			CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

			if (pZipper->UnzipBufferToBuffer(&AutoBuffer, pSrcBuffer, lSrcMaxLen, NULL))
			{
				// return data
				char* pTgtBuffer = AutoBuffer.GetData();
				ULONG uLength = AutoBuffer.GetDataSize();
				*pbstrOut = ::SysAllocStringByteLen(pTgtBuffer, uLength);
				hr = S_OK;
			}
			else
			{
				Error(L"Unable to unzip data");
			}

			delete pZipper;
		}
		else
		{
			Error(L"Error accessing variant data");
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::XMLNODECompressToSafeArray(IUnknown* pUnknownXMLNode, VARIANT *pvarOut)
{
	HRESULT hr = E_FAIL;
	VariantInit(pvarOut);

	try
	{
		MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		VARIANT vstrValue;
		VariantInit(&vstrValue);
		if (SUCCEEDED(XMLDOMNodePtr->get_nodeTypedValue(&vstrValue)))
		{
			hr = SafeArrayCompressToSafeArray(vstrValue, pvarOut);
		}
		::VariantClear(&vstrValue);
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::SafeArrayDecompressToXMLNODE(VARIANT varIn, IUnknown* pUnknownXMLNode)
{
	HRESULT hr = E_FAIL;

	try
	{
		// extract source
		VARIANT var;
		// ... create variant
		VariantInit(&var);
		if (SUCCEEDED(SafeArrayDecompressToSafeArray(varIn, &var)))
		{
			// ... and store it in the xml node
			MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
			hr = XMLDOMNodePtr->put_nodeTypedValue(var);
			if (SUCCEEDED(hr))
			{
				hr = S_OK;
			}	
			else
			{
				Error(L"Error writing data to XML node");
			}
		}
		VariantClear(&var);
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::XMLNODEBINBASE64CompressToSafeArray(IUnknown* pUnknownXMLNode, VARIANT *pvarOut)
{
	HRESULT hr = E_FAIL;
	try
	{
		MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		hr = XMLDOMNodePtr->put_dataType(_bstr_t(L"bin.base64"));
		if (SUCCEEDED(hr))
		{
			hr = XMLNODECompressToSafeArray(pUnknownXMLNode, pvarOut);
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}
	return hr;
}

STDMETHODIMP CdmsCompression1::SafeArrayDecompressToXMLNODEBINBASE64(VARIANT varIn, IUnknown* pUnknownXMLNode)
{
	HRESULT hr = E_FAIL;

	//::DebugBreak();
	try
	{
		MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		hr = XMLDOMNodePtr->put_dataType(_bstr_t(L"bin.base64"));
		if (SUCCEEDED(hr))
		{
			hr = SafeArrayDecompressToXMLNODE(varIn, pUnknownXMLNode);
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}
	return hr;
}

STDMETHODIMP CdmsCompression1::XMLNODEBINHEXCompressToSafeArray(IUnknown* pUnknownXMLNode, VARIANT *pvarOut)
{
	HRESULT hr = E_FAIL;
	try
	{
		MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		hr = XMLDOMNodePtr->put_dataType(_bstr_t(L"bin.hex"));
		if (SUCCEEDED(hr))
		{
			hr = XMLNODECompressToSafeArray(pUnknownXMLNode, pvarOut);
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}
	return hr;
}

STDMETHODIMP CdmsCompression1::SafeArrayDecompressToXMLNODEBINHEX(VARIANT varIn, IUnknown* pUnknownXMLNode)
{
	HRESULT hr = E_FAIL;
	try
	{
		MSXML::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		hr = XMLDOMNodePtr->put_dataType(_bstr_t(L"bin.hex"));
		if (SUCCEEDED(hr))
		{
			hr = SafeArrayDecompressToXMLNODE(varIn, pUnknownXMLNode);
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}
	return hr;
}

STDMETHODIMP CdmsCompression1::SafeArrayCompressToSafeArray(VARIANT varIn, VARIANT *pvarOut)
{
	HRESULT hr = E_FAIL;
	VariantInit(pvarOut);

	enum
	{
		VALUETYPE_SAFEARRAYBYTES,
		VALUETYPE_BSTR,
		VALUETYPE_NOTSUPPORTED
	} eValueType = VALUETYPE_NOTSUPPORTED;

	try
	{
		// extract source
		long lSrcMaxLen = 0;
		char* pSrcBuffer = NULL;
		LONG lLBound = 0;
		LONG lUBound = 0;
		CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
		if (V_VT(&varIn) == (VT_ARRAY | VT_UI1) && // array of bytes
			SafeArrayGetDim(V_ARRAY(&varIn)) == 1 && // one dimensional 
			SUCCEEDED(SafeArrayGetLBound(V_ARRAY(&varIn), 1, &lLBound)) &&
			SUCCEEDED(SafeArrayGetUBound(V_ARRAY(&varIn), 1, &lUBound)) &&
			SUCCEEDED(SafeArrayAccessUnaccessData.Access(V_ARRAY(&varIn), reinterpret_cast<void**>(&pSrcBuffer))))
		{
			eValueType = VALUETYPE_SAFEARRAYBYTES;
			lSrcMaxLen = lUBound - lLBound + 1;
		}
		else if (V_VT(&varIn) == VT_BSTR)
		{
			lSrcMaxLen = ::SysStringByteLen(V_BSTR(&varIn));
			pSrcBuffer = reinterpret_cast<char*>(V_BSTR(&varIn));
			eValueType = VALUETYPE_BSTR;
		}
		else
		{
			eValueType = VALUETYPE_NOTSUPPORTED;
		}

		// pSrcBuffer is a pointer to a buffer 
		// lSrcMaxLen set
		if (eValueType != VALUETYPE_NOTSUPPORTED)
		{
			// zip
			CAutoBuffer AutoBuffer;
			AutoBuffer.SetGrowBy(lSrcMaxLen);

			CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

			if (pZipper->ZipBufferToBuffer(&AutoBuffer, pSrcBuffer, lSrcMaxLen, NULL))
			{
				// return data
				char* pTgtBuffer = AutoBuffer.GetData();
				ULONG uLength = AutoBuffer.GetDataSize();

				//Set the type to an array of unsigned chars (OLE SAFEARRAY)
				pvarOut->vt = VT_ARRAY | VT_UI1;
				SAFEARRAYBOUND  rgsabound[1];
				rgsabound[0].cElements = uLength;
				rgsabound[0].lLbound = 0;
				pvarOut->parray = SafeArrayCreate(VT_UI1,1,rgsabound);
				if(pvarOut->parray != NULL)
				{	
					void* pArrayData = NULL;
					CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
					SafeArrayAccessUnaccessData.Access(pvarOut->parray, &pArrayData);
					memcpy(pArrayData, pTgtBuffer, uLength);
					hr = S_OK;
				}
				else
				{
					Error(L"Unable to create safe array");
				}
			}
			else
			{
				Error(L"Unable to zip data");
			}

			delete pZipper;
		}
		else
		{
			Error(L"Unsupported variant type passed as input");
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::SafeArrayDecompressToSafeArray(VARIANT varIn, VARIANT* pvarOut)
{
	HRESULT hr = E_FAIL;

	enum
	{
		VALUETYPE_SAFEARRAYBYTES,
		VALUETYPE_BSTR,
		VALUETYPE_NOTSUPPORTED
	} eValueType = VALUETYPE_NOTSUPPORTED;

	try
	{
		// extract source
		long lSrcMaxLen = 0;
		char* pSrcBuffer = NULL;
		LONG lLBound = 0;
		LONG lUBound = 0;
		CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
		if (V_VT(&varIn) == (VT_ARRAY | VT_UI1) && // array of bytes
			SafeArrayGetDim(V_ARRAY(&varIn)) == 1 && // one dimensional 
			SUCCEEDED(SafeArrayGetLBound(V_ARRAY(&varIn), 1, &lLBound)) &&
			SUCCEEDED(SafeArrayGetUBound(V_ARRAY(&varIn), 1, &lUBound)) &&
			SUCCEEDED(SafeArrayAccessUnaccessData.Access(V_ARRAY(&varIn), reinterpret_cast<void**>(&pSrcBuffer))))
		{
			eValueType = VALUETYPE_SAFEARRAYBYTES;
			lSrcMaxLen = lUBound - lLBound + 1;
		}
		else if (V_VT(&varIn) == VT_BSTR)
		{
			lSrcMaxLen = ::SysStringByteLen(V_BSTR(&varIn));
			pSrcBuffer = reinterpret_cast<char*>(V_BSTR(&varIn));
			eValueType = VALUETYPE_BSTR;
		}
		else
		{
			eValueType = VALUETYPE_NOTSUPPORTED;
		}

		// pSrcBuffer is a pointer to a buffer 
		// lSrcMaxLen set
		if (eValueType != VALUETYPE_NOTSUPPORTED)
		{
			// unzip
			CAutoBuffer AutoBuffer;
			AutoBuffer.SetGrowBy(lSrcMaxLen * 3); // assume 3 * compression

			CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

			if (pZipper->UnzipBufferToBuffer(&AutoBuffer, pSrcBuffer, lSrcMaxLen, NULL))
			{
				// return data
				char* pTgtBuffer = AutoBuffer.GetData();
				ULONG uLength = AutoBuffer.GetDataSize();
				//Set the type to an array of unsigned chars (OLE SAFEARRAY)
				pvarOut->vt = VT_ARRAY | VT_UI1;
				SAFEARRAYBOUND  rgsabound[1];
				rgsabound[0].cElements = uLength;
				rgsabound[0].lLbound = 0;
				pvarOut->parray = SafeArrayCreate(VT_UI1,1,rgsabound);
				if(pvarOut->parray != NULL)
				{	
					void* pArrayData = NULL;
					CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
					SafeArrayAccessUnaccessData.Access(pvarOut->parray, &pArrayData);
					memcpy(pArrayData, pTgtBuffer, uLength);
					hr = S_OK;
				}
				else
				{
					Error(L"Unable to create safe array");
				}
			}
			else
			{
				Error(L"Unable to unzip data");
			}

			delete pZipper;
		}
		else
		{
			Error(L"Error accessing variant data");
		}
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::FileCompressToFile(BSTR bstrSrcFile, BSTR bstrTgtFile)
{
	HRESULT hr = E_FAIL;

	try
	{
		CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

		if (pZipper->ZipFileToFile(bstrTgtFile, bstrSrcFile))
		{
			hr = S_OK;
		}
		else
		{
			Error(L"Failure zipping file");
		}

		delete pZipper;
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}

STDMETHODIMP CdmsCompression1::FileDecompressToFile(BSTR bstrSrcFile, BSTR bstrTgtFile)
{
	HRESULT hr = E_FAIL;

	try
	{
		CZipper* pZipper = CZipper::CreateZipper(m_bstrCompressionMethod);

		if (pZipper->UnzipFileToFile(bstrTgtFile, bstrSrcFile))
		{
			hr = S_OK;
		}
		else
		{
			Error(L"Failure unzipping file");
		}

		delete pZipper;
	}
	catch(...)
	{
		Error(L"Exception caught");
	}

	return hr;
}


