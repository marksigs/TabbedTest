/* <VERSION  CORELABEL="" LABEL="020.02.06.19.00" DATE="19/06/2002 18:51:20" VERSION="384" PATH="$/CodeCore/Code/FileVersioningSystem/omStream/omStream1.cpp"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			omStream1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      13/02/01    Initial version
//	LD		05/07/01	Base64ToObject modified to dataType property of XML node
//						passed in, and now accepts a variant value of type BSTR as
//						well as a safe array of bytes
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OmStream.h"
#include "omStream1.h"

#import "msxml3.dll" rename_namespace("xmlnamespace")

/////////////////////////////////////////////////////////////////////////////
// ComStream1

STDMETHODIMP ComStream1::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IomStream1
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP ComStream1::ObjectToBase64(IUnknown *pUnknownObject, IUnknown* pUnknownXMLNode)
{
	HGLOBAL hGlobal = NULL;
	HRESULT hr = E_FAIL;

	try
	{
		// attempt to retrieve the persistance interface from the object
		IPersistStoragePtr PersistStoragePtr(pUnknownObject);
		if (PersistStoragePtr)
		{
			// create IStorage based on an array of bytes in memory
			const size_t nInitialSize = 0;
			hGlobal = GlobalAlloc(GMEM_MOVEABLE, nInitialSize);
			if (hGlobal)
			{
				ILockBytesPtr LockBytesPtr;
				const BOOL bDeleteOnRelease = FALSE;
				hr = ::CreateILockBytesOnHGlobal(hGlobal, bDeleteOnRelease, &LockBytesPtr);
				if (SUCCEEDED(hr))
				{
					IStoragePtr StoragePtr;
					hr = ::StgCreateDocfileOnILockBytes(LockBytesPtr, 
						STGM_CREATE | STGM_DIRECT | STGM_WRITE | STGM_SHARE_EXCLUSIVE,
						0, &StoragePtr);
					if (SUCCEEDED(hr))
					{							
						// save the persistant object to the storage
						hr = PersistStoragePtr->Save(StoragePtr, FALSE);
						if (SUCCEEDED(hr))
						{
							hr = StoragePtr->Commit(STGC_ONLYIFCURRENT);
							if (SUCCEEDED(hr))
							{
								// save the clsid in the storage
								CLSID clsid;
								PersistStoragePtr->GetClassID(&clsid);
								hr = ::WriteClassStg(StoragePtr, clsid);
							}
						}

						if (SUCCEEDED(hr))
						{
							// access the stored object in the storage, and copy it to the xml node
							LPVOID lpvGlobalData = ::GlobalLock(hGlobal);
							if (lpvGlobalData)
							{
								DWORD dwSize = ::GlobalSize(hGlobal);

								VARIANT vstrValue;
								VariantInit(&vstrValue);
								V_VT(&vstrValue) = VT_ARRAY | VT_UI1;
								SAFEARRAYBOUND  rgsabound[1];
								rgsabound[0].cElements = dwSize;
								rgsabound[0].lLbound = 0;
								V_ARRAY(&vstrValue) = SafeArrayCreate(VT_UI1, 1, rgsabound);
								if(V_ARRAY(&vstrValue) != NULL)
								{	
									LPVOID lpvArrayData = NULL;
									hr = SafeArrayAccessData(V_ARRAY(&vstrValue), reinterpret_cast<void**>(&lpvArrayData));
									if (SUCCEEDED(hr))
									{
										memcpy(lpvArrayData, lpvGlobalData, dwSize);
										SafeArrayUnaccessData(V_ARRAY(&vstrValue));

										try
										{
											xmlnamespace::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
											hr = XMLDOMNodePtr->put_dataType(_bstr_t(L"bin.base64"));
											if (SUCCEEDED(hr))
											{
												hr = XMLDOMNodePtr->put_nodeTypedValue(vstrValue);
												if (SUCCEEDED(hr))
												{
													hr = S_OK;
												}	
												else
												{
													Error(L"Error writing data to XML node");
												}
											}
											else
											{
												Error(L"Error writing data type to XML node");
											}
										}
										catch (...)
										{
											Error(L"Error writing to XML node");
										}
									}
									else
									{
										Error(L"Error accessing safe array");
									}
								}
								else
								{
									Error(L"Error creating a safe array of bytes");
								}
								::VariantClear(&vstrValue);
								::GlobalUnlock(hGlobal);
							}
							else
							{
								Error(L"Unable to access memory");
							}
						}
						else
						{
							Error(L"Unable to write to storage");
						}
					}
					else
					{
						Error(L"Unable to create doc file on lock bytes");
					}
				}
				else
				{
					Error(L"Unable to create lock bytes");
				}
			}
			else
			{
				Error(L"Unable to allocate memory");
			}
		}
		else
		{
			Error(L"Unable to retrieve IPersistStorage");
		}
	}
	catch (...)
	{
		Error(L"Exception caught");
	}

	::GlobalFree(hGlobal);

	return hr;
}

STDMETHODIMP ComStream1::Base64ToObject(IUnknown* pUnknownXMLNode, IUnknown **ppUnknownObject)
{
	*ppUnknownObject = NULL;
	HRESULT hr = E_FAIL;
	
	enum
	{
		VALUETYPE_SAFEARRAYBYTES,
		VALUETYPE_BSTR,
		VALUETYPE_NOTSUPPORTED
	} eValueType = VALUETYPE_NOTSUPPORTED;
	
	try
	{
		xmlnamespace::IXMLDOMNodePtr XMLDOMNodePtr = pUnknownXMLNode;
		if (XMLDOMNodePtr->GetdataType() != _variant_t(L"bin.base64"))
		{
			Error(L"Node passed in is not of type bin.base64 (dataType property of the XML node is not bin.base64)");
		}
		else
		{
			VARIANT vstrValue;
			VariantInit(&vstrValue);
			if (SUCCEEDED(XMLDOMNodePtr->get_nodeTypedValue(&vstrValue)))
			{
				LONG lLBound = 0;
				LONG lUBound = 0;
				LPVOID lpvArrayData = NULL;
				
				if (V_VT(&vstrValue) == (VT_ARRAY | VT_UI1) && // array of bytes
					SafeArrayGetDim(V_ARRAY(&vstrValue)) == 1 && // one dimensional 
					SUCCEEDED(SafeArrayGetLBound(V_ARRAY(&vstrValue), 1, &lLBound)) &&
					SUCCEEDED(SafeArrayGetUBound(V_ARRAY(&vstrValue), 1, &lUBound)) &&
					SUCCEEDED(SafeArrayAccessData(V_ARRAY(&vstrValue), reinterpret_cast<void**>(&lpvArrayData))))
				{
					eValueType = VALUETYPE_SAFEARRAYBYTES;
				}
				else if ((V_VT(&vstrValue) == VT_BSTR))
				{
					lLBound = 0;
					lUBound = ::SysStringLen(V_BSTR(&vstrValue)) * sizeof(WCHAR);
					lpvArrayData = V_BSTR(&vstrValue);
					eValueType = VALUETYPE_BSTR;
				}
				else
				{
					eValueType = VALUETYPE_NOTSUPPORTED;
				}
				
				// lpvArrayData is a pointer to a buffer 
				// lLBound set
				// lUBound set
				if (eValueType != VALUETYPE_NOTSUPPORTED)
				{
					size_t nInitialSize = lUBound - lLBound;
					HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, nInitialSize);
					if (hGlobal)
					{
						LPVOID lpvGlobalData = ::GlobalLock(hGlobal);
						if (lpvGlobalData)
						{
							memcpy(lpvGlobalData, lpvArrayData, nInitialSize);
							ILockBytesPtr LockBytesPtr;
							hr = CreateILockBytesOnHGlobal(hGlobal, FALSE, &LockBytesPtr);
							if (SUCCEEDED(hr))
							{
								IStoragePtr StoragePtr;
								hr = StgOpenStorageOnILockBytes(LockBytesPtr, NULL,
									STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, NULL,
									0, &StoragePtr);
								if (SUCCEEDED(hr))
								{
									CLSID clsid;
									::ReadClassStg(StoragePtr, &clsid);
									IPersistStoragePtr PersistStoragePtr(clsid);
									if (PersistStoragePtr)
									{
										HRESULT hr = PersistStoragePtr->Load(StoragePtr);
										if (SUCCEEDED(hr))
										{
											IUnknownPtr UnknownPtr = PersistStoragePtr;
											*ppUnknownObject = UnknownPtr;
											UnknownPtr.Detach();
											hr = S_OK;
										}
										else
										{
											Error(L"Error loading from storage");
										}
									}
									else
									{
										Error(L"Unable to retrieve IPersistStorage");
									}
								}
								else
								{
									Error(L"Unable to open storage on lock bytes");
								}
							}
							else
							{
								Error(L"Unable to create lock bytes");
							}
						}
						else
						{
							Error(L"Unable to access memory");
						}
						::GlobalUnlock(hGlobal);
						if (eValueType == VALUETYPE_SAFEARRAYBYTES)
						{
							::SafeArrayUnaccessData(V_ARRAY(&vstrValue));
						}
					}
					else
					{
						Error(L"Unable to allocate memory");
					}
				}
				else
				{
					Error(L"Unsupported variant type passed as input");
				}
			}
			else
			{
				Error(L"Error extracting data from XML node");
			}
			::VariantClear(&vstrValue);
		}

 	}
	catch (...)
	{
		Error(L"Exception caught");
	}

	return hr;
}
