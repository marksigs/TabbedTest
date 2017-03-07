// PkZip.cpp : Implementation of CPkZip
#include "stdafx.h"
#include <comdef.h>
#include "DmsCompression.h"
#include "PkZip.h"
#include "ZLibZipper.h"
#include "ZLIB\msg\src\vstudio\vc6\zlibwapi\unzip.h"
#include "SafeArrayAccessUnaccessData.h"


/////////////////////////////////////////////////////////////////////////////
// CPkZip

STDMETHODIMP CPkZip::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IPkZip
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



///////////////////////////////////////////////////////////////////////////////
//	ZipFile:
//		Zip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
STDMETHODIMP CPkZip::ZipFile(BSTR bstrZipFile, BSTR bstrFile, BOOL noPaths)
{
	HRESULT hr = E_FAIL;

	if (wcslen(bstrZipFile) == 0)
	{
		Error(L"Invalid parameter: bstrZipFile");
	}
	else
	{
		try
		{
			CZipper* pZipper = CZipper::CreateZipper(L"ZLIB");

			if (pZipper->ZipFile(bstrZipFile, bstrFile, noPaths))
			{
				hr = S_OK;
			}
			else
			{
				_bstr_t error(L"Failure zipping file: ");
				error += pZipper->GetLastErrorText();
				Error(static_cast<wchar_t*>(error));
			}

			delete pZipper;
		}
		catch(...)
		{
			Error(L"Exception caught");
		}
	}

	return hr;
}

///////////////////////////////////////////////////////////////////////////////
//	UnzipFile:
//		Unzip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
STDMETHODIMP CPkZip::UnzipFile(BSTR bstrZipFile, BSTR bstrFile, BSTR bstrUnzipDir)
{
	HRESULT hr = E_FAIL;

	if (wcslen(bstrZipFile) == 0)
	{
		Error(L"Invalid parameter: bstrZipFile");
	}
	else
	{
		try
		{
			CZipper* pZipper = CZipper::CreateZipper(L"ZLIB");
 
			if (pZipper->UnzipFile(bstrZipFile, bstrFile, bstrUnzipDir))
			{
				hr = S_OK;
			}
			else
			{
				_bstr_t error(L"Failure unzipping file: ");
				error += pZipper->GetLastErrorText();
				Error(static_cast<wchar_t*>(error));
			}

			delete pZipper;
		}
		catch(...)
		{
			Error(L"Exception caught");
		}
	}

	return hr;
}

STDMETHODIMP CPkZip::UnzipFileToSafeArray(BSTR bstrZipFile, BSTR bstrFile, VARIANT* pvarUnzippedFiles, LONG* pdwUnzippedFilesCount)
{
	HRESULT hr = E_FAIL;

	if (wcslen(bstrZipFile) == 0)
	{
		Error(L"Invalid parameter: bstrZipFile");
	}
	else if (pvarUnzippedFiles == NULL)
	{
		Error(L"Invalid parameter: pvarUnzippedFiles");
	}
	else if (pdwUnzippedFilesCount == NULL)
	{
		Error(L"Invalid parameter: pdwUnzippedFilesCount");
	}
	else
	{
		try
		{
			// The unzipped files buffer (pUnzippedFiles) has the following structure for each unzipped file
			// it contains:
			//
			// Unzipped file name (null terminated string).
			// 4 bytes (long): Number of bytes in unzipped file.
			// Unzipped file (binary).
			//
			// The unzipped files occur immediately one after the other.

			HANDLE hFile = INVALID_HANDLE_VALUE;
			if ((hFile = ::CreateFile(bstrZipFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, NULL, NULL)) != INVALID_HANDLE_VALUE)
			{
				// Assume a compression ratio of 10 to 1.
				long lUnzippedFilesMaxSize = ::GetFileSize(hFile, NULL) * 10;
				::CloseHandle(hFile);

				CZLibZipper* pZipper = reinterpret_cast<CZLibZipper*>(CZipper::CreateZipper(L"ZLIB"));
		 
				char* pUnzippedFiles = NULL;

				int result = UNZ_OK;
				do
				{
					*pdwUnzippedFilesCount = 0;

					if (pUnzippedFiles != NULL)
					{
						_ASSERTE(_CrtIsValidHeapPointer(pUnzippedFiles));
						free(pUnzippedFiles);
					}

					pUnzippedFiles = (char*)malloc(lUnzippedFilesMaxSize);
        
					result = 
						pZipper->UnzipFileToBuffer(
							bstrZipFile, 
							bstrFile, 
							pUnzippedFiles,			// Output buffer to hold unzipped files.
							&lUnzippedFilesMaxSize,	// Maximum size of output buffer.
							pdwUnzippedFilesCount);	// Number of unzipped files.

					_ASSERTE(_CrtIsValidHeapPointer(pUnzippedFiles));

				} while (result == UNZ_MOREMEM);	// UNZ_MOREMEM == buffer not big enough, so increase its size.

				if (result == UNZ_OK)
				{
					// Set the type to an array of unsigned chars (OLE SAFEARRAY).
					pvarUnzippedFiles->vt = VT_ARRAY | VT_UI1;
					SAFEARRAYBOUND rgsabound[1];
					rgsabound[0].cElements = lUnzippedFilesMaxSize;
					rgsabound[0].lLbound = 0;
					pvarUnzippedFiles->parray = SafeArrayCreate(VT_UI1, 1, rgsabound);
					if (pvarUnzippedFiles->parray != NULL)
					{	
						void* pArrayData = NULL;
						CSafeArrayAccessUnaccessData SafeArrayAccessUnaccessData;
						SafeArrayAccessUnaccessData.Access(pvarUnzippedFiles->parray, &pArrayData);
						memcpy(pArrayData, pUnzippedFiles, lUnzippedFilesMaxSize);
						hr = S_OK;
					}
					else
					{
						Error(L"Unable to create safe array: pvarUnzippedFiles");
						hr = E_FAIL;
					}

					hr = S_OK;
				}
				else
				{
					_bstr_t error(L"Failure unzipping file: ");
					error += pZipper->GetLastErrorText();
					Error(static_cast<wchar_t*>(error));
				}

				if (pUnzippedFiles != NULL)
				{
					free(pUnzippedFiles);
				}

				delete pZipper;
			}
			else
			{
				Error("Unable to open zip file");
			}
		}
		catch(...)
		{
			Error(L"Exception caught");
		}
	}

	return hr;
}

