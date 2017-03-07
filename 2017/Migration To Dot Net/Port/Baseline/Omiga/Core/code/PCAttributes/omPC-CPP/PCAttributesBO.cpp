///////////////////////////////////////////////////////////////////////////////
//	FILE:			PCAttributesBO.cpp
//	DESCRIPTION: 	
//
//	SYSTEM:	    	Omiga
//	COPYRIGHT:		(c) 2006, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		23/08/2006	CORE293 refactored.
//						UNICODE strings only. 
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#pragma message("Requires Microsoft Platform SDK.")
#pragma message("Platform SDK include and lib directories must be before ATL directories in Tools, Options, Directories.")
#include <Iphlpapi.h>
#include <IPIfCons.h>
#include <comdef.h>
#include "omPC.h"
#include "PCAttributesBO.h"

#import "msxml3.dll"


static void AddPrinterNode(MSXML2::IXMLDOMDocumentPtr xmlResponse, _bstr_t printerName, bool defaultIndicator);
static _bstr_t GetDefaultPrinterFromRegistry();
static _bstr_t GetDefaultPrinterFromApi();
static _bstr_t GetWin32APIErrorText(const DWORD dwError);

/////////////////////////////////////////////////////////////////////////////
// CPCAttributesBO


STDMETHODIMP CPCAttributesBO::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IPCAttributesBO
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

STDMETHODIMP CPCAttributesBO::NameOfPC(BSTR *pbstrPCName)
{
	const unsigned long len = MAX_COMPUTERNAME_LENGTH + 1;
	wchar_t pstr[len] = L"";

	if (::GetComputerName(pstr, const_cast<unsigned long*>(&len)))
	{
		_bstr_t computerName(pstr);

		*pbstrPCName = computerName.copy();
	}

	return S_OK;
}


// Return a list of the local printers on this PC
STDMETHODIMP CPCAttributesBO::FindLocalPrinterList(BSTR bstrXMLRequest, BSTR* pbstrPrinterList)
{
	HRESULT hResult = S_OK;
	_bstr_t response = L"";

	try
	{
		DWORD dwBytesNeeded = 0;
		DWORD dwReturned = 0;
		PRINTER_INFO_4* pinfo4 = NULL;
		const unsigned int dwFlags = PRINTER_ENUM_CONNECTIONS | PRINTER_ENUM_LOCAL;

		while (!EnumPrinters(dwFlags, NULL, 4, (LPBYTE)pinfo4, dwBytesNeeded, &dwBytesNeeded, &dwReturned))
		{
			if (pinfo4 != NULL) 
			{
				LocalFree(pinfo4);
			}
			DWORD error = GetLastError();
			if (error == ERROR_INSUFFICIENT_BUFFER)
			{
				pinfo4 = (PRINTER_INFO_4*)LocalAlloc(LPTR, dwBytesNeeded);
			}
			else
			{
				throw new _bstr_t(L"EnumPrinters error: " + GetWin32APIErrorText(error));
			}
		}

		MSXML2::IXMLDOMDocumentPtr xmlResponse(__uuidof(MSXML2::DOMDocument));
		MSXML2::IXMLDOMNodePtr xmlNode = xmlResponse->createElement(L"RESPONSE");
		xmlResponse->appendChild(xmlNode);

		// If there are no printers found then return an error. or there was another error
		if (dwReturned > 0)
		{
			// if no default printer then the string will be zero lenght
			_bstr_t defaultPrinterName = GetDefaultPrinterFromApi();
			// Will only add it if there was one found.

			AddPrinterNode(xmlResponse, defaultPrinterName, true);
    
			for (unsigned int i = 0; i < dwReturned; i++)
			{        
				// if this is not the default printer then add it to the xml.
				// the default printer should already be in the list by the time you get here
				if (wcsicmp(defaultPrinterName, pinfo4[i].pPrinterName) != 0)
				{
					AddPrinterNode(xmlResponse, pinfo4[i].pPrinterName, false);
				}
			}
    
		}
		else
		{
			if (pinfo4 != NULL)
			{
				LocalFree(pinfo4);
			}
			throw new _bstr_t(L"No printers found");
		}

		if (pinfo4 != NULL)
		{
			LocalFree(pinfo4);
		}

		xmlResponse->documentElement->setAttribute(L"TYPE", L"SUCCESS");

		response = xmlResponse->xml;
	}
	catch(_com_error& e)
	{
		response = L"<RESPONSE TYPE=\"APPERR\"><ERROR><ERRORNUMBER>500</ERRORNUMBER><ERRORSOURCE>PCAttributes.FindLocalPrinterList</ERRORSOURCE><ERRORDESCRIPTION>";
		response += e.Description();
		response += L"</ERRORDESCRIPTION></ERROR></RESPONSE>";
	}
	catch(_bstr_t* e)
	{
		response = L"<RESPONSE TYPE=\"APPERR\"><ERROR><ERRORNUMBER>500</ERRORNUMBER><ERRORSOURCE>PCAttributes.FindLocalPrinterList</ERRORSOURCE><ERRORDESCRIPTION>";
		response += *e;
		response += L"</ERRORDESCRIPTION></ERROR></RESPONSE>";
		delete e;
	}
	catch(...)
	{
		response = L"<RESPONSE TYPE=\"APPERR\"><ERROR><ERRORNUMBER>500</ERRORNUMBER><ERRORSOURCE>PCAttributes.FindLocalPrinterList</ERRORSOURCE><ERRORDESCRIPTION>Unknown error</ERRORDESCRIPTION></ERROR></RESPONSE>";
	}

	*pbstrPrinterList = response.copy();

	return hResult;
}

// If the newTag is > zero length then add it to the xml string.
// Return false unless you successfully add a printer name that is not a zero length string.
void AddPrinterNode(MSXML2::IXMLDOMDocumentPtr xmlResponse, _bstr_t printerName, bool defaultIndicator = false)
{
	if (printerName.length() > 0)
	{
		MSXML2::IXMLDOMNodePtr xmlPrinter = xmlResponse->createElement(L"PRINTER");
		MSXML2::IXMLDOMNodePtr xmlPrinterName = xmlResponse->createElement(L"PRINTERNAME");
		MSXML2::IXMLDOMNodePtr xmlDefaultIndicator = xmlResponse->createElement(L"DEFAULTINDICATOR");
		xmlResponse->documentElement->appendChild(xmlPrinter);
		xmlPrinter->appendChild(xmlPrinterName);
		xmlPrinter->appendChild(xmlDefaultIndicator);

		xmlPrinterName->text = printerName;
		xmlDefaultIndicator->text = defaultIndicator ? L"1" : L"0";
	}
	else
	{
		throw new _bstr_t(L"Printer name length is 0");
	}
}


_bstr_t GetDefaultPrinterFromRegistry()
{
	_bstr_t printerName = L"";

	HKEY hKey = NULL;
	DWORD error = 0;
	if ((error = RegOpenKeyEx(
			HKEY_CURRENT_USER,
			L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows",
			0, KEY_READ, &hKey)) == ERROR_SUCCESS)
	{
		const ULONG lRet = 256;
		UCHAR szValue[lRet];

		error = RegQueryValueEx(hKey, L"Device", NULL, NULL, szValue, const_cast<ULONG*>(&lRet));
		RegCloseKey(hKey);
		
		if (error == ERROR_SUCCESS)
		{
			// Parse the string to get the printer name out
			wchar_t* commaPos = NULL;
			if ((commaPos = wcschr(reinterpret_cast<wchar_t*>(szValue), L',')) != NULL)
			{
				*commaPos = L'\0';
				printerName = reinterpret_cast<wchar_t*>(szValue);
			}
		}
		else
		{
			throw new _bstr_t(L"Unable to read registry key HKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows" + GetWin32APIErrorText(error));
		}
	}
	else
	{
		throw new _bstr_t(L"Unable to open registry key HKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows\\Device: " + GetWin32APIErrorText(error));
	}

	return printerName;
}

_bstr_t GetDefaultPrinterFromApi()
{
	_bstr_t defaultPrinter = L"";

	DWORD chBuffer = 0;

	// Get required buffer size.
	::GetDefaultPrinter(NULL, &chBuffer);
	wchar_t* pszBuffer = new wchar_t[chBuffer];

	if (pszBuffer != NULL)
	{
		if (::GetDefaultPrinter(pszBuffer, &chBuffer))
		{
			defaultPrinter = pszBuffer;
		}
		else
		{
			delete []pszBuffer;
			throw new _bstr_t(L"Unable to get default printer: " + GetWin32APIErrorText(GetLastError()));
		}

		delete []pszBuffer;
	}

	return defaultPrinter;
}


_bstr_t GetWin32APIErrorText(const DWORD dwError) 
{
	const unsigned int maxLen = 256;
	wchar_t messageBuffer[maxLen] = L"";

	if (
		::FormatMessage( 
			FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ARGUMENT_ARRAY,
			NULL,
			dwError,
			LANG_NEUTRAL,
			messageBuffer,
			maxLen,
			NULL))
	{
		if (messageBuffer)
		{
			// remove cr and newline character
			messageBuffer[wcscspn(messageBuffer, L"\r\n")] = L'\0';
		}
	}

	return messageBuffer;
}

STDMETHODIMP CPCAttributesBO::GetMACAddress(BSTR *pbstrMACAddress)
{
	ULONG outBufLen = 0;
	DWORD result = 0;
	_bstr_t macAddress = L"";
	_bstr_t loopbackAddress = L"";
	
	if ((result = ::GetAdaptersInfo(NULL, &outBufLen)) == ERROR_BUFFER_OVERFLOW)
	{
		PIP_ADAPTER_INFO pipAdapterInfo = reinterpret_cast<IP_ADAPTER_INFO*>(new BYTE[outBufLen]);
		
		if ((result = ::GetAdaptersInfo(pipAdapterInfo, &outBufLen)) == ERROR_SUCCESS)
		{
			// The adapter types are in the order in which we want them to be checked; 
			// we will return the MAC address of the first one found.
			unsigned int adapterTypes[] = 
			{ 
				MIB_IF_TYPE_ETHERNET,
				MIB_IF_TYPE_TOKENRING,
				MIB_IF_TYPE_FDDI,
				MIB_IF_TYPE_PPP,
				MIB_IF_TYPE_SLIP,
				MIB_IF_TYPE_LOOPBACK,
				MIB_IF_TYPE_OTHER
			};

			for (int adapterType = 0; adapterType < (sizeof(adapterTypes) / sizeof(adapterTypes[0])) && macAddress.length() == 0; adapterType++)
			{
				PIP_ADAPTER_INFO pipThisAdapterInfo = pipAdapterInfo;
				while (pipThisAdapterInfo && macAddress.length() == 0)
				{
					if (pipThisAdapterInfo->Type == adapterTypes[adapterType] && pipThisAdapterInfo->AddressLength > 0)
					{
						for (unsigned int addressByte = 0; addressByte < pipThisAdapterInfo->AddressLength; addressByte++)
						{
							wchar_t byteStr[3] = L"";
							swprintf(byteStr, L"%.02X", pipThisAdapterInfo->Address[addressByte]);
							macAddress += byteStr;
						}
						strlwr(pipThisAdapterInfo->Description);
						if (strstr(pipThisAdapterInfo->Description, "loopback"))
						{
							// Do not use "MS LoopBack Adapter" unless there are no other adapters.
							loopbackAddress = macAddress;
							macAddress = L"";
						}
					}
					pipThisAdapterInfo = pipThisAdapterInfo->Next;
				}
			}

			if (macAddress.length() == 0 && loopbackAddress.length() > 0)
			{
				// "MS LoopBack Adapter" is the only one available, so use its MAC address.
				macAddress = loopbackAddress;
			}
		}

		delete []pipAdapterInfo;
	}

	if (macAddress.length() > 0)
	{
		*pbstrMACAddress = macAddress.copy();
	}

	return S_OK;
}

STDMETHODIMP CPCAttributesBO::Sleep(int dwMilliseconds)
{
	::Sleep(dwMilliseconds);
	return S_OK;
}
