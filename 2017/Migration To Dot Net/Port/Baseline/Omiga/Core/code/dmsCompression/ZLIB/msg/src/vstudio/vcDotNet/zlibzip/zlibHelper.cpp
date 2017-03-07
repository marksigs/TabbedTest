#include "StdAfx.h"
#include "zlibHelper.h"
#include "zlibzip.h"


// Replacement for printf, suitable for a Win32 DLL.
// The resulting string is copied to a string in Thread Local Storage, i.e., it is thread safe.
int printfZlibHelper(const char * format, ...)
{
	int result = -1;

	char* error = GetTlsErrorBuffer();
	if (error != NULL)
	{
		va_list arglist;

		va_start(arglist, format);

		result = vsprintf(error, format, arglist);

		va_end(arglist);
	}

	return result;
}
