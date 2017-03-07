
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the ZLIBZIP_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// ZLIBZIP_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef ZLIBZIP_EXPORTS
#define ZLIBZIP_API __declspec(dllexport)
#else
#define ZLIBZIP_API __declspec(dllimport)
#endif

static const int MAXLEN_ERROR = 1024;
extern char* GetTlsErrorBuffer();


ZLIBZIP_API int minizipFile(
	const wchar_t* zipFileName,	// Zip file.
	const wchar_t* fileName,	// File to zip.
	BOOL overwrite,				// Overwrite existing zip file?
	BOOL append,				// Append to existing zip file?
	int compressLevel,			// -1 = default, 0 = store only, 1 = faster, 9 = better.
	const wchar_t* password,	// Password may be NULL.
	BOOL nopaths,				// Do not include file path name in zip file?
	wchar_t* error,				// Receives any error.
	int maxlenError);

ZLIBZIP_API int miniunzFile(
	const wchar_t* zipFileName,	// Zip file.
	const wchar_t* fileName,	// File to extract; may be NULL.
	BOOL extractWithout,		// Extract without pathname (junk paths)?
	BOOL extractWith,			// Extract with pathname?
	const wchar_t* password,	// Password may be NULL.
	const wchar_t* unzipDir,	// Location in which to unzip files.
	char* pUnzippedFiles,		// Output buffer to hold unzipped files.
	long* plUnzippedFilesMaxSize,// Maximum size of output buffer.
	long* plUnzippedFilesCount,	// Number of unzipped files.
	wchar_t* error,				// Receives any error.
	int maxlenError);
