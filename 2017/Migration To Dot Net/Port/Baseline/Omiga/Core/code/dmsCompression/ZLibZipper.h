///////////////////////////////////////////////////////////////////////////////
//	FILE:			ZLibZipper.h
//	DESCRIPTION: 	ZLib wrapper compression class.
//	COPYRIGHT:		(c) 2004, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		12/05/04	First version
///////////////////////////////////////////////////////////////////////////////

#ifndef ZLIBZIPPER_H
#define ZLIBZIPPER_H

#include "Zipper.h"

class COMPAPI CZLibZipper : public CZipper
{
public:
	CZLibZipper();
	virtual ~CZLibZipper();

	virtual BOOL ZipFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName,
		BOOL noPaths);

	virtual BOOL UnzipFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName,
		const wchar_t* pszUnzipDir);

	virtual int UnzipFileToBuffer(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName,
		char* pUnzippedFiles,			// Output buffer to hold unzipped files.
		long* plUnzippedFilesMaxSize,	// Maximum size of output buffer.
		long* plUnzippedFilesCount);	// Number of unzipped files.

protected:
	BOOL Compress();
	BOOL Decompress();
};


#endif

