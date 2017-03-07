Attribute VB_Name = "FVS_constants"
'header ----------------------------------------------------------------------------------
'Workfile:      FVS_Constants.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Helper functions for error handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'LD     26/01/01    Initial creation
'IK     20/04/2006  CORE261 omGemini integration
'------------------------------------------------------------------------------------------
Option Explicit
' constants match those defined in the stored procedure / package
Public Const FVS_ROOTFOLDERGUID = "A4D240E9F68B11d482BA005004E8D1A7"
Public Const STORAGE_NATIVE = 0
Public Const STORAGE_MSGCOMPRESS = 1
    
Public Const FVS_CHUNKSIZE As Long = 2000
'BG 01/10/02 SYS5619 Compression code
Public Const STORAGE_DMSCOMPRESSION1 = 1
Public Const STORAGE_DMSCOMPRESSION1BINBASE64 = 2
Public Const STORAGE_DMSCOMPRESSION1BINHEX = 3
'BG 01/10/02 SYS5619 Compression code - END

Public Const DELIVERYTYPE_DOC = 10
Public Const DELIVERYTYPE_PDF = 20
Public Const DELIVERYTYPE_RTF = 30
Public Const DELIVERYTYPE_XML = 40

'IK_20/04/2006_CORE261
Public Const DOCUMENTMANAGEMENTSYSTEMTYPE_DMS = 1
Public Const DOCUMENTMANAGEMENTSYSTEMTYPE_FILENET = 2
Public Const DOCUMENTMANAGEMENTSYSTEMTYPE_GEMINI = 3

