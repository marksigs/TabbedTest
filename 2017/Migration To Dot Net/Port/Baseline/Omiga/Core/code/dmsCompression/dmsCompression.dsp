# Microsoft Developer Studio Project File - Name="dmsCompression" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=dmsCompression - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "dmsCompression.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "dmsCompression.mak" CFG="dmsCompression - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "dmsCompression - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "dmsCompression - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/Code/Black Box/dmsCompression", CVDAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "dmsCompression - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "dmsCompression___Win32_Debug"
# PROP BASE Intermediate_Dir "dmsCompression___Win32_Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "vc6\DebugU"
# PROP Intermediate_Dir "vc6\DebugU"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "_DEBUG" /D "_UNICODE" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /FR /YX"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "_DEBUG" /D "_UNICODE" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /D "VC6" /FR /YX"stdafx.h" /FD /GZ /c
# ADD BASE RSC /l 0x809 /d "_DEBUG" /d "VC6"
# ADD RSC /l 0x809 /d "_DEBUG" /d "VC6"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /base:"0x6D60000" /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Performing registration
OutDir=.\vc6\DebugU
TargetPath=.\vc6\DebugU\dmsCompression.dll
InputPath=.\vc6\DebugU\dmsCompression.dll
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode DLL on Windows 95 
	:end 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "dmsCompression - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "dmsCompression___Win32_Release"
# PROP BASE Intermediate_Dir "dmsCompression___Win32_Release"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "vc6\ReleaseUMinDependency"
# PROP Intermediate_Dir "vc6\ReleaseUMinDependency"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /Zi /O1 /D "NDEBUG" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /YX"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /Zi /O1 /D "NDEBUG" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /D "VC6" /FR /YX"stdafx.h" /FD /c
# ADD BASE RSC /l 0x809 /d "NDEBUG" /d "VC6"
# ADD RSC /l 0x809 /d "NDEBUG" /d "VC6"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /base:"0x6D60000" /subsystem:windows /dll /debug /machine:I386
# Begin Custom Build - Performing registration
OutDir=.\vc6\ReleaseUMinDependency
TargetPath=.\vc6\ReleaseUMinDependency\dmsCompression.dll
InputPath=.\vc6\ReleaseUMinDependency\dmsCompression.dll
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode DLL on Windows 95 
	:end 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "dmsCompression - Win32 Debug"
# Name "dmsCompression - Win32 Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AutoBuff.cpp
# End Source File
# Begin Source File

SOURCE=.\AutoZipBuff.cpp
# End Source File
# Begin Source File

SOURCE=.\Crc.cpp
# End Source File
# Begin Source File

SOURCE=.\dmsCompression.cpp
# End Source File
# Begin Source File

SOURCE=.\dmsCompression.def
# End Source File
# Begin Source File

SOURCE=.\dmsCompression.idl
# ADD BASE MTL /tlb ".\dmsCompression.tlb" /h "dmsCompression.h" /iid "dmsCompression_i.c" /Oicf
# ADD MTL /tlb ".\dmsCompression.tlb" /h "dmsCompression.h" /iid "dmsCompression_i.c" /Oicf
# End Source File
# Begin Source File

SOURCE=.\dmsCompression.rc
# End Source File
# Begin Source File

SOURCE=.\dmsCompression1.cpp
# End Source File
# Begin Source File

SOURCE=.\PkZip.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD BASE CPP /Yc"stdafx.h"
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AutoBuff.h
# End Source File
# Begin Source File

SOURCE=.\AutoZipBuff.h
# End Source File
# Begin Source File

SOURCE=.\COMP.H
# End Source File
# Begin Source File

SOURCE=.\Crc.h
# End Source File
# Begin Source File

SOURCE=.\dmsCompression1.h
# End Source File
# Begin Source File

SOURCE=.\PkZip.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\SafeArrayAccessUnaccessData.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\stdhdr.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\dmsCompression1.rgs
# End Source File
# Begin Source File

SOURCE=.\PkZip.rgs
# End Source File
# End Group
# Begin Group "Zipper"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\Compapi.cpp
# End Source File
# Begin Source File

SOURCE=.\CompApi.h
# End Source File
# Begin Source File

SOURCE=.\CompAPIZipper.cpp
# End Source File
# Begin Source File

SOURCE=.\CompAPIZipper.h
# End Source File
# Begin Source File

SOURCE=.\Compress.fns
# End Source File
# Begin Source File

SOURCE=.\ZipData.h
# End Source File
# Begin Source File

SOURCE=.\ZipDef.h
# End Source File
# Begin Source File

SOURCE=.\Zipper.cpp
# End Source File
# Begin Source File

SOURCE=.\Zipper.h
# End Source File
# Begin Source File

SOURCE=.\ZLIB\msg\src\vstudio\vc7\zlib.h
# End Source File
# Begin Source File

SOURCE=.\ZLibZipper.cpp
# End Source File
# Begin Source File

SOURCE=.\ZLibZipper.h
# End Source File
# End Group
# Begin Group "Library Files"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\ZLIB\msg\src\vstudio\vc6\zlibwapi\Debug\zlibwapi.lib

!IF  "$(CFG)" == "dmsCompression - Win32 Debug"

!ELSEIF  "$(CFG)" == "dmsCompression - Win32 Release"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ZLIB\msg\src\vstudio\vc6\zlibwapi\Release\zlibwapi.lib

!IF  "$(CFG)" == "dmsCompression - Win32 Debug"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "dmsCompression - Win32 Release"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ZLIB\msg\src\vstudio\vc6\zlibzip\Debug\zlibzip.lib

!IF  "$(CFG)" == "dmsCompression - Win32 Debug"

!ELSEIF  "$(CFG)" == "dmsCompression - Win32 Release"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ZLIB\msg\src\vstudio\vc6\zlibzip\Release\zlibzip.lib

!IF  "$(CFG)" == "dmsCompression - Win32 Debug"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "dmsCompression - Win32 Release"

!ENDIF 

# End Source File
# End Group
# End Target
# End Project
