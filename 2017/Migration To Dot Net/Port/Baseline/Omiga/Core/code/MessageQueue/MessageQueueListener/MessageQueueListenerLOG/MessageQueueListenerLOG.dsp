# Microsoft Developer Studio Project File - Name="MessageQueueListenerLOG" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=MessageQueueListenerLOG - Win32 Unicode Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "MessageQueueListenerLOG.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "MessageQueueListenerLOG.mak" CFG="MessageQueueListenerLOG - Win32 Unicode Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "MessageQueueListenerLOG - Win32 Unicode Debug" (based on "Win32 (x86) Application")
!MESSAGE "MessageQueueListenerLOG - Win32 Unicode Release MinDependency" (based on "Win32 (x86) Application")
!MESSAGE "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2K" (based on "Win32 (x86) Application")
!MESSAGE "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2003" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/dev/MessageQueue/MessageQueueListener/MessageQueueListenerLOG", FIRAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Debug"

# PROP BASE Use_MFC 1
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "DebugU"
# PROP BASE Intermediate_Dir "DebugU"
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "DebugU"
# PROP Intermediate_Dir "DebugU"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_UNICODE" /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 version.lib /nologo /entry:"wWinMainCRTStartup" /subsystem:windows /debug /machine:I386 /pdbtype:sept
# Begin Custom Build - Performing registration
OutDir=.\DebugU
TargetPath=.\DebugU\MessageQueueListenerLOG.exe
InputPath=.\DebugU\MessageQueueListenerLOG.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	"$(TargetPath)" /Console 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode EXE on Windows 95 
	:end 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependency"

# PROP BASE Use_MFC 1
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "ReleaseUMinDependency"
# PROP BASE Intermediate_Dir "ReleaseUMinDependency"
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "ReleaseUMinDependency"
# PROP Intermediate_Dir "ReleaseUMinDependency"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 version.lib /nologo /subsystem:windows /map /debug /machine:I386
# Begin Custom Build - Performing registration
OutDir=.\ReleaseUMinDependency
TargetPath=.\ReleaseUMinDependency\MessageQueueListenerLOG.exe
InputPath=.\ReleaseUMinDependency\MessageQueueListenerLOG.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	"$(TargetPath)" /Console 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode EXE on Windows 95 
	:end 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Use_MFC 1
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "MessageQueueListenerLOG___Win32_Unicode_Release_MinDependencyW2K"
# PROP BASE Intermediate_Dir "MessageQueueListenerLOG___Win32_Unicode_Release_MinDependencyW2K"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "ReleaseUMinDependencyW2K"
# PROP Intermediate_Dir "ReleaseUMinDependencyW2K"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /Yu"stdafx.h" /FD /c
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 version.lib /nologo /subsystem:windows /map /debug /machine:I386
# ADD LINK32 version.lib /nologo /subsystem:windows /map /debug /machine:I386
# Begin Custom Build - Performing registration
OutDir=.\ReleaseUMinDependencyW2K
TargetPath=.\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.exe
InputPath=.\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	"$(TargetPath)" /Console 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode EXE on Windows 95 
	:end 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2003"

# PROP BASE Use_MFC 1
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "MessageQueueListenerLOG___Win32_Unicode_Release_MinDependencyW2003"
# PROP BASE Intermediate_Dir "MessageQueueListenerLOG___Win32_Unicode_Release_MinDependencyW2003"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 1
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "ReleaseUMinDependencyW2003"
# PROP Intermediate_Dir "ReleaseUMinDependencyW2003"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /Yu"stdafx.h" /FD /c
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 version.lib /nologo /subsystem:windows /map /debug /machine:I386
# ADD LINK32 version.lib /nologo /subsystem:windows /map /debug /machine:I386
# Begin Custom Build - Performing registration
OutDir=.\ReleaseUMinDependencyW2003
TargetPath=.\ReleaseUMinDependencyW2003\MessageQueueListenerLOG.exe
InputPath=.\ReleaseUMinDependencyW2003\MessageQueueListenerLOG.exe
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	if "%OS%"=="" goto NOTNT 
	if not "%OS%"=="Windows_NT" goto NOTNT 
	"$(TargetPath)" /Console 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	echo Server registration done! 
	goto end 
	:NOTNT 
	echo Warning : Cannot register Unicode EXE on Windows 95 
	:end 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "MessageQueueListenerLOG - Win32 Unicode Debug"
# Name "MessageQueueListenerLOG - Win32 Unicode Release MinDependency"
# Name "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2K"
# Name "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2003"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\MessageQueueListenerLOG.cpp
# End Source File
# Begin Source File

SOURCE=.\MessageQueueListenerLOG.idl
# ADD MTL /tlb ".\MessageQueueListenerLOG.tlb" /h "MessageQueueListenerLOG.h" /iid "MessageQueueListenerLOG_i.c" /Oicf
# End Source File
# Begin Source File

SOURCE=.\MessageQueueListenerLOG.rc
# End Source File
# Begin Source File

SOURCE=.\MessageQueueListenerLOG1.cpp
# End Source File
# Begin Source File

SOURCE=.\mpheap.cpp
# End Source File
# Begin Source File

SOURCE=.\mutex.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD BASE CPP /Yc"stdafx.h"
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolManager.cpp
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolManagerFunctionQueue.cpp
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolMessage.cpp
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolMessageFunctionQueue.cpp
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolThread.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\MessageQueueListenerLOG1.h
# End Source File
# Begin Source File

SOURCE=.\mpheap.h
# End Source File
# Begin Source File

SOURCE=.\mutex.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolManager.h
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolManagerFunctionQueue.h
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolMessage.h
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolMessageFunctionQueue.h
# End Source File
# Begin Source File

SOURCE=.\ThreadPoolThread.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\MessageQueueListenerLOG.rgs
# End Source File
# Begin Source File

SOURCE=.\MessageQueueListenerLOG1.rgs
# End Source File
# End Group
# Begin Source File

SOURCE=.\messages.mc

!IF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Debug"

# Begin Custom Build
InputPath=.\messages.mc
InputName=messages

BuildCmds= \
	mc -c $(InputName)

"$(InputName).rc" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependency"

# Begin Custom Build
InputPath=.\messages.mc
InputName=messages

BuildCmds= \
	mc -c $(InputName)

"$(InputName).rc" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2K"

# Begin Custom Build
InputPath=.\messages.mc
InputName=messages

BuildCmds= \
	mc -c $(InputName)

"$(InputName).rc" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "MessageQueueListenerLOG - Win32 Unicode Release MinDependencyW2003"

# Begin Custom Build
InputPath=.\messages.mc
InputName=messages

BuildCmds= \
	mc -c $(InputName)

"$(InputName).rc" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ENDIF 

# End Source File
# End Target
# End Project
