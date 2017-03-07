# Microsoft Developer Studio Project File - Name="Files" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Generic Project" 0x010a

CFG=Files - Win32 Unicode Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Files.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Files.mak" CFG="Files - Win32 Unicode Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Files - Win32 Unicode Debug" (based on "Win32 (x86) Generic Project")
!MESSAGE "Files - Win32 Unicode Release MinDependency" (based on "Win32 (x86) Generic Project")
!MESSAGE "Files - Win32 Unicode Release MinDependencyW2K" (based on "Win32 (x86) Generic Project")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/dev/MessageQueue/MessageQueueListener/Files", ILRAAAAA"
# PROP Scc_LocalPath "."
MTL=midl.exe

!IF  "$(CFG)" == "Files - Win32 Unicode Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Unicode Debug"
# PROP BASE Intermediate_Dir "Unicode Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Unicode Debug"
# PROP Intermediate_Dir "Unicode Debug"
# PROP Target_Dir ""

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependency"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Unicode Release MinDependency"
# PROP BASE Intermediate_Dir "Unicode Release MinDependency"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Unicode Release MinDependency"
# PROP Intermediate_Dir "Unicode Release MinDependency"
# PROP Target_Dir ""

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Files___Win32_Unicode_Release_MinDependencyW2K"
# PROP BASE Intermediate_Dir "Files___Win32_Unicode_Release_MinDependencyW2K"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Files___Win32_Unicode_Release_MinDependencyW2K"
# PROP Intermediate_Dir "Files___Win32_Unicode_Release_MinDependencyW2K"
# PROP Target_Dir ""

!ENDIF 

# Begin Target

# Name "Files - Win32 Unicode Debug"
# Name "Files - Win32 Unicode Release MinDependency"
# Name "Files - Win32 Unicode Release MinDependencyW2K"
# Begin Group "Debug"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueListener\DebugU\MessageQueueListener.xml
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\DebugU\MessageQueueListenerLOG.log
# End Source File
# End Group
# Begin Group "ReleaseUMinDependency"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependency\MessageQueueListener.xml
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependency\MessageQueueListenerLOG.log
# End Source File
# End Group
# Begin Group "ReleaseUMinDependencyW2K"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependencyW2K\MessageQueueListener.xml
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.log
# End Source File
# End Group
# Begin Group "Queue Information"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test2_MessageQueueComponentVC2_MSMQ.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test2_MessageQueueComponentVC2Success_MSMQ.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2_MSDAORA.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2_MSMQ.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2_ORAOLEDB.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2_Registry.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2_SQLOLEDB.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2Success_MSDAORA.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2Success_MSMQ.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2Success_ORAOLEDB.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2Success_Registry.xml
# End Source File
# Begin Source File

SOURCE=..\..\..\messagequeuetest\QueueInformation\Test_MessageQueueComponentVC2Success_SQLOLEDB.xml
# End Source File
# End Group
# Begin Group "Sample Configuration Files"

# PROP Default_Filter ""
# Begin Source File

SOURCE="..\MessageQueueListener\SampleConfigurationFiles\MessageQueueListener.xml - MSMQ1_Test"
# End Source File
# Begin Source File

SOURCE="..\MessageQueueListener\SampleConfigurationFiles\MessageQueueListener.xml - MSMQ1_Test (DIRECT OS)"
# End Source File
# Begin Source File

SOURCE="..\MessageQueueListener\SampleConfigurationFiles\MessageQueueListener.xml - OMMQ1_MSDAORA_Test"
# End Source File
# Begin Source File

SOURCE="..\MessageQueueListener\SampleConfigurationFiles\MessageQueueListener.xml - OMMQ1_ORAOLEDB_Test"
# End Source File
# Begin Source File

SOURCE="..\MessageQueueListener\SampleConfigurationFiles\MessageQueueListener.xml - OMMQ1_SQLOLEDB_Test"
# End Source File
# End Group
# Begin Group "Binaries (NT4)"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependency\MessageQueueComponentVC.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependency\MessageQueueComponentVC.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependency\MessageQueueComponentVC.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependency\MessageQueueListener.exe
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependency\MessageQueueListener.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependency\MessageQueueListener.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependency\MessageQueueListenerLOG.exe
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependency\MessageQueueListenerLOG.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependency\MessageQueueListenerLOG.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependency\MessageQueueListenerMTS.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependency\MessageQueueListenerMTS.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependency\MessageQueueListenerMTS.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependency\MessageQueueListenerPRF.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependency\MessageQueueListenerPRF.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependency\MessageQueueListenerPRF.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependency\MessageQueueComponentVC.lib
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependency\MessageQueueListenerMTS.lib
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependency\MessageQueueListenerPRF.lib
# End Source File
# End Group
# Begin Group "Binaries (W2K)"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependencyW2K\MessageQueueComponentVC.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependencyW2K\MessageQueueComponentVC.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependencyW2K\MessageQueueComponentVC.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependencyW2K\MessageQueueListener.exe
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependencyW2K\MessageQueueListener.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ReleaseUMinDependencyW2K\MessageQueueListener.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.exe
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\ReleaseUMinDependencyW2K\MessageQueueListenerLOG.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependencyW2K\MessageQueueListenerMTS.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependencyW2K\MessageQueueListenerMTS.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependencyW2K\MessageQueueListenerMTS.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependencyW2K\MessageQueueListenerPRF.dll
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependencyW2K\MessageQueueListenerPRF.map
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependencyW2K\MessageQueueListenerPRF.pdb
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\ReleaseUMinDependencyW2K\MessageQueueComponentVC.lib
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\ReleaseUMinDependencyW2K\MessageQueueListenerMTS.lib
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\ReleaseUMinDependencyW2K\MessageQueueListenerPRF.lib
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MessageQueueComponentVC\MessageQueueComponentVC.rc
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\MessageQueueListener.rc

!IF  "$(CFG)" == "Files - Win32 Unicode Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependency"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\MessageQueueListenerLOG.rc

!IF  "$(CFG)" == "Files - Win32 Unicode Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependency"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\MessageQueueListenerMTS.rc

!IF  "$(CFG)" == "Files - Win32 Unicode Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependency"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\MessageQueueListenerPRF.rc

!IF  "$(CFG)" == "Files - Win32 Unicode Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependency"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "Files - Win32 Unicode Release MinDependencyW2K"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\Resource.h
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\Resource.h
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\Resource.h
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\Resource.h
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerPRF\resource.h
# End Source File
# End Group
# Begin Source File

SOURCE=..\MessageQueueComponentVC\MessageQueueComponentVC1.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\MessageQueueComponentVC2.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueComponentVC\MessageQueueComponentVC2Success.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\MessageQueueListener.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\MessageQueueListener1.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\MessageQueueListenerLOG.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerLOG\MessageQueueListenerLOG1.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListenerMTS\MessageQueueListenerMTS1.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ThreadPoolManagerMSMQ1.rgs
# End Source File
# Begin Source File

SOURCE=..\MessageQueueListener\ThreadPoolManagerOMMQ1.rgs
# End Source File
# End Target
# End Project
