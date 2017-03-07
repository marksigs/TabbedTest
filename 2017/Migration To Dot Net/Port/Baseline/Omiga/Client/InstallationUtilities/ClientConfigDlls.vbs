'------------------------------------------------------------------------------------------
' Filename:     ClientConfigDlls.vbs

' Description:
'   VB Script that creates PrintManager package and installs the appropriate components into them.
'   If the package already exists it is deleted and recreated.
'   A log of packages installed is written to InstallPrintManager.log in the working directory.
'   Completion of the installation is indicated via a messagebox.
'   The script was based on InstDLL.vbs, a file provided as part of the Microsoft Transaction Server Samples.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     07/02/01 Created based on InstallOmiga4.vbs
'GHun   30/09/04 Merged in CopyClientConfigDlls, unregister and delete old DLLs, include BMCompletions.dll
'		 and default to library application
'------------------------------------------------------------------------------------------
Option Explicit

Dim strMsg
Dim strMsgBoxTitle
Dim objArgs
Dim strDLLPath
Dim objFSO
Dim txtLogFile
Dim FileNames(11)
Dim i
Dim strFile
Dim strSystem32
Dim objShell

strMsgBoxTitle = "ClientConfigDll Manager Installation"

' Get the component path if passed in as an argument

Set objArgs = WScript.Arguments

If objArgs.Count > 0 Then
    strDLLPath = objArgs(0)
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = objFSO.CreateTextFile("ClientConfigDllManager.log", True)

strMsg = "InstallClientConfig BMidsInt starting"

MsgBox strMsg, vbInformation, strMsgBoxTitle
txtLogFile.WriteLine (strMsg & " at " & Now() & vbCrLf)

FileNames(0) = "BMPrinting.dll"
FileNames(1) = "omAdminRules.dll"
FileNames(2) = "omAPRules.dll"
FileNames(3) = "omCDRules.dll"
FileNames(4) = "omCMRules.dll"
FileNames(5) = "omCompRules.dll"
FileNames(6) = "omDataIngestionRules.dll"
FileNames(7) = "omPDM.dll"
FileNames(8) = "omRARules.dll"
FileNames(9) = "omTMRules.dll"
FileNames(10) = "BMCompletions.dll"

Do While Len(strDLLPath) = 0
    strDLLPath = InputBox("Please enter the location of the Client Config components", _
                              strMsgBoxTitle, _
                              "C:\Program Files\Marlborough Stirling\Omiga 4\DLL\")
Loop

strSystem32 = objFSO.GetSpecialFolder(1)    'SystemFolder
Set objShell = CreateObject("WScript.Shell")

For i = 0 To UBound(FileNames) - 1
    strFile = strDLLPath & FileNames(i)
    'If the file already exists then unregister and delete it
    If objFSO.FileExists(strFile) Then
        objShell.Run strSystem32 & "\regsvr32.exe /u /s " & strFile, , True
        objFSO.DeleteFile (strFile)
    End If
    objFSO.CopyFile "\\msgchhost5\Omiga VSS\BMIDSCodeVSS\OUTPUT\BMIDSSystemTest\Build\" + FileNames(i), strDLLPath
Next

InstallApp "ClientConfigDllManager", strDLLPath

strMsg = "Install ClientConfigDllManager Complete"
MsgBox strMsg, vbInformation, strMsgBoxTitle
txtLogFile.WriteLine (vbCrLf & strMsg & " at " & Now())
txtLogFile.Close

Set objFSO = Nothing
Set objShell = Nothing

Sub InstallApp(vstrAppName, vstrDllPath)

    Dim catalog
    Dim apps
    Dim newApp
    Dim numApps
    Dim i

    strMsg = "Installing package " & vstrAppName & " ... "
    txtLogFile.WriteLine (strMsg)

    ' First, we create the catalog object
    Set catalog = CreateObject("COMAdmin.COMAdminCatalog.1")
    
    ' Then we get the applications collection
    Set apps = catalog.GetCollection("Applications")
    apps.Populate
    
    ' Remove all packages that go by the same name as the package we wish to install
    numApps = apps.Count
    For i = numApps - 1 To 0 Step -1
        If apps.Item(i).Value("Name") = vstrAppName Then
            apps.Remove (i)
        End If
    Next
    
    ' Commit our deletions
    apps.SaveChanges
    
    ' Add a new package
    Set newApp = apps.Add
    newApp.Value("Name") = vstrAppName
    'newApp.Value("SecurityEnabled") = "N"
    'Set the activation to be library application
    newApp.Value("Activation") = 0

    ' Commit new application
    apps.SaveChanges
    
    ' Refresh packages
    apps.Populate
    
    '-----------------------------------------------
    ' Install components
    '-----------------------------------------------
    InstallComponent catalog, vstrAppName, vstrDllPath & "BMPrinting.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omAdminRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omAPRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omCDRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omCMRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omCompRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omDataIngestionRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omPDM.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omRARules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "omTMRules.dll"
    InstallComponent catalog, vstrAppName, vstrDllPath & "BMCompletions.dll"
    
    strMsg = "Application " & vstrAppName & " installed"
    txtLogFile.WriteLine (strMsg)

End Sub
            
Sub InstallComponent(vcatalog, vstrAppName, vstrFileName)
    Dim strMsg
    
    strMsg = "Installing component " & vstrFileName
    txtLogFile.WriteLine (strMsg)
    vcatalog.InstallComponent vstrAppName, vstrFileName, "", ""
End Sub
