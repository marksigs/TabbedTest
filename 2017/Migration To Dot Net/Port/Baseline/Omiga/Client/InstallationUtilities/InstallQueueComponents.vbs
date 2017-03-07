'Attribute VB_Name = "Module1"
' Filename:     InstallQueueComponents.vbs

' Description:
'       VB Script that creates an "OmiQueue" package and installs the appropriate components into it.
'	These components are the ones the Completions Block Adapter (CBA) needs in order to write to
'	an Omiga queue.
'       If the package already exists it is deleted and recreated.
'       A log of components installed is written to file InstallQueueComponents.log in the working directory.
'       Completion of the installation is indicated via a messagebox.
'       The script was based on InstDLL.vbs, a file provided as part of the Microsoft Transaction Server Samples.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'RF     12/03/02        Created
'RF	15/03/02	SYS4205 omCOG moved to a separate script
'RF	02/04/02	SYS4349 omCOG is once again installed by this script
'------------------------------------------------------------------------------------------


Dim strMsg

' Get the component path if passed in as an argument

Set objArgs = WScript.Arguments

If objArgs.Count > 0 Then
        strDllPath = objArgs(0)
End If


Set fso = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = fso.CreateTextFile("InstallQueueComponents.log", True)

strMsg = "InstallQueueComponents starting"

MsgBox strMsg, 0, "Omiga 4 MTS Installation"
txtLogFile.WriteLine (strMsg & " at " & Now())
txtLogFile.WriteLine ("")

While Len(strDllPath) = 0
	strDllPath = InputBox("Please enter the location of the Omiga 4 components", _
        	"Omiga 4 MTS Installation", _
                "C:\Program Files\Marlborough Stirling\Omiga 4\DLL\")
Wend

Call InstallPackageEx("OmiQueue", strDllPath)

strMsg = "InstallQueueComponents complete"
MsgBox strMsg, 0, "Omiga 4 MTS Installation"
txtLogFile.WriteLine ("")
txtLogFile.WriteLine (strMsg & " at " & Now())
txtLogFile.Close


Sub InstallPackageEx(vstrPackageName, vstrDllPath)

        strMsg = "Installing package " & vstrPackageName & " ... "
        txtLogFile.WriteLine (strMsg)

        ' First, we create the catalog object
        Dim catalog
        Set catalog = CreateObject("MTSAdmin.Catalog.1")
        
        ' Then we get the packages collection
        Dim packages
        Set packages = catalog.GetCollection("Packages")
        packages.Populate
        
        ' Remove all packages that go by the same name as the package we wish to install
        numPackages = packages.Count
        For i = numPackages - 1 To 0 Step -1
            If packages.Item(i).Value("Name") = vstrPackageName Then
                packages.Remove (i)
            End If
        Next
        
        ' Commit our deletions
        packages.SaveChanges
        
        ' Add a new package
        Dim newPackage
        Set newPackage = packages.Add
        newPackage.Value("Name") = vstrPackageName
        newPackage.Value("SecurityEnabled") = "N"

        ' Commit new package
        packages.SaveChanges
        
        ' Refresh packages
        packages.Populate
        
        ' Get components collection for new package
        Dim components
        Set components = packages.GetCollection("ComponentsInPackage", newPackage.Value("ID"))

        '-----------------------------------------------
        ' Install components
        '-----------------------------------------------

        Dim util
        Set util = components.GetUtilInterface
                
        Call InstallComponent(util, vstrDllPath & "omBase.dll")
        Call InstallComponent(util, vstrDllPath & "omCOG.dll")
        Call InstallComponent(util, vstrDllPath & "omMQ.dll")
        Call InstallComponent(util, vstrDllPath & "omToOMMQ.dll")
 
        ' Refresh components
        components.Populate
        
        ' Commit the changes to the components
        components.SaveChanges

        strMsg = "Package " & vstrPackageName & " installed"
        txtLogFile.WriteLine (strMsg)

End Sub
                        
Sub InstallComponent(vutil, vstrFileName)
        strMsg = "Installing component " & vstrFileName
        txtLogFile.WriteLine (strMsg)
        vutil.InstallComponent vstrFileName, "", ""
End Sub
