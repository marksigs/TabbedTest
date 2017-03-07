' Filename: 	InstallAlpha+.vbs

' Description: 
'	VB Script that creates the Alpha+ COM+ package and installs the appropriate components into them.
'	If the package already exists it is deleted and recreated.
'	A log of packages installed is written to InstallAlpha+.log in the working directory.
'	Completion of the installation is indicated via a messagebox.
'	The script was based on InstDLL.vbs, a file provided as part of the Microsoft Transaction Server Samples.
'------------------------------------------------------------------------------------------
'History:
'
'Prog  	Date     Description
'GHun  	25/03/04 BMIDS755 Created based on InstallOmiga4.vbs
'------------------------------------------------------------------------------------------

dim strMsg
dim strMsgBoxTitle
dim strIdentity
dim strPassword
dim objArgs
dim strDLLPath
dim fso
dim txtLogFile

strMsgBoxTitle = "Alpha+ COM+ Installation"

' Get the component path if passed in as an argument 

Set objArgs = WScript.Arguments

if objArgs.Count > 0 then
	strDLLPath = objArgs(0)
end if 

Set fso = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = fso.CreateTextFile("InstallAlpha+.log", True)

strMsg = "InstallAlpha+ starting"
'MsgBox strMsg,0,strMsgBoxTitle

txtLogFile.WriteLine(strMsg & " at " & Now())
txtLogFile.WriteLine("")

while len(strDllPath) = 0 
	strDllPath = InputBox("Please enter the location of the Alpha+ Calculations Engine components", _
               	              strMsgBoxTitle, _
                       	      "C:\Program Files\Marlborough Stirling\Alpha+ Calculations Engine\CalcEng\")
wend

call InstallPackageEx("Alpha+", strDllPath)

strMsg = "InstallAlpha+ complete"
MsgBox strMsg,vbOKOnly + vbInformation,strMsgBoxTitle
txtLogFile.WriteLine("")
txtLogFile.WriteLine(strMsg & " at " & Now())
txtLogFile.Close

sub InstallPackageEx(vstrAppName, vstrDllPath)

	Dim catalog
	Dim apps
	Dim newApp
	Dim numApps
	Dim i

	strMsg = "Installing package " & vstrAppName & " ... "
	txtLogFile.WriteLine(strMsg)

	' First, we create the catalog object

	Set catalog = CreateObject("COMAdmin.COMAdminCatalog.1")
	
	' Then we get the applications collection
	Set apps = catalog.GetCollection("Applications")
	apps.Populate
	
	' Remove all packages that go by the same name as the package we wish to install
	numApps = apps.Count
	For i = numApps - 1 To 0 Step -1
	    If apps.Item(i).Value("Name") = vstrAppName Then
	        apps.Remove(i)
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

	' Commit new package
	apps.SaveChanges
	
	' Refresh packages
	apps.Populate
	
	'-----------------------------------------------	
	' Install components 
	'-----------------------------------------------	
	call InstallComponent(catalog, vstrAppName, vstrDllPath & "AlphaMortgage.dll")
	call InstallComponent(catalog, vstrAppName, vstrDllPath & "AlphaCOMPlus.dll")

	strMsg = "Application " & vstrAppName & " installed"
	txtLogFile.WriteLine(strMsg)
end sub
			
sub InstallComponent(vcatalog, vstrAppName, vstrFileName)
	strMsg = "Installing component " & vstrFileName
	txtLogFile.WriteLine(strMsg)
	vcatalog.InstallComponent vstrappName, vstrFileName, "", ""
end sub
