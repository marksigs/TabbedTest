' Filename: 	InstallPrintManager.vbs

' Description: 
'	VB Script that creates PrintManager package and installs the appropriate components into them.
'	If the package already exists it is deleted and recreated.
'	A log of packages installed is written to InstallPrintManager.log in the working directory.
'	Completion of the installation is indicated via a messagebox.
'	The script was based on InstDLL.vbs, a file provided as part of the Microsoft Transaction Server Samples.
'------------------------------------------------------------------------------------------
'History:
'
'Prog  	Date     Description
'RF    	07/02/01 Created based on InstallOmiga4.vbs
'AD 	18/03/03 change it back to be a server package - BMIDS00399
'------------------------------------------------------------------------------------------

dim strMsg
dim strMsgBoxTitle
dim strIdentity
dim strPassword


strMsgBoxTitle = "Print Manager Installation"

' Get the component path if passed in as an argument 

Set objArgs = WScript.Arguments

if objArgs.Count > 0 then
	strDLLPath = objArgs(0)
end if 

Set fso = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = fso.CreateTextFile("InstallPrintManager.log", True)

strMsg = "InstallPrintManager starting"

MsgBox strMsg,0,strMsgBoxTitle
txtLogFile.WriteLine(strMsg & " at " & Now())
txtLogFile.WriteLine("")

while len(strDllPath) = 0 
	strDllPath = InputBox("Please enter the location of the Print Manager components", _
               	              strMsgBoxTitle, _
                       	      "C:\Program Files\Marlborough Stirling\Omiga 4\DLL\")
wend

call InstallPackageEx("PrintManager", strDllPath)

strMsg = "InstallPrintManager complete"
MsgBox strMsg,0,strMsgBoxTitle
txtLogFile.WriteLine("")
txtLogFile.WriteLine(strMsg & " at " & Now())
txtLogFile.Close

sub InstallPackageEx(vstrPackageName, vstrDllPath)

	strMsg = "Installing package " & vstrPackageName & " ... "
	txtLogFile.WriteLine(strMsg)

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
	        packages.Remove(i)
	    End If
	Next
	
	' Commit our deletions
	packages.SaveChanges
	
	' Add a new package
	Dim newPackage
	Set newPackage = packages.Add
	newPackage.Value("Name") = vstrPackageName
	newPackage.Value("SecurityEnabled") = "N"

       ' AD 18/03/2003 change it back to be a server package - BMIDS00399

        newPackage.Value("Activation") = "Local"


'	while len(strIdentity) = 0 
'		strIdentity = InputBox("Please enter the identity the Print Manager will run under", _
'        	       	              strMsgBoxTitle, _
'                	       	      "")
'	wend
'	
'
'	while len(strPassword) = 0 
'		strPassword = InputBox("Please enter the password for the Print Manager identity", _
'        	       	              strMsgBoxTitle, _
'                	       	      "")
'	wend
'
'        newPackage.Value("Identity") = strIdentity
'
'        newPackage.Value("Password") = strPassword


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
		
	call InstallComponent(util, vstrDllPath & "omFVS.dll")
	call InstallComponent(util, vstrDllPath & "omDPS.dll")
	call InstallComponent(util, vstrDllPath & "omPM.dll")

	' Refresh components
	components.Populate
	
	' Commit the changes to the components
	components.SaveChanges

	strMsg = "Package " & vstrPackageName & " installed"
	txtLogFile.WriteLine(strMsg)

end sub
			
sub InstallComponent(vutil,vstrFileName)
	strMsg = "Installing component " & vstrFileName
	txtLogFile.WriteLine(strMsg)
	vutil.InstallComponent vstrFileName, "", ""
end sub
