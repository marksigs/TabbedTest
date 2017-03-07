'Attribute VB_Name = "Module1"
' Filename:     InstallOmiga4.vbs

' Description:
'       VB Script that creates Omiga 4 packages and installs the appropriate components into them.
'       If the package already exists it is deleted and recreated.
'       A log of packages installed is written to InstallOmiga4.log in the working directory.
'       Completion of the installation is indicated via a messagebox.
'       Running of unit test/developer version is controlled by flag blnTestEnviron.
'       The script was based on InstDLL.vbs, a file provided as part of the Microsoft Transaction Server Samples.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'RF     01/11/99        Created
'RF     01/11/99        "Base" is created as a library package
'RF     05/11/99        Added QuickQuote package
'RF     08/11/99        Write to a log file ("InstallOmiga4.log") in working directory
'RF     08/11/99        Added extra logging information.
'RF     10/11/99        Added CostModelling and MortgageProduct packages in build environment.
'                       Allowed for different dll paths for development and build installations.
'RF     10/11/99        Tested build version.
'RF     22/11/99        Added CostModelling and MortgageProduct packages in developer environment.
'RF     29/11/99        New dll path in test environment.
'RF     01/12/99        Added new function InstallPackageEx as all components should be in the same package.
'RF     01/12/99        InstallPackageEx now uses new function InstallComponent.
'RF     09/12/99        Added AIP component in test environment.
'RF     24/01/00        Change to dll directory in test environment.
'RF     11/02/00        Added BC component in test environment.
'APS    15/03/00        Added PAF component in test environment.
'PSC    30/03/00        Added omCC component
'AY     13/04/00        Added om4To3,omEXP,omPoke,omRA,omRARules,omRB and omSub components
'APS    25/05/00        Removed omCR component
'JLD    25/05/00        Added omHU Egg HomeUse component
'JLD    11/08/00        Added omDM core Decision Manager component
'PSC    05/09/00        Enable input of component location
'PSC    05/09/00        Change so that component location is read as the first parameter else
'                       ask for it
'PSC    18/12/00        Add Task Management components
'IK     20/12/00        replace omTaskManagement component with omTm components
'RF     17/01/01        Added omDPS and omPM components
'IVW    18/01/01        Added new payproc.dll to MTS installation.
'IVW    02/02/01        Added new component OmTmRules.
'PF     23/02/01        Removed omPoke and added omAppProc SYS1978
'PF     27/02/01        Removed omDPS and omPM SYS1983
'GF     14/03/01        Added omCDRules and omROT SYS2065
'PC     27/03/01        SYS2180 Add omAdmin
'GF     18/04/01        SYS1878 add new component omAPRules
'GF     02/05/01        SYS2247 & SYS2298 added omPrint, omPDM, omPC to installation
'GF     03/05/01        SYS2247 Remove omPC - not a .dll!
'GF     04/05/01        SYS2247 Install omToOMMQ.dll ( requires .tlb )
'PF     21/05/01        SYS2288 Remove omHU
'RF     23/01/02        SYS3812 Added new component ODITransformer.
'DM	21/02/02	SYS4130 Added omToCBA because it has not been included in MTS
'DM     25/02/02        SYS4146 Added omFromOmmq because it was not added into COM+ Services.
'RF	05/03/02	SYS4215 Added omMQ
'SG	29/05/02	SYS4767 MSMS to Core integration - added omBACS
'
'BMids Changes
'~~~~~~~~~~~~~
'AW	17/05/02	Removed components not required for BMIDS
'MDC	28/06/02	BMIDS00121 - Added omHI
'AW	11/07/02	BMIDS00189 - Re-added omIM
'AD	21/08/02	SYS5412 - Added in omCaseTrack
'AD	06/09/02	BMIDS00218 - Added in the 2 Data Ingestion components.
'AD	17/09/02	Removed references to MTS
'MV	04/10/02	BMIDS00502 - Added omInt
'PSC	18/10/02	BMIDS00674 - Amend to remove surplus components
'AD	05/11/2002	BMIDS00848 - Remove components
'MDC 05/11/2002 BMIDS00654 - Add new component omIC
'AD	18/03/03	BM0399 Added in omMutex.
'AD 	15/05/03	BM00549 - Make sure it is a server package.
'------------------------------------------------------------------------------------------


Dim blnTestEnviron

Dim strMsg

' Get the component path if passed in as an argument

Set objArgs = WScript.Arguments

If objArgs.Count > 0 Then
        strDllPath = objArgs(0)
End If

'*****************************************
'todo - comment/uncomment the appropriate lines
'*****************************************
blnTestEnviron = True           ' test environment (e.g unit test or system test)
'blnTestEnviron = false         ' developer environment

Set fso = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = fso.CreateTextFile("InstallOmiga4.log", True)

If blnTestEnviron = True Then
        strMsg = "InstallOmiga4 starting for test environment"
Else
        strMsg = "InstallOmiga4 starting for developer environment"
End If

MsgBox strMsg, 0, "Omiga 4 Installation"
txtLogFile.WriteLine (strMsg & " at " & Now())
txtLogFile.WriteLine ("")

If blnTestEnviron = True Then

        While Len(strDllPath) = 0
                strDllPath = InputBox("Please enter the location of the Omiga 4 components", _
                                      "Omiga 4 Installation", _
                                      "C:\Program Files\Marlborough Stirling\Omiga 4\DLL\")
        Wend

        Call InstallPackageEx("Omiga4", strDllPath)

Else
        While Len(strDllPath) = 0
                strDllPath = InputBox("Please enter the location of the Omiga 4 components", _
                                      "Omiga 4 Installation", _
                                      "C:\PROJECTS\dev\")
        Wend

        Call InstallPackage("Audit", strDllPath & "Audit\omAU\omAU.dll")
        Call InstallPackage("Base", strDllPath & "Base\omBase\omBase.dll")
        Call InstallPackage("CostModelling", strDllPath & "CostModelling\omCM\omCM.dll")
        'call InstallPackage("CustReg", strDllPath & "custreg\middletier\omiga4CustReg\omCR.dll")
        Call InstallPackage("MortgageProduct", strDllPath & "MortgageProduct\omMP\omMP.dll")
        Call InstallPackage("QuickQuote", strDllPath & "QuickQuote\omiga4QuickQuote\omQQ.dll")
        Call InstallPackage("Organisation", strDllPath & "Organisation\omOrg\omOrg.dll")
End If

strMsg = "InstallOmiga4 complete"
MsgBox strMsg, 0, "Omiga 4 Installation"
txtLogFile.WriteLine ("")
txtLogFile.WriteLine (strMsg & " at " & Now())
txtLogFile.Close


Sub InstallPackage(vstrPackageName, vstrDllPath)

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

        If vstrPackageName = "Base" Then
                ' make it a library package
                newPackage.Value("Activation") = "Local"
        End If
        
        ' Commit new package
        packages.SaveChanges
        
        ' Refresh packages
        packages.Populate
        
        ' Get components collection for new package
        Dim components
        Set components = packages.GetCollection("ComponentsInPackage", newPackage.Value("ID"))
        
        ' Install components
        Dim util
        Set util = components.GetUtilInterface
        
        util.InstallComponent vstrDllPath, "", ""
        
        ' Refresh components
        components.Populate
        
        ' Commit the changes to the components
        components.SaveChanges

        'MsgBox vstrPackageName & " installed",0, "Omiga 4 Installation"
        strMsg = vstrPackageName & " installed"
        txtLogFile.WriteLine (strMsg)

End Sub


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
        newPackage.Value("Activation") = "Local"

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
                
 
        Call InstallComponent(util, vstrDllPath & "om4To3.dll")
        Call InstallComponent(util, vstrDllPath & "omAdmin.dll")
        Call InstallComponent(util, vstrDllPath & "omAIP.dll")
        Call InstallComponent(util, vstrDllPath & "omApp.dll")
        Call InstallComponent(util, vstrDllPath & "omAQ.dll")
        Call InstallComponent(util, vstrDllPath & "omAU.dll")
        Call InstallComponent(util, vstrDllPath & "omBase.dll")
        Call InstallComponent(util, vstrDllPath & "omBC.dll")
        Call InstallComponent(util, vstrDllPath & "omCC.dll")
        Call InstallComponent(util, vstrDllPath & "omCE.dll")
        Call InstallComponent(util, vstrDllPath & "omCF.dll")
        Call InstallComponent(util, vstrDllPath & "omCM.dll")
        Call InstallComponent(util, vstrDllPath & "omCust.dll")
        Call InstallComponent(util, vstrDllPath & "omEXP.dll")
        Call InstallComponent(util, vstrDllPath & "omLC.dll")
        Call InstallComponent(util, vstrDllPath & "omMP.dll")
        Call InstallComponent(util, vstrDllPath & "omMQ.dll")
        Call InstallComponent(util, vstrDllPath & "omOrg.dll")
        Call InstallComponent(util, vstrDllPath & "omPAF.dll")
        Call InstallComponent(util, vstrDllPath & "omPP.dll")
        Call InstallComponent(util, vstrDllPath & "omQQ.dll")
        Call InstallComponent(util, vstrDllPath & "omRA.dll")
        Call InstallComponent(util, vstrDllPath & "omRB.dll")
        Call InstallComponent(util, vstrDllPath & "omTP.dll")
        Call InstallComponent(util, vstrDllPath & "omTM.dll")
        Call InstallComponent(util, vstrDllPath & "msgTM.dll")
        Call InstallComponent(util, vstrDllPath & "omPayProc.dll")
        Call InstallComponent(util, vstrDllPath & "omComp.dll")
        Call InstallComponent(util, vstrDllPath & "omAppProc.dll")
        Call InstallComponent(util, vstrDllPath & "omROT.dll")
        Call InstallComponent(util, vstrDllPath & "omPrint.dll")
        Call InstallComponent(util, vstrDllPath & "omToOMMQ.dll")
        Call InstallComponent(util, vstrDllPath & "omCOG.dll")
        Call InstallComponent(util, vstrDllPath & "omBatch.dll")
        Call InstallComponent(util, vstrDllPath & "omRC.dll") 'Added by Tony Wilkes 21 March 2002 SYS4303
	Call InstallComponent(util, vstrDllPath & "omFromOmmq.dll")
	Call InstallComponent(util, vstrDllPath & "omBACS.dll")	'SG 29/05/02 SYS4767 
	Call InstallComponent(util, vstrDllPath & "omHI.dll")	'MDC 28/06/02 BMIDS00121
	Call InstallComponent(util, vstrDllPath & "omIM.dll")   'AW	11/07/02 BMIDS00189
	Call InstallComponent(util, vstrDllPath & "omCaseTrack.dll")   'AD	21/08/02 SYS5412
	Call InstallComponent(util, vstrDllPath & "omDataIngestion.dll")	'AD 06/09/02 BMIDS00218 
	Call InstallComponent(util, vstrDllPath & "omDataIngestionRules.dll")	'AD 06/09/02 BMIDS00218 
	Call InstallComponent(util, vstrDllPath & "omInt.dll")  'MV 04/10/2002 BMIDS00502
	Call InstallComponent(util, vstrDllPath & "omCMRules.dll")  'PSC 18/10/2002 BMIDS00674
	Call InstallComponent(util, vstrDllPath & "omIC.dll")  'BMIDS00654 MDC 05/11/2002
	Call InstallComponent(util, vstrDllPath & "omMutex.dll")  'BM0399 MDC 18/03/2003


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