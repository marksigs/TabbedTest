' Filename: 	UninstallOmiga4.vbs

' Description: 
'	VB Script that deletes the Omiga 4 MTS package.
'------------------------------------------------------------------------------------------
'History:
'
'Prog  	Date     Description
'RF    	01/11/99 Created
'RF	05/11/99 Added QuickQuote package
'RF	10/11/99 Added CostModelling and MortgageProduct packages. 
'RF	29/02/00 Sole package is now "Omiga4".
'------------------------------------------------------------------------------------------

'call UninstallPackage("Audit")
'call UninstallPackage("Base")
'call UninstallPackage("CostModelling")
'call UninstallPackage("CustReg")
'call UninstallPackage("MortgageProduct")
'call UninstallPackage("QuickQuote")
'call UninstallPackage("Organisation")

call UninstallPackage("Omiga4")
MsgBox("End of UninstallOmiga4")

sub UninstallPackage(vstrPackageName)

	' First, we create the catalog object
	Dim catalog
	Set catalog = CreateObject("MTSAdmin.Catalog.1")
	
	' Then we get the packages collection
	Dim packages
	Set packages = catalog.GetCollection("Packages")
	packages.Populate
	
	' Remove all packages that go by the same name as that specified
	numPackages = packages.Count
	For i = numPackages - 1 To 0 Step -1
	    If packages.Item(i).Value("Name") = vstrPackageName Then
	        packages.Remove(i)
	    End If
	Next
	
	' Commit our deletions
	packages.SaveChanges
	
	MsgBox(vstrPackageName & " removed")

end sub
	
			