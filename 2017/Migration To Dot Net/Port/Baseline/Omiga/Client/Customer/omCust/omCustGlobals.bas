Attribute VB_Name = "omCustGlobals"
Option Explicit


'------------------------------------------------------------------------------------------
'BMids Specific History:
'
'Prog   Date        Description
'GHun   26/06/2002  BMIDS00005 Added oeCustNoContactsFound error constant
'------------------------------------------------------------------------------------------

Public Enum CUSTOMERERROR
    oeCustNoContactsFound = 7000
End Enum

Public Sub Main()
    adoBuildDbConnectionString
End Sub



