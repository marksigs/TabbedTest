Attribute VB_Name = "omCMGlobals"
'-------------------------------------------------------------------------
'BMids Specific History:
'Prog   Date        AQR      Decription
'GHun   23/04/2004  BMIDS736 Added oeAlphaPlusError
'-------------------------------------------------------------------------
Option Explicit

Public Enum INVALIDINCENTIVEAMOUNT
    oeInvalidIncentiveAmount = 7008
End Enum

'BMIDS736 GHun
Public Const oeAlphaPlusError As Integer = 9000
'BMIDS736 End

Private Sub Main()
    adoBuildDbConnectionString
End Sub
