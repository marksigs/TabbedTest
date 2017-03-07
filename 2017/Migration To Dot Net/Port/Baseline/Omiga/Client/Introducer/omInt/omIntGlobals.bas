Attribute VB_Name = "omIntGlobals"
'Workfile:      omIntGlobals.bas
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Introducer Processing General module.

'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MO     20/09/2002  Created
'-------------------------------------------------------------------------------------------------------

Option Explicit

Public Sub Main()
    ' adoAssist
    adoLoadSchema
    adoBuildDbConnectionString
End Sub
