Attribute VB_Name = "omLockManagerGlobals"
Option Explicit

' If this is declared in StdData.bas then all other components have to be rebuilt.
Public Const gstrLOCKMANAGER_COMPONENT = "omLockManager"

Public Sub Main()
    adoBuildDbConnectionString
End Sub

