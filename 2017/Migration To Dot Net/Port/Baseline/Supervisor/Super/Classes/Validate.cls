VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Validate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Function IsDateValid(sDate As String) As Boolean
    Dim sTmp As String
    
    sTmp = StripUnderscoreAndSlash(sDate)

    If Len(sTmp) > 0 Then
        IsDateValid = True
    Else
        IsDateValid = False
    End If

End Function


