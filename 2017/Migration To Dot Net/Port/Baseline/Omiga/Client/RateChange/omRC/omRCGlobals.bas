Attribute VB_Name = "omRCGlobals"
Option Explicit

Public Enum RCERROR
    ' Rate Change specific errors
    oeRCBatchLockingError = 5000
End Enum



Public Sub Main()
    ' adoAssist
    adoLoadSchema
    adoBuildDbConnectionString
End Sub

Public Sub WriteRejectReport()
'BMIDS622 WriteRejectReport does not do anything, but cannot be deleted because OOSS will just replace it
''#TASK - Printing
'On Error GoTo WriteRejectReport_Exit
'
'Const cstrFunctionName As String = "WriteRejectReport"
'
'    'Do Reporting stuff
'
'WriteRejectReport_Exit:
'
'    'errCheckError strFunctionName, TypeName(Me)
'
End Sub


