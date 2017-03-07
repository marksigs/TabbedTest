Attribute VB_Name = "omPrintGlobals"
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
' IK_BM0200 traceAssist support
Public gobjTrace As traceAssist
Public Sub Main()
    
    ' IK_BM0200 traceAssist support
    Set gobjTrace = New traceAssist
    ' adoAssist
    adoBuildDbConnectionString
    adoLoadSchema
End Sub
