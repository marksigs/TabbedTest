Attribute VB_Name = "FileVersioningGlobals"
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
' ik_bm0200 traceAssist support
Public gobjTrace As traceAssist
Public Sub Main()
    
'   ik_bm0200 traceAssist support
    Set gobjTrace = New traceAssist
    ' adoAssist
    adoBuildDbConnectionString
    adoLoadSchema
End Sub
