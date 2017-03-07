Attribute VB_Name = "omPDMGlobal"
Option Explicit
'TW     18/10/2005  MAR223
Public gobjTrace As traceAssist
Public Sub Main()
    Set gobjTrace = New traceAssist
    DBAssistBuildDbConnectionString
End Sub
