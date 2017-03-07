Attribute VB_Name = "omPackDataGlobals"
'TW     18/10/2005      MAR223
'PSC    11/01/2006      MAR994 Add adoBuildDbConnectionString
'GHun   07/06/2006      MAR1819 Added Option Explicit and fixed module name
Option Explicit

Public gobjTrace As traceAssist
Public Sub Main()
    Set gobjTrace = New traceAssist
    adoBuildDbConnectionString
End Sub
