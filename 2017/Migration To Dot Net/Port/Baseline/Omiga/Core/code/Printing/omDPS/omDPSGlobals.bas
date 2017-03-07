Attribute VB_Name = "omDPSGlobals"
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' IK    17/02/2003  BM0200 - add TraceAssist support
' PSC   20/09/2004  BBG1438 Added LoadPrinterAttributes
'------------------------------------------------------------------------------------------
' ik_bm0200
Option Explicit
Public gobjTrace As traceAssist

Public gxmldocPrinterAttributes As FreeThreadedDOMDocument40 ' PSC 20/09/2004 BBG1438

Public Sub Main()
    
    ' ik_bm0200
    Set gobjTrace = New traceAssist
    adoBuildDbConnectionString
    LoadPrinterAttributes       ' PSC 20/09/2004 BBG1438
End Sub

' PSC 20/09/2004 BBG1438 - Start
Private Sub LoadPrinterAttributes()

    Dim strFileSpec As String
    
    Set gxmldocPrinterAttributes = xmlCreateDOMObject()
    
    strFileSpec = App.Path & "\PrinterAttributes.xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    gxmldocPrinterAttributes.load strFileSpec

End Sub
' PSC 20/09/2004 BBG1438 - End
