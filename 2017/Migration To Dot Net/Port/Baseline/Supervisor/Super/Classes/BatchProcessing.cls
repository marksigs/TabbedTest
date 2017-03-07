VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : BatchProcessing
' Description   : Contains all batch processes
'
' Change history
' Prog      Date        Description
' AA        27/06/2001  Created
' SA        10/01/2002  SYS3327 omBatch.Launchbatch requires attribute based xml
'                        Pass in Date/Time to Launch method
' STB       05/02/2002  SYS3327 omBatch.Launchbatch requires UserID, UnitID and
'                       ChannelID in the request.
' SA        13/02/2002  Add new parameter to LaunchBatch to tell if it's a "rerun" or not.
' STB       10/04/2002  SYS4357 DateTime datatypes are implicitly converted to Amertican format.
' GHun      15/04/2002  SYS4368 Changed Launch ReRun parameter to BatchRunNo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Launch
' Description   : Calls the omiga4 Batch Component to place a batch on the queue
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' SYS4368 Changed Launch bReRun parameter to lngBatchRunNo
'Public Sub Launch(vBatchNo As Variant, dExecutionDateTime As Date, Optional bReRun As Boolean)
Public Sub Launch(vBatchNo As Variant, dExecutionDateTime As Date, Optional lngBatchRunNo As Long = -1)

    On Error GoTo Failed
    Dim objXMLBatch As New FreeThreadedDOMDocument
    Dim objRequestElement As IXMLDOMElement
    Dim objListElement As IXMLDOMElement
    Dim objBatchElement As IXMLDOMElement
    Dim objXMLResp As New FreeThreadedDOMDocument
    Dim nResponse As Integer
    Dim clsOmiga4 As New Omiga4Support
    Dim xmlResponseNode As IXMLDOMNode
    Dim sResponse As String
        
    Set objRequestElement = objXMLBatch.createElement("REQUEST")

    'SYS3327
    'objRequestElement.appendChild(objXMLBatch.createElement("OPERATION")).Text = "LaunchBatch"
    objRequestElement.setAttribute "OPERATION", "LaunchBatch"
    objRequestElement.setAttribute "USERID", g_sSupervisorUser
    objRequestElement.setAttribute "UNITID", g_sUnitID
    objRequestElement.setAttribute "CHANNELID", g_sChannelID
    objRequestElement.setAttribute "USERAUTHORITYLEVEL", 99
    objRequestElement.setAttribute "MACHINEID", GetMachineID()
    
    Set objBatchElement = objXMLBatch.createElement("BATCH")
    
    'SYS3327 Should be an attribute and should have date/time too
    'objBatchElement.appendChild(objXMLBatch.createElement("BATCHNUMBER")).Text = vBatchNo
    objBatchElement.setAttribute "BATCHNUMBER", vBatchNo
    
    'SYS4357 - DateTimes are implicitly converted to American format in the DOM.
    objBatchElement.setAttribute "EXECUTIONDATETIME", Format$(dExecutionDateTime, "dd/mm/yyyy Hh:Nn:ss")
    'SYS4357 - End.
    
    'SYS4073 Send in FailedBatchRunNumber if a rerun.
    'SYS4368 Use the FailedBatchRunNumber, not the BatchNumber
    'If a FailedBatchRunNumber has been passed in, it should be greater than 0
    If lngBatchRunNo > 0 Then
        'objBatchElement.setAttribute "FAILEDBATCHNUMBER", vBatchNo
        objBatchElement.setAttribute "FAILEDBATCHRUNNUMBER", lngBatchRunNo
    End If
    'SYS4368 End
    
    objRequestElement.appendChild objBatchElement
    
    objXMLBatch.appendChild objRequestElement
    
    sResponse = clsOmiga4.RunASP(objXMLBatch.xml, "LaunchBatch.asp")
    
    ' Check the return code
    ' DJP SQL Server port, use Supervisor error object instead
    g_clsXMLAssist.CheckXMLResponse sResponse, True

    ' Check the return code
    nResponse = g_clsXMLAssist.CheckXMLResponse(sResponse, True)

    If nResponse <> 0 Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

