Attribute VB_Name = "om4PayProcGlobals"
Option Explicit

Public Enum PPERROR
'    (omPayProc) Payment Processing specific errors
    oePPFeesOutstanding = 4700
    oePPNotAuthCancelBalance = 4701
    oePPComboValueIdNotFound = 4702
    oePPNotAuthUpdateDisb = 4703
    oePPNotAuthCreateDisb = 4704
    oePPBatchLockingErr = 4705
    oePPBatchPaymentPayeeInvalid = 4706
    oePPBatchNotInitAdvance = 4707
    oePPBatchInvalidJobType = 4708
End Enum


' PSC 19/09/2002 SYS4184
Public Enum BMidsPPERROR
    oePPBatchGPCompStageNotFound = 4711
    oePPBatchMoveToCompStageFailed = 4712
    oePPBatchStageMissed = 4713
    'BM0339 MDC 06/03/2003
    oePPBatchIncorrectPaymentStatus = 4714
    oePPBatchInvalidPaymentMethod = 4715
    oePPBatchMissingPayeeDetails = 4716
    'BM0339 MDC 06/03/2003 - End
End Enum

Public Sub Main()
    ' adoAssist
    adoLoadSchema
    adoBuildDbConnectionString
End Sub

Public Sub WriteRejectReport()
'#TASK - Printing
On Error GoTo WriteRejectReport_Exit

Const strFunctionName As String = "WriteRejectReport"

    'Do Reporting stuff

WriteRejectReport_Exit:

    'errCheckError strFunctionName, TypeName(Me)

End Sub

Public Function GetFactFindNumberForApplication(ByVal strAppNo As String) As String

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
Dim xmlTempNode As IXMLDOMNode
Dim xmlElement As IXMLDOMElement
Dim objAppBO As ApplicationBO
Dim objContext As ObjectContext

Dim lngRet As Long
Dim strResponse As String

    Set objContext = GetObjectContext()

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppBO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")

    'Get the ApplicationFactFindNumber for this application
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("APPLICATION")
    xmlTempRequest.appendChild xmlElement
    Set xmlTempNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlTempNode.Text = strAppNo
    xmlElement.appendChild xmlTempNode
    strResponse = objAppBO.GetApplicationData(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse

    GetFactFindNumberForApplication = xmlGetMandatoryNodeText(xmlDoc.documentElement, _
            "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/APPLICATIONFACTFINDNUMBER")

GetFactFindNumberForApplicationExit:
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempNode = Nothing
    Set xmlElement = Nothing
    Set objAppBO = Nothing
    Set objContext = Nothing
    Exit Function

End Function
'INR BMIDS628
Public Sub LogWarningMessage(ByVal strBatchNumber As String, ByVal strBatchRunNumber As String, _
                                                                    ByVal strMessage As String)

On Error GoTo LogWarningMessageExit

Dim strBatchLogFilePath As String
Dim intFileNo As Integer
Dim blnFileOpen As Boolean
Dim strBatchLogFile As String

    'Write an warning entry in the event log ====================================================
    App.LogEvent vbCrLf & strMessage, vbLogEventTypeWarning


    'Append an item to the batch run warnings log ===============================================
    strBatchLogFilePath = GetGlobalParamString("BatchLogFilePath")
    
    'Checking for the Directory Existence to write the outputFile
    'If does not Exist then Create One
    If Dir(strBatchLogFilePath, vbDirectory) = "" Then
        MkDir strBatchLogFilePath
    End If
    
    strBatchLogFile = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber & "_Warnings"
    
    intFileNo = FreeFile
    Open strBatchLogFile & ".log" For Append As #intFileNo
    blnFileOpen = True
    
    Print #intFileNo, ""
    Print #intFileNo, strMessage
    Print #intFileNo, "------------------------------------------------------"

LogWarningMessageExit:
    If blnFileOpen Then
        Close #intFileNo
    End If

End Sub


