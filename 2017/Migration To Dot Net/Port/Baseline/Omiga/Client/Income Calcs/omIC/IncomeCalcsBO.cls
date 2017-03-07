VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IncomeCalcsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      IncomeCalcsBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Income Calculations Business Object which 'supports transactions' only.
'               Any transactioned requests will be delegated to IncomeCalcsTxBO

'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'-------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
'
' MDC   03/07/02    BMIDS00654 Maximum Borrowing Calculations
'-------------------------------------------------------------------------------------------
'Mars History:
'
' Prog  Date        Description
'
' GHun  10/07/2006  MAR1891 Changed RunIncomeCalcs and CalculateMaximumBorrowing to reduce locking
'-------------------------------------------------------------------------------------------
'Epsom History:
'
' Prog  Date        Description
'
' PSC   19/02/2007 EP2_1409 Return warning messages from RunIncomeCalcs
'-------------------------------------------------------------------------------------------

Implements IIncomeCalcsBO
Implements ObjectControl

Private m_objXMLAssist As XMLAssist
Private m_objErrAssist As ErrAssist
Private m_objContext As ObjectContext

Public Function RunIncomeCalculation(ByVal vstrXMLRequest As String) As String

On Error GoTo RunIncomeCalculationVbErr

    Const strFunctionName As String = "RunIncomeCalculation"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    'Create default response block
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    'Delegate to DOM based method
    Set xmlTempResponseNode = IIncomeCalcsBO_RunIncomeCalculation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RunIncomeCalculation = xmlTempResponseNode.xml
    
    m_objContext.SetComplete

RunIncomeCalculationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

RunIncomeCalculationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    RunIncomeCalculation = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RunIncomeCalculationExit

End Function

Public Function CalculateAllowableIncome(ByVal vstrXMLRequest As String) As String

On Error GoTo CalculateAllowableIncomeVbErr

    Const strFunctionName As String = "CalculateAllowableIncome"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    'Create default response block
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    'Delegate to DOM based method
    Set xmlTempResponseNode = IIncomeCalcsBO_CalculateAllowableIncome(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateAllowableIncome = xmlTempResponseNode.xml
    
    m_objContext.SetComplete

CalculateAllowableIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

CalculateAllowableIncomeVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAllowableIncome = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateAllowableIncomeExit

End Function

Public Function CalculateMaximumBorrowing(ByVal vstrXMLRequest As String) As String

On Error GoTo CalculateMaximumBorrowingVbErr

Const strFunctionName As String = "CalculateMaximumBorrowing"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    'Create default response block
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    'Delegate to DOM based method
    Set xmlTempResponseNode = IIncomeCalcsBO_CalculateMaximumBorrowing(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateMaximumBorrowing = xmlResponseElem.xml
    
    m_objContext.SetComplete

CalculateMaximumBorrowingExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalculateMaximumBorrowingVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateMaximumBorrowing = m_objErrAssist.CreateErrorResponse

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function

Public Function GetIncomeMultipleSet(ByVal vstrXMLRequest As String) As String

On Error GoTo GetIncomeMultipleSetVbErr

    Const strFunctionName As String = "GetIncomeMultipleSet"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = xmlCreateDOMObject()
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DOM based method and attach returned data to our response
    Set xmlTempResponseNode = IIncomeCalcsBO_GetIncomeMultipleSet(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXMLAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    GetIncomeMultipleSet = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetIncomeMultipleSetExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetIncomeMultipleSetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If

    GetIncomeMultipleSet = m_objErrAssist.CreateErrorResponse

    
    ' go to clean-up section
    Resume GetIncomeMultipleSetExit

End Function


Private Function IIncomeCalcsBO_CalculateAllowableIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAllowableIncomeVbErr
    
    Const strFunctionName As String = "IIncomeCalcsBO_CalculateAllowableIncome"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIIncomeCalcsTxBO As IIncomeCalcsTxBO

    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIIncomeCalcsTxBO = m_objContext.CreateInstance(App.Title & ".IncomeCalcsTxBO")
    
    'Validate that required elements are present
    ValidateIncomeCalcsRequest vxmlRequest
    
    Set xmlTempResponseNode = objIIncomeCalcsTxBO.CalculateAllowableIncome(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IIncomeCalcsBO_CalculateAllowableIncome = xmlTempResponseNode
    m_objContext.SetComplete
    
CalculateAllowableIncomeExit:
    Set objIIncomeCalcsTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CalculateAllowableIncomeVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsBO_CalculateAllowableIncome = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateAllowableIncomeExit

End Function

Private Function IIncomeCalcsBO_CalculateMaximumBorrowing(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateMaximumBorrowingVbErr
        
    Const strFunctionName As String = "IIncomeCalcsBO_CalculateMaximumBorrowing"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIIncomeCalcsTxBO As IIncomeCalcsTxBO
    Dim xmlRequest As IXMLDOMElement
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set objIIncomeCalcsTxBO = m_objContext.CreateInstance(App.Title & ".IncomeCalcsTxBO")
    
    'Validate that required elements are present
    ValidateIncomeCalcsRequest vxmlRequest
    
    'MAR1891 GHun Get MaxBorrowing data up front before a new transaction is created
    Set xmlRequest = vxmlRequest.cloneNode(True)
    
    Set xmlTempResponseNode = GetMaxBorrowing(vxmlRequest)
    If Not xmlTempResponseNode Is Nothing Then
        xmlRequest.appendChild xmlTempResponseNode
    End If
    
    Set xmlTempResponseNode = objIIncomeCalcsTxBO.CalculateMaximumBorrowing(xmlRequest)
    'MAR1891 End
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IIncomeCalcsBO_CalculateMaximumBorrowing = xmlResponseElem
    m_objContext.SetComplete
    
CalculateMaximumBorrowingExit:
    Set objIIncomeCalcsTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CalculateMaximumBorrowingVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsBO_CalculateMaximumBorrowing = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function

Private Sub ValidateIncomeCalcsRequest(ByVal vxmlRequest As IXMLDOMNode)

On Error GoTo ValidateIncomeCalcsRequestVbErr

Const strFunctionName As String = "ValidateIncomeCalcsRequest"

Dim strValue As String

    strValue = m_objXMLAssist.GetMandatoryElementText(vxmlRequest, "//APPLICATION/APPLICATIONNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(vxmlRequest, "//APPLICATION/APPLICATIONFACTFINDNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(vxmlRequest, "//CUSTOMERLIST/CUSTOMER/CUSTOMERNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(vxmlRequest, "//CUSTOMERLIST/CUSTOMER/CUSTOMERVERSIONNUMBER")
    strValue = m_objXMLAssist.GetMandatoryElementText(vxmlRequest, "//CUSTOMERLIST/CUSTOMER/CUSTOMERROLETYPE")

    Exit Sub
    
ValidateIncomeCalcsRequestVbErr:
    ' Re-raise the error
    Err.Raise Err.Number, strFunctionName, Err.Description

End Sub

Private Function IIncomeCalcsBO_GetIncomeMultipleSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetIncomeMultipleSetVbErr
    
    Const strFunctionName As String = "IIncomeCalcsBO_GetIncomeMultipleSet"
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIIncomeCalcDO As IIncomeCalcsDO
    
    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIIncomeCalcDO = m_objContext.CreateInstance(App.Title & ".IncomeCalcsDO")
                
    If vxmlRequest.nodeName = "INCOMEMULTIPLESET" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("INCOMEMULTIPLESET").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "INCOMEMULTIPLESET" & " tag not found"
    End If
        
    Set xmlDataNode = objIIncomeCalcDO.GetIncomeMultipleSet(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode

    Set IIncomeCalcsBO_GetIncomeMultipleSet = xmlResponseElem.cloneNode(True)
    
    m_objContext.SetComplete
    
GetIncomeMultipleSetExit:
    
    Set xmlRequestNode = Nothing
    Set objIIncomeCalcDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetIncomeMultipleSetVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If

    Set IIncomeCalcsBO_GetIncomeMultipleSet = m_objErrAssist.CreateErrorResponseEx()

    ' go to clean-up section
    Resume GetIncomeMultipleSetExit

End Function

Private Function IIncomeCalcsBO_RunIncomeCalculation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo RunIncomeCalculationVbErr
    
    Const strFunctionName As String = "IIncomeCalcsBO_RunIncomeCalculation"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIIncomeCalcsTxBO As IIncomeCalcsTxBO
    'MAR1891 GHun
    Dim objApplicationQuoteBO As ApplicationQuoteBO
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim strResponse As String
    Dim xmlRequest As IXMLDOMElement
    'MAR1891 End

    Set xmlOut = xmlCreateDOMObject()
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Validate that required elements are present
    ValidateIncomeCalcsRequest vxmlRequest
    
    'MAR1891 GHun Run CalcAllowableInc up front before a new transaction is created
    Set objApplicationQuoteBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    strResponse = objApplicationQuoteBO.CalcAllowableInc(vxmlRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    m_objErrAssist.CheckResponse xmlTempResponseDoc.documentElement, xmlResponseElem, True
    
    Set xmlTempResponseNode = xmlOut.createElement("CALCALLOWABLEINC")
    xmlTempResponseNode.appendChild xmlTempResponseDoc.documentElement
    Set xmlRequest = vxmlRequest.cloneNode(True)
    xmlRequest.appendChild xmlTempResponseNode
    
    'Get MaxBorrowing data up front before a new transaction is created
    Set xmlTempResponseNode = GetMaxBorrowing(vxmlRequest)
    If Not xmlTempResponseNode Is Nothing Then
        xmlRequest.appendChild xmlTempResponseNode
    End If
    'MAR1891 End
    
    Set objIIncomeCalcsTxBO = m_objContext.CreateInstance(App.Title & ".IncomeCalcsTxBO")
    Set xmlTempResponseNode = objIIncomeCalcsTxBO.RunIncomeCalculation(xmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' PSC 19/02/2007 EP2_1409 - Start
    AttachResponseData xmlResponseElem, xmlTempResponseNode
    Set IIncomeCalcsBO_RunIncomeCalculation = xmlResponseElem
    ' PSC 19/02/2007 EP2_1409 - End
    
    m_objContext.SetComplete
    
RunIncomeCalculationExit:
    Set objIIncomeCalcsTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    'MAR1891 GHun
    Set objApplicationQuoteBO = Nothing
    Set xmlTempResponseDoc = Nothing
    'MAR1891 End
    
    Exit Function

RunIncomeCalculationVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IIncomeCalcsBO_RunIncomeCalculation = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume RunIncomeCalculationExit

End Function

'MAR1891 GHun
Private Function GetMaxBorrowing(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMElement
    Const cstrFunctionName  As String = "GetMaxBorrowing"
    Dim xmlElem             As IXMLDOMElement
    Dim objParameterDO      As GlobalParameterDO
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim blnContinue         As Boolean
    Dim strTemp             As String
    
On Error GoTo GetMaxBorrowingErr
    blnContinue = False
    
    Set xmlElem = xmlGetMandatoryNode(vxmlRequest, "INCOMECALCULATION")
    If xmlGetAttributeText(xmlElem, "CALCULATEMAXBORROWING") = "1" Then
        blnContinue = True
    Else
        Set objParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        If objParameterDO.GetCurrentParameterByType("IncMultToleranceChk", "BOOLEAN") Then
            strTemp = objParameterDO.GetCurrentParameterByType("IncomeTolerenceMethod", "STRING")
            If UCase$(strTemp) <> "MAXBORROW" Then
                blnContinue = True
            End If
        End If
    End If
    
    If blnContinue Then
        Set xmlElem = xmlGetMandatoryNode(xmlElem, "APPLICATION")
        strTemp = GetMaxBorrowingValues(xmlGetNodeText(xmlElem, "APPLICATIONNUMBER"), _
                        xmlGetNodeText(xmlElem, "APPLICATIONFACTFINDNUMBER"))
        Set xmlDoc = xmlLoad(strTemp, cstrFunctionName)
        
        Set GetMaxBorrowing = xmlGetNode(xmlDoc, "RESPONSE/MAXBORROWINGVALUES")
    Else
        Set GetMaxBorrowing = Nothing
    End If
    
GetMaxBorrowingErr:
    Set xmlElem = Nothing
    Set objParameterDO = Nothing
    Set xmlDoc = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'MAR1891 End

'MAR1891 GHun This function has been moved here, rather than the DO, as I don't want to change the interface for a patch release
Private Function GetMaxBorrowingValues(ByVal vstrApplicationNumber As String, ByVal vstrApplicationFactFindNumber As String) As String
    Const cstrFunctionName As String = "GetMaxBorrowingValues"

    Dim objIADOAssist  As IADOAssist
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    
On Error GoTo GetMaxBorrowingValuesExit
    
    Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objIADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .open
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.open
       
    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandText = "USP_GETMAXBORROWINGVALUES"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(vstrApplicationNumber), vstrApplicationNumber)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , CInt(vstrApplicationFactFindNumber))
        Set .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
       
    adoConnection.Close
    
    GetMaxBorrowingValues = adoStream.ReadText()
    
    adoStream.Close
        
GetMaxBorrowingValuesExit:
    If adoConnection.State = adStateOpen Then
       adoConnection.Close
    End If
    
    Set adoConnection = Nothing
    Set objIADOAssist = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
'MAR1891 End

Private Sub ObjectControl_Activate()

    Set m_objXMLAssist = New XMLAssist
    Set m_objErrAssist = New ErrAssist
    Set m_objContext = GetObjectContext()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objXMLAssist = Nothing
    Set m_objErrAssist = Nothing
    Set m_objContext = Nothing

End Sub
