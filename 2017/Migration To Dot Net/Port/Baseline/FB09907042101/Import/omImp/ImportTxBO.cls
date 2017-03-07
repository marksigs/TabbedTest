VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ImportTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ImportTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     23/03/00    Creation
'SR     03/05/00    New method 'CreateApplication'
'SR     08/05/00    New Method 'UpdateApplication'
'SR     05/06/00    Modified method 'UpdateApplication' - Now calls 'UpdateSubmissionStatus'
'                   instead of 'CreateSubmissionStatus'
'MC     20/06/00    Unlock application and customers when completed CreateApplication
'------------------------------------------------------------------------------------------

Option Explicit

Implements IImportTxBO
Implements ObjectControl
 
'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Private Function IImportTxBO_CreateApplication( _
        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IImportTxBO_CreateApplication"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCustNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode, xmlTableNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlOutCustomerList As IXMLDOMNode
    
    Dim strResponse As String, strCustomerNumber As String
    Dim strCustVersionNumber As String
    Dim strApplNumber As String, strAFFNumber As String, strTypeOfApplication As String
    
    Dim objCustomerBO As CustomerBO
    Dim objApplManagerBO As ApplicationManagerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objCustomerBO = New CustomerBO
        Set objApplManagerBO = New ApplicationManagerBO
    Else
        Set objCustomerBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objApplManagerBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    End If
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    '-----------------------------------------------------------------------------------
    'Create Full Application Details and store the number of the application generated
    '-----------------------------------------------------------------------------------
    strResponse = objApplManagerBO.CreateFullApplicationDetails(vxmlRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    strApplNumber = xmlDoc.selectSingleNode("//APPLICATIONNUMBER").Text
    strAFFNumber = xmlDoc.selectSingleNode("//APPLICATIONFACTFINDNUMBER").Text
    
    ' Add Application number and FactFind number to Response
    Set xmlTableNode = xmlOut.createElement("APPLICATION")
    xmlResponseElem.appendChild xmlTableNode
    
    Set xmlElement = xmlTableNode.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlElement.Text = strApplNumber
    
    Set xmlElement = xmlTableNode.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlElement.Text = strAFFNumber
    '-------------------------------------------------------------------------
    ' Create Full Customer Details and add this customer to the
    ' application generated above
    '-------------------------------------------------------------------------
    Dim intOrder As Integer
    intOrder = 0
    
    Set xmlOutCustomerList = xmlOut.createElement("CUSTOMERLIST")
    xmlResponseElem.appendChild xmlOutCustomerList
    
    strTypeOfApplication = m_objXmlAssist.GetTagValue(vxmlRequest, "TYPEOFAPPLICATION")
    
    Set xmlCustNodeList = vxmlRequest.selectNodes(".//CUSTOMER")
    For Each xmlNode In xmlCustNodeList
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("CREATE"))
        xmlElement.appendChild xmlNode
        strResponse = objCustomerBO.CreateFullCustomerDetails(xmlNewRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
        strCustomerNumber = xmlDoc.selectSingleNode("//CUSTOMERNUMBER").Text
        strCustVersionNumber = xmlDoc.selectSingleNode("//CUSTOMERVERSIONNUMBER").Text
        
        ' Add above generated values to Response
        Set xmlTableNode = xmlOut.createElement("CUSTOMER")
        xmlOutCustomerList.appendChild xmlTableNode
        
        Set xmlElement = xmlTableNode.appendChild(xmlOut.createElement("CUSTOMERNUMBER"))
        xmlElement.Text = strCustomerNumber
        
        Set xmlElement = xmlTableNode.appendChild(xmlOut.createElement("CUSTOMERVERSIONNUMBER"))
        xmlElement.Text = strCustVersionNumber
        
        ' Build the request and add customer to application
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
        Set xmlTableNode = xmlOut.createElement("APPLICATION")
        xmlNewRequest.appendChild xmlTableNode
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
        xmlElement.Text = strApplNumber
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
        xmlElement.Text = strAFFNumber
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("CUSTOMERNUMBER"))
        xmlElement.Text = strCustomerNumber
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("CUSTOMERVERSIONNUMBER"))
        xmlElement.Text = strCustVersionNumber

        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("CUSTOMERROLETYPE"))
        xmlElement.Text = "1"  'Customer is an applicant
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("CUSTOMERORDER"))
        intOrder = intOrder + 1
        xmlElement.Text = intOrder
        
        Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("TYPEOFAPPLICATION"))
        xmlElement.Text = strTypeOfApplication
        
        strResponse = objApplManagerBO.AddCustomerToApplication(xmlNewRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
        Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Next xmlNode

' ** FIX ME LOAN DETAILS (Awaiting completion of spec)

    'Create Application Stage record
    Set xmlElement = CreateApplicationStage(vxmlRequest, strApplNumber, strAFFNumber)
    m_objErrAssist.CheckResponse xmlElement, xmlResponseElem, True
    
    'Set Submission Status
    Set xmlElement = CreateSubmissionStatus(vxmlRequest, strApplNumber, strAFFNumber, "AIPImpComplete")
    m_objErrAssist.CheckResponse xmlElement, xmlResponseElem, True
    
    'MC 20/06/2000. Unlock Application and Customers associated with it
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlElement = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlElement.Text = strApplNumber
    strResponse = objApplManagerBO.UnlockApplicationAndCustomers(xmlNewRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    
    Set IImportTxBO_CreateApplication = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlCustNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElement = Nothing
    Set objCustomerBO = Nothing
    Set objApplManagerBO = Nothing
    
    Exit Function
CreateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportTxBO_CreateApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function CreateApplicationStage( _
                    ByVal vxmlRequest As IXMLDOMNode, _
                    ByVal strApplNumber As String, ByVal strAFFNumber As String) As IXMLDOMNode

On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "CreateApplicationStage"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode
    
    Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objApplicationBO As ApplicationBO
    
    If m_objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
    Else
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONSTAGE"))
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlElement.Text = strApplNumber
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlElement.Text = strAFFNumber
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("STAGENUMBER"))
    xmlElement.Text = "35"
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("STAGENAME"))
    xmlElement.Text = "Agreement in Principle Complete"
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("DATETIME"))
    xmlElement.Text = Format(Now, "DD/MM/YYYY HH:MM:SS")
    
    strResponse = objApplicationBO.CreateApplicationStage(xmlNewRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set CreateApplicationStage = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set objApplicationBO = Nothing
    
    Exit Function
CreateVbErr:
        If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set CreateApplicationStage = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function CreateSubmissionStatus( _
                    ByVal vxmlRequest As IXMLDOMNode, _
                    ByVal strApplNumber As String, ByVal strAFFNumber As String, _
                    ByVal strComboValidation As String) As IXMLDOMNode

On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "CreateSubmissionStatus"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode
    
    Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objSubmtBO As SubmitBO
    Dim objComboDO As ComboDO
    
    If m_objContext Is Nothing Then
        Set objSubmtBO = New SubmitBO
        Set objComboDO = New ComboDO
    Else
        Set objSubmtBO = m_objContext.CreateInstance(gstrSubmission & ".SubmitBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONSUBMISSIONSTATUS"))
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlElement.Text = strApplNumber
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlElement.Text = strAFFNumber
    
    ' Get the Combo ValueID for the validation passed
    Dim strValueId As String
    strValueId = objComboDO.GetFirstComboValueId("SubmissionStatus", strComboValidation)

    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("SUBMISSIONSTATUS"))
    xmlElement.Text = strValueId
       
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("SUBMISSIONDATETIME"))
    xmlElement.Text = Format(Now, "DD/MM/YYYY HH:MM:SS")
    
    strResponse = objSubmtBO.CreateSubmissionStatus(xmlNewRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set CreateSubmissionStatus = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set objSubmtBO = Nothing
    Set objComboDO = Nothing
    
    Exit Function
CreateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set CreateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function UpdateSubmissionStatus( _
                    ByVal vxmlRequest As IXMLDOMNode, _
                    ByVal strApplNumber As String, ByVal strAFFNumber As String, _
                    ByVal strComboValidation As String) As IXMLDOMNode

On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "UpdateSubmissionStatus"
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode
    
    Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objSubmtBO As SubmitBO
    Dim objComboDO As ComboDO
    
    If m_objContext Is Nothing Then
        Set objSubmtBO = New SubmitBO
        Set objComboDO = New ComboDO
    Else
        Set objSubmtBO = m_objContext.CreateInstance(gstrSubmission & ".SubmitBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONSUBMISSIONSTATUS"))
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlElement.Text = strApplNumber
    
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlElement.Text = strAFFNumber
    
    ' Get the Combo ValueID for the validation passed
    Dim strValueId As String
    strValueId = objComboDO.GetFirstComboValueId("SubmissionStatus", strComboValidation)

    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("SUBMISSIONSTATUS"))
    xmlElement.Text = strValueId
       
    Set xmlElement = xmlTableElem.appendChild(xmlOut.createElement("SUBMISSIONDATETIME"))
    xmlElement.Text = Format(Now, "DD/MM/YYYY HH:MM:SS")
    
    strResponse = objSubmtBO.UpdateSubmissionStatus(xmlNewRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    Set UpdateSubmissionStatus = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set objSubmtBO = Nothing
    Set objComboDO = Nothing
    
    Exit Function
UpdateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set UpdateSubmissionStatus = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function

Private Function IImportTxBO_ImportApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ImportVbErr

    Const strFunctionName As String = "IImportTxBO_ImportApplication"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlRequestNode2 As IXMLDOMNode
    Dim xmlCreateElem As IXMLDOMNode
    Dim xmlCustRoleTypeElem As IXMLDOMNode
    Dim xmlCustomerOrderElem As IXMLDOMElement
    
    Dim xmlCustNodeList As IXMLDOMNodeList
    Dim xmlApplicationNode As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomerKeys As IXMLDOMElement
    
    Dim objXmlAssist As XMLAssist
    Dim objICustomerBO As ICustomerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                                
    Set xmlCustNodeList = vxmlRequest.selectNodes(".//CUSTOMER")
    If xmlCustNodeList.length = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                            oeMissingPrimaryTag, " CUSTOMER tag not found"
    End If
   
    Set xmlApplicationNode = vxmlRequest.selectSingleNode(".//APPLICATION")
    If xmlApplicationNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                            oeMissingPrimaryTag, " APPLICATION tag not found"
    End If
    
    ' Create the requset node to attach the CUSTOMER NODE to
    Set objXmlAssist = New XMLAssist
    Set xmlRequestNode = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlRequestNode2 = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlCreateElem = xmlOut.createElement("CREATE")
    xmlRequestNode.appendChild xmlCreateElem
    
    If m_objContext Is Nothing Then
        Set objICustomerBO = New CustomerBO
    Else
        Set objICustomerBO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        End If
    
    Set xmlCustomerKeys = xmlOut.createElement("CUSTOMERLIST")
    
    ' For every customer, create full customer details
    ' and, append the customer keys and CustomerRoleType to 'CustomerKeys' node
    Dim strCustomerRoleType As String
    
    For Each xmlNode In xmlCustNodeList
        xmlCreateElem.appendChild xmlNode
        Set xmlTempResponseNode = objICustomerBO.CreateFullCustomerDetails(xmlRequestNode)
        m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        ' Add 'CustomerRoleType' and CustomerOrder to the Customer Node in Response
        Set xmlCustRoleTypeElem = xmlOut.createElement("CUSTOMERROLETYPE")
'        strCustomerRoleType = objXmlAssist.GetTagValue(xmlNode, "CUSTOMERROLETYPE", True)
'        If strCustomerRoleType = "" Then strCustomerRoleType = 1
'        xmlCustRoleTypeElem.Text = strCustomerRoleType
        xmlCustRoleTypeElem.Text = xmlNode.selectSingleNode(".//CUSTOMERROLETYPE").Text
        xmlTempResponseNode.firstChild.appendChild xmlCustRoleTypeElem
        
        Set xmlCustomerOrderElem = xmlOut.createElement("CUSTOMERORDER")
        xmlCustomerOrderElem.Text = xmlNode.selectSingleNode(".//CUSTOMERORDER").Text
        xmlTempResponseNode.firstChild.appendChild xmlCustomerOrderElem
        
        xmlCustomerKeys.appendChild xmlTempResponseNode.selectSingleNode(".//CUSTOMER")
        
        Set xmlRequestNode = xmlRequestNode2
        Set xmlCreateElem = xmlOut.createElement("CREATE")
        xmlRequestNode.appendChild xmlCreateElem
    Next xmlNode
    
    Call CreateApplAndAddCustomers(vxmlRequest, xmlCustomerKeys)
    
    Set IImportTxBO_ImportApplication = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ImportExit:

    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestNode2 = Nothing
    Set xmlCreateElem = Nothing
    
    Set xmlCustNodeList = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlNode = Nothing
    Set xmlCustomerKeys = Nothing
    
    Set objXmlAssist = Nothing
    Set objICustomerBO = Nothing
    
    Exit Function

ImportVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportTxBO_ImportApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ImportExit

End Function

Private Function CreateApplAndAddCustomers(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal xmlCustomerKeys As IXMLDOMElement)
'----------------------------------------------------------------------------
'Description : Create the application and add all the customers (passes in)
'              to this application.
'Pass        : vxmlRequest - Request node to generate the application
'              xmlCustomerKeys - Customers created previously
'----------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "CreateApplAndAddCustomers"
    
    Dim objXmlAssist As XMLAssist
    Dim objIApplMgrBO As IApplicationManagerBO
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlApplElem As IXMLDOMNode
    Dim xmlApplNumberElem As IXMLDOMElement
    Dim xmlAFFNUmberElem As IXMLDOMElement
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    If m_objContext Is Nothing Then
        Set objIApplMgrBO = New ApplicationManagerBO
    Else
        Set objIApplMgrBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    End If

'   Create Application-Full-Details and extract the keys
    Set xmlTempResponseNode = objIApplMgrBO.CreateFullApplicationDetails(vxmlRequest)
    m_objErrAssist.ReRaiseResponseError xmlTempResponseNode
    
    Dim strApplNumber As String, strAFFNumber As String
        
    strApplNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    ' Create the Request for adding customer to application
    Set objXmlAssist = New XMLAssist
    Set xmlNewRequest = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempRequest = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlNodeList = xmlCustomerKeys.selectNodes(".//CUSTOMER")
    Set xmlApplNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    xmlApplNumberElem.Text = strApplNumber
    Set xmlAFFNUmberElem = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlAFFNUmberElem.Text = strAFFNumber
    
    For Each xmlNode In xmlNodeList
        Set xmlApplElem = xmlOut.createElement("APPLICATION")
        xmlNewRequest.appendChild xmlApplElem
        
        xmlApplElem.appendChild xmlNode.selectSingleNode("CUSTOMERNUMBER")
        xmlApplElem.appendChild xmlNode.selectSingleNode("CUSTOMERVERSIONNUMBER")
        xmlApplElem.appendChild xmlNode.selectSingleNode("CUSTOMERROLETYPE")
        xmlApplElem.appendChild xmlNode.selectSingleNode("CUSTOMERORDER")
    
        xmlApplElem.appendChild xmlApplNumberElem
        xmlApplElem.appendChild xmlAFFNUmberElem
        
        Set xmlTempResponseNode = objIApplMgrBO.AddCustomerToApplication(xmlNewRequest)
        m_objErrAssist.ReRaiseResponseError xmlTempResponseNode
        
        Set xmlNewRequest = xmlTempRequest
    Next xmlNode
    
CreateExit:
    Set objXmlAssist = Nothing
    Set objIApplMgrBO = Nothing
    Set xmlOut = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlApplElem = Nothing
    Set xmlApplNumberElem = Nothing
    Set xmlAFFNUmberElem = Nothing
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
            
    Exit Function
CreateVbErr:
    Set objXmlAssist = Nothing
    Set objIApplMgrBO = Nothing
    Set xmlOut = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlApplElem = Nothing
    Set xmlApplNumberElem = Nothing
    Set xmlAFFNUmberElem = Nothing
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IImportTxBO_UpdateApplication( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IImportTxBO_UpdateApplication"
    
    Dim xmlOut As New FreeThreadedDOMDocument40, xmlDoc As FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode, xmlTableElem As IXMLDOMNode
    
    Dim strResponse As String, strCondition As String
    Dim strApplNumber As String, strAFFNumber As String
    Dim lngErrNumber As Long
    Dim blnRecordExists As Boolean
    
    Dim objIAdoAssist As IADOAssist
    Dim objNewPropertyBO As NewPropertyBO
    Dim objApplicationBO As ApplicationBO
    Dim objApplThirdPartyBO As ApplicationThirdPartyBO
    Dim objCustomerBO As CustomerBO, objCFBO As CustomerFinancialBO
    Dim objCEBO As CustomerEmploymentBO
    
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objNewPropertyBO = New NewPropertyBO
        Set objIAdoAssist = New ADOAssist
        Set objApplicationBO = New ApplicationBO
        Set objApplThirdPartyBO = New ApplicationThirdPartyBO
        Set objCustomerBO = New CustomerBO
        Set objCFBO = New CustomerFinancialBO
        Set objCEBO = New CustomerEmploymentBO
    Else
        Set objNewPropertyBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
        Set objIAdoAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objApplicationBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        Set objApplThirdPartyBO = _
            m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationThirdPartyBO")
        Set objCustomerBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
        Set objCFBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        Set objCEBO = _
            m_objContext.CreateInstance(gstrCUSTOMER_EMPLOYMENT & ".CustomerEmploymentBO")
    End If
    
    '-------------------------------------------------------------------------------------------------
    ' Update NewPropertyDetails. If error 'No Records affected' is returned, create NewPropertyDetails
    '--------------------------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//NEWPROPERTY") Is Nothing Then
        strResponse = objNewPropertyBO.UpdateNewPropertyDetails(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        lngErrNumber = m_objErrAssist.CheckResponse(xmlDoc.documentElement, xmlResponseElem, False)
    
        If lngErrNumber <> 0 Then
            lngErrNumber = m_objErrAssist.GetOmigaErrorNumber(lngErrNumber)
        
            If lngErrNumber = oeNoRowsAffected Then
                strResponse = objNewPropertyBO.CreateNewPropertyDetails(vxmlRequest.xml)
                Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
                m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
            Else
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, lngErrNumber
            End If
        End If
    End If
    '-----------------------------------------------------------------------------------
    ' Check whether corresponding record exist in 'NewPropertyAddress', if so update it
    ' else create one
    '-----------------------------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//NEWPROPERTYADDRESS") Is Nothing Then
        strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
        strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
        strCondition = " APPLICATIONNUMBER = '" & strApplNumber & "' AND " _
                       & " APPLICATIONFACTFINDNUMBER = " & strAFFNumber
                   
        blnRecordExists = objIAdoAssist.CheckRecordExists("NEWPROPERTYADDRESS", strCondition)
    
        If blnRecordExists Then
            strResponse = objNewPropertyBO.UpdateNewPropertyAddress(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        Else
            strResponse = objNewPropertyBO.CreateNewPropertyAddress(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        End If
    
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        
    End If
    
'** FIX ME - Commit to Buy - Waiting for spec completion
    
    '---------------------------------------------------
    'Create Other Insurance Company
    '----------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//OTHERINSURANCECOMPANY") Is Nothing Then
        strResponse = objApplicationBO.CreateOtherInsuranceCompany(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
    
    '-------------------------------------------------------------
    'Save Insurance Disclosures
    '-------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//DECLARATION") Is Nothing Then
        strResponse = objApplicationBO.SaveDeclaration(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
    
    '-------------------------------------------------------------
    ' Create Estate Agents records
    '-------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//APPLICATIONESTATEAGENT") Is Nothing Then
        strResponse = objApplThirdPartyBO.CreateEstateAgent(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
    
    '-------------------------------------------------------------
    ' Create Legal Representatives records
    '-------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//APPLICATIONLEGALREP") Is Nothing Then
        strResponse = objApplThirdPartyBO.CreateLegalRep(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If

    '---------------------------------------------------------------
    ' Update Application Fact Find - Application Notes
    '---------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//APPLICATIONFACTFIND") Is Nothing Then
        strResponse = objApplicationBO.UpdateApplicationFactFind(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If

    '-------------------------------------------------------------
    ' Create BankBuildingSociety Records
    '-------------------------------------------------------------
    If Not vxmlRequest.selectSingleNode(".//APPLICATIONBANKBUILDINGSOC") Is Nothing Then
        strResponse = objApplThirdPartyBO.CreateBankBuildingSociety(vxmlRequest.xml)
        Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If

    '------------------------------------------------------------------------
    ' Save data related to customers
    '------------------------------------------------------------------------
    Set xmlNodeList = vxmlRequest.selectNodes(".//CUSTOMER")
    For Each xmlNode In xmlNodeList
        '----------------------------------------------------------------------
        ' Update Customer Version
        '----------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//CUSTOMERVERSION") Is Nothing Then
            strResponse = objCustomerBO.UpdateCustomerVersion(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

        '---------------------------------------------------------------------
        ' Save Dependants for customer
        '---------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//DEPENDANT") Is Nothing Then
            strResponse = objApplicationBO.SaveDependantForCustomer(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

        '----------------------------------------------------------------------
        ' Create Bank Card details
        '----------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//BANKCREDITCARD") Is Nothing Then
            strResponse = objCFBO.CreateBankCard(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

        '----------------------------------------------------------------------
        ' Save Employed details
        '----------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//EMPLOYMENT") Is Nothing Then
            strResponse = objCEBO.SaveEmploymentDetails(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

        '----------------------------------------------------------------------
        ' Create Tenancy details
        '----------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//TENANCY") Is Nothing Then
            strResponse = objCustomerBO.CreateTenancy(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

        '----------------------------------------------------------------------
        ' Save Accountant details
        '----------------------------------------------------------------------
        If Not xmlNode.selectSingleNode(".//ACCOUNTANT") Is Nothing Then
            strResponse = objCEBO.SaveAccountantDetails(vxmlRequest.xml)
            Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
            m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
        End If

    Next xmlNode
    
'** FIX ME - Loan Details - Waiting for completion of Spec
    
    ' Create Application Stage with stage number of 55
    Set xmlNewRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlNewRequest.appendChild(xmlOut.createElement("APPLICATIONSTAGE"))
    
    Set xmlNode = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONNUMBER"))
    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    
    Set xmlNode = xmlTableElem.appendChild(xmlOut.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlNode.Text = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    Set xmlNode = xmlTableElem.appendChild(xmlOut.createElement("STAGENUMBER"))
    xmlNode.Text = "55"
    
    Set xmlNode = xmlTableElem.appendChild(xmlOut.createElement("STAGENAME"))
    xmlNode.Text = "Mortgage Application Complete"
    
    Set xmlNode = xmlTableElem.appendChild(xmlOut.createElement("DATETIME"))
    xmlNode.Text = Format(Now, "DD/MM/YYYY HH:MM:SS")
    
    strResponse = objApplicationBO.CreateApplicationStage(xmlNewRequest.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    'Set the submission status of this application to "MAImpComplete"
    Set xmlNode = UpdateSubmissionStatus(vxmlRequest, strApplNumber, strAFFNumber, "MAImpComplete")
    m_objErrAssist.CheckResponse xmlNode, xmlResponseElem, True
    
    Set IImportTxBO_UpdateApplication = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlTableElem = Nothing
    Set xmlNewRequest = Nothing
    
    Set xmlResponseElem = Nothing
    Set objNewPropertyBO = Nothing
    Set objApplicationBO = Nothing
    Set objIAdoAssist = Nothing
    Set objApplThirdPartyBO = Nothing
    Set objCustomerBO = Nothing
    Set objCFBO = Nothing
    Set objCEBO = Nothing
    
    Exit Function
UpdateVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportTxBO_UpdateApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub

