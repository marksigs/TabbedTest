VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LifeSubQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LifeSubQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Dependencies:  List any other dependent components
'               e.g. LifeSubQuoteTxBO, LifeSubQuoteDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     26/11/1999  Created
'PSC    20/11/1999  Amend to return the created sub quote number
'PSC    01/02/2000  Added SaveRecostedBenefits()
'SR     25/04/2000  New mehtod - GetData
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
Implements ILifeSubQuoteBO

Private Const cstrTABLE_NAME = "LIFESUBQUOTE"

Public Function SaveRecostedBenefits(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Deletes the life benefits for a subquote and creates the
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveRecostedBenefitsVbErr

    Const strFunctionName As String = "SaveRecostedBenefits"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = ILifeSubQuoteBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = ILifeSubQuoteBO_SaveRecostedBenefits(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveRecostedBenefits = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveRecostedBenefitsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveRecostedBenefitsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveRecostedBenefits = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveRecostedBenefitsExit
    
End Function


Public Function UpdateLifeBenefit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeBenefitVbErr

    Const strFunctionName As String = "UpdateLifeBenefit"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

    Dim objLifeSubQuoteTxBO As LifeSubQuoteTxBO

    If objContext Is Nothing Then
        Set objLifeSubQuoteTxBO = New LifeSubQuoteTxBO
    Else
        Set objLifeSubQuoteTxBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    objLifeSubQuoteTxBO.UpdateLifeBenefit vstrXMLRequest

    UpdateLifeBenefit = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateLifeBenefitExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objLifeSubQuoteTxBO = Nothing

    Exit Function

UpdateLifeBenefitVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateLifeBenefit = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume UpdateLifeBenefitExit

End Function

Public Function DeleteLifeBenefit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the
'               table 'LIFEBENEFIT'
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeBenefitVbErr

    Const strFunctionName As String = "DeleteLifeBenefit"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objLifeSubQuoteTxBO As LifeSubQuoteTxBO

    If objContext Is Nothing Then
        Set objLifeSubQuoteTxBO = New LifeSubQuoteTxBO
    Else
        Set objLifeSubQuoteTxBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '   call Business Transaction Object DeleteLifeBenefit function
   objLifeSubQuoteTxBO.DeleteLifeBenefit (vstrXMLRequest)

    DeleteLifeBenefit = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteLifeBenefitExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objLifeSubQuoteTxBO = Nothing
    Set objXmlAssist = Nothing

    Exit Function

DeleteLifeBenefitVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteLifeBenefit = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    '   go to clean-up section
    Resume DeleteLifeBenefitExit

End Function

Public Function CreateFirstSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object - Table (LifeSubQuote)
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr
    
    Const strFunctionName As String = "CreateFirstSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim xmlTxBODoc As FreeThreadedDOMDocument40
    
    Dim objLifeSubQuoteTxBO As LifeSubQuoteTxBO
    
    Dim strXmlOut As String

    If objContext Is Nothing Then
        Set objLifeSubQuoteTxBO = New LifeSubQuoteTxBO
    Else
        Set objLifeSubQuoteTxBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   call Business Transaction Object Create function
    strXmlOut = objLifeSubQuoteTxBO.Create(vstrXMLRequest, "FIRST")
    Set xmlTxBODoc = objXmlAssist.load(strXmlOut, TypeName(Me), strFunctionName)
    xmlResponseElem.appendChild xmlTxBODoc.documentElement
        
    CreateFirstSubQuote = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objLifeSubQuoteTxBO = Nothing
    Set xmlTxBODoc = Nothing

    Exit Function

CreateFirstSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateFirstSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstSubQuoteExit

End Function

Public Function CreateNewSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object - Table (LifeSubQuote)
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr
    
    Const strFunctionName As String = "CreateNewSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlElement As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist

    Dim objLifeSubQuoteTxBO As LifeSubQuoteTxBO
    Dim strXmlOut As String
    
    If objContext Is Nothing Then
        Set objLifeSubQuoteTxBO = New LifeSubQuoteTxBO
    Else
        Set objLifeSubQuoteTxBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteTxBO")
    End If

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    '   call Business Transaction Object Create function
    strXmlOut = objLifeSubQuoteTxBO.Create(vstrXMLRequest, "NEW")
    Set xmlDoc = objXmlAssist.load(strXmlOut, TypeName(Me), strFunctionName)
           
    xmlResponseElem.appendChild xmlDoc.documentElement
    
    CreateNewSubQuote = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objLifeSubQuoteTxBO = Nothing

    Exit Function

CreateNewSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewSubQuoteExit

End Function

Private Function ILifeSubQuoteBO_GetData( _
            ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "ILifeSubQuoteBO_GetData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objILifeSubQuoteDO As ILifeSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objILifeSubQuoteDO = New LifeSubQuoteDO
    Else
        Set objILifeSubQuoteDO = objContext.CreateInstance(App.Title & ".LifeSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = "LIFESUBQUOTE" Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName("LIFESUBQUOTE").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "LIFESUBQUOTE" & " tag not found"
    End If
        
    Set xmlDataNode = objILifeSubQuoteDO.GetData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ILifeSubQuoteBO_GetData = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetDataExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objILifeSubQuoteDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ILifeSubQuoteBO_GetData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function ILifeSubQuoteBO_SaveRecostedBenefits(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveRecostedBenefitsVbErr
    
    Const strFunctionName As String = "ILifeSubQuoteBO_SaveRecostedBenefits"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objILifeSubQuoteTxBO As ILifeSubQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objILifeSubQuoteTxBO = New LifeSubQuoteTxBO
    Else
        Set objILifeSubQuoteTxBO = objContext.CreateInstance(App.Title & ".LifeSubQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objILifeSubQuoteTxBO.SaveRecostedBenefits(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set ILifeSubQuoteBO_SaveRecostedBenefits = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveRecostedBenefitsExit:

    Set objErrAssist = Nothing
    Set objILifeSubQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveRecostedBenefitsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ILifeSubQuoteBO_SaveRecostedBenefits = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveRecostedBenefitsExit

End Function

Private Function ILifeSubQuoteBO_Validate(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "ILifeSubQuoteBO_Validate"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeElem As IXMLDOMElement
'    Dim xmlDataNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
'    Dim xmlDataElement As IXMLDOMElement
'    Dim xmlDataList As IXMLDOMNodeList
'    Dim objILifeSubQuoteDO As ILifeSubQuoteDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

'    If objContext Is Nothing Then
'        Set objIOmiga4DO = New Omiga4DO
'    Else
'        Set objIOmiga4DO = objContext.CreateInstance(App.Title & ".Omiga4DO")
'    End If
'
'    ' Get node to add derived data to and apply validation to
'    If vxmlRequest.nodeName = "UPDATE" Or vxmlRequest.nodeName = "CREATE" Then
'        Set xmlRequestTypeElem = vxmlRequest
'    ElseIf vxmlRequest.getElementsByTagName("UPDATE").length > 0 Then
'        Set xmlRequestTypeElem = vxmlRequest.getElementsByTagName("UPDATE").Item(0)
'    ElseIf vxmlRequest.getElementsByTagName("CREATE").length > 0 Then
'        Set xmlRequestTypeElem = vxmlRequest.getElementsByTagName("CREATE").Item(0)
'    End If
'
'    If xmlRequestTypeElem Is Nothing Then
'        objErrAssist.ThrowError TypeName(Me), _
'                                strFunctionName, _
'                                oeMissingPrimaryTag, _
'                                "CREATE or UPDATE tag missing"
'    End If
'
'    Set xmlDataList = xmlRequestTypeElem.getElementsByTagName(cstrROOT_NODE_NAME)
'
'    For Each xmlDataElement In xmlDataList
'            objIOmiga4DO.AddDerivedData xmlDataElement
'
'            ' TODO ************************************************************************
'            '   add hard coded validation here, & raise error if it fails
'            ' *****************************************************************************
'    Next
'
'    '-----------------------------------------------------------------------------
'    ' hard-coded validation OK
'    ' return the response
'    '-----------------------------------------------------------------------------
    
    Set ILifeSubQuoteBO_Validate = xmlResponseElem.cloneNode(True)
    
ValidateExit:

    Set xmlRequestTypeElem = Nothing
'    Set xmlDataNode = Nothing
'    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
'    Set objILifeSubQuoteDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
'    Set xmlDataElement = Nothing
'    Set xmlDataList = Nothing
    
    Exit Function

ValidateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ILifeSubQuoteBO_Validate = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with the table
'               'Application Stage'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetData"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = ILifeSubQuoteBO_GetData(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetData = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDataExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetDataVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

