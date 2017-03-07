VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuickQuoteApplicantDetailsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuoteApplicantDetailsBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (QuickQuoteApplicantDetailsTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:
'               QuickQuoteApplicantDetailsTxBO,
'               QuickQuoteApplicantDetailsDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     01/08/99    Created
'JLD    25/10/99    Issue 32 read the outstanding loan amount and the total monthly repayment
'                   from the string passed in, not the database
'DJP    29/02/00    Get ApplicationBO from omCR not omAPP
'APS    22/05/00    SYS0779 Removed Customer Registration component
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

Private Const cintMAX_QQ_APPLICANTS = 2

'=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "QUICKQUOTEAPPLICANTDETAILS"

Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       Create          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "Create"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objErrAssist As New ErrAssist
    Dim objXmlNode As IXMLDOMNode

   Dim objQuickQuoteApplicantDetailsTxBO As QuickQuoteApplicantDetailsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsTxBO = New QuickQuoteApplicantDetailsTxBO
    Else
        Set objQuickQuoteApplicantDetailsTxBO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    
    '---------------------------------------------------------------------------
    ' perform hard-coded validation & call scripting interface
    '---------------------------------------------------------------------------
    Validate vstrXMLRequest
    
    '   call Business Transaction Object Create function
    objQuickQuoteApplicantDetailsTxBO.Create (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Create = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsTxBO = Nothing

    Exit Function

CreateVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Create = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume CreateExit

End Function

Public Function Update(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data update
' return:       Update          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "Update"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        
' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsTxBO As QuickQuoteApplicantDetailsTxBO
    
    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsTxBO = New QuickQuoteApplicantDetailsTxBO
    Else
        Set objQuickQuoteApplicantDetailsTxBO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    
    If Not objXmlResponseElem Is Nothing Then
        Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    End If
    
    ' Call scripting interface
    Validate vstrXMLRequest
    
    objQuickQuoteApplicantDetailsTxBO.Update vstrXMLRequest
    
    '   if we get here, everything has completed OK
    If Not objXmlResponseElem Is Nothing Then
        objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    End If
    
    Update = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objQuickQuoteApplicantDetailsTxBO = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As FreeThreadedDOMDocument40
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Update = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume UpdateExit
    
End Function


Public Function Delete(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       Delete          xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    
    Const strFunctionName As String = "Delete"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    
    Dim objErrAssist As New ErrAssist

    Dim objQuickQuoteApplicantDetailsTxBO As QuickQuoteApplicantDetailsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsTxBO = New QuickQuoteApplicantDetailsTxBO
    Else
        Set objQuickQuoteApplicantDetailsTxBO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
           
    '   call Business Transaction Object Delete function
   objQuickQuoteApplicantDetailsTxBO.Delete (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Delete = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsTxBO = Nothing

    Exit Function

DeleteVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Delete = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume DeleteExit

End Function

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       GetData         xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "GetData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    

' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
                
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    '   call Data Object GetData function
    Dim strXMLData As String
    
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objQuickQuoteApplicantDetailsDO.GetData(objXmlRequestNode.xml)
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If

    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    GetData = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDataExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing

    Exit Function

GetDataVbErr:
    
    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetDataExit

End Function

Private Sub Validate(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "Validate"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlDerived As New FreeThreadedDOMDocument40
    objXmlDerived.validateOnParse = False
    objXmlDerived.setProperty "NewParser", True
    Dim objXmlBeforeData As IXMLDOMElement
    Dim objXmlAfterData As IXMLDOMElement
    Dim objXmlAfterNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist

    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
    
' TODO ************************************************************************
'   add processing here
'   e.g.
'   examine input xml request & extract before image node (if present)
'   & after image node
' *****************************************************************************
    
    If objXmlIn.getElementsByTagName("UPDATE").length > 0 Then
        Dim int0 As Integer
        For int0 = 0 To objXmlIn.getElementsByTagName("UPDATE").length - 1
            Select Case objXmlIn.getElementsByTagName("UPDATE").Item(int0).Attributes.getNamedItem("TYPE").Text
                Case "BEFORE"
                    Set objXmlBeforeData = objXmlIn.getElementsByTagName("UPDATE").Item(int0).firstChild
                Case "AFTER"
                    Set objXmlAfterNode = objXmlIn.getElementsByTagName("UPDATE").Item(int0)
                    Set objXmlAfterData = objXmlAfterNode.firstChild
            End Select
        Next
    Else
        If objXmlIn.getElementsByTagName("CREATE").length > 0 Then
            Set objXmlAfterNode = objXmlIn.getElementsByTagName("CREATE").Item(int0)
            Set objXmlAfterData = objXmlAfterNode.firstChild
        End If
    End If
    
'   replace 'after' image with xml plus any derived values

    If Not objXmlAfterData Is Nothing Then
        Dim strXMLAfterData As String
        strXMLAfterData = objQuickQuoteApplicantDetailsDO.AddDerivedData(objXmlAfterData.xml)
        Set objXmlDerived = objXmlAssist.load(strXMLAfterData, TypeName(Me), strFunctionName)
        objXmlAfterNode.replaceChild objXmlDerived.documentElement, objXmlAfterData
    End If
    
    '   -----------------------------------------------------------------------------
    '   hard-coded validation OK
    '   call scripting interface
    '   -----------------------------------------------------------------------------

    'Dim objScriptInterface As New omBase.ScriptInterface
    Dim objScriptInterface As ScriptInterface
    Set objScriptInterface = New ScriptInterface
    
    objScriptInterface.RunScript App.Title, TypeName(Me), _
                                 strFunctionName, vstrXMLRequest
    
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing

Exit Sub

ValidateVbErr:

    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlBeforeData = Nothing
    Set objXmlAfterData = Nothing
    Set objXmlAfterNode = Nothing
    Set objXmlDerived = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
    
Public Function SaveQuickQuoteData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Validate and save a set of Quick Quote data, including some ApplicationFactFind data
'   (Purchase Price, O/S Loan Amount and Total Monthly Repayments)
' pass:
'   vstrXMLRequest
'       xml Request. Format:
'       <REQUEST>
'            <QUICKQUOTEAPPLICANTDETAILSLIST>
'                <QUICKQUOTEAPPLICANTDETAILS>
'                    <QUICKQUOTEOUTGOINGSLIST>
'                        <QUICKQUOTEOUTGOINGS>
'                        </QUICKQUOTEOUTGOINGS>
'                    </QUICKQUOTEOUTGOINGSLIST>
'                    <QUICKQUOTEPERSONALDEBTSLIST>
'                        <QUICKQUOTEPERSONALDEBTS>
'                        </QUICKQUOTEPERSONALDEBTS>
'                    </QUICKQUOTEPERSONALDEBTSLIST>
'                </QUICKQUOTEAPPLICANTDETAILS>
'            </QUICKQUOTEAPPLICANTDETAILSLIST>
'            <UPDATE TYPE="AFTER">
'               <APPLICATIONFACTFIND>
'               </APPLICATIONFACTFIND>
'            </UPDATE>
'       </REQUEST>
' return:
'   SaveQuickQuoteData  xml Response data stream containing results of operation
'                       either: TYPE="SUCCESS" and xml representation of data
'                       or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SaveQuickQuoteDataVbErr
    
    Const strFunctionName As String = "SaveQuickQuoteData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objQQAppDetailsTxBO As QuickQuoteApplicantDetailsTxBO
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim objXmlResponseNode As IXMLDOMNode
    Dim objXmlTempNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim strResponse As String
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlResponseNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If objContext Is Nothing Then
        Set objQQAppDetailsTxBO = New QuickQuoteApplicantDetailsTxBO
    Else
        Set objQQAppDetailsTxBO = objContext.CreateInstance( _
            App.Title & ".QuickQuoteApplicantDetailsTxBO")
    End If
    
    ' check how many applicant details have been supplied
    Dim intNumApplicants As Integer
    intNumApplicants = _
        objXmlIn.getElementsByTagName("QUICKQUOTEAPPLICANTDETAILS").length
    If intNumApplicants > cintMAX_QQ_APPLICANTS Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter
    End If

    '------------------------------------------------------------------------------------------
    ' Check if Quick Quote is allowed for this application fact find
    '------------------------------------------------------------------------------------------
    
    CheckIfQuickQuoteAllowed vstrXMLRequest
    
    '------------------------------------------------------------------------------------------
    ' Save the data
    '------------------------------------------------------------------------------------------
    
    objQQAppDetailsTxBO.SaveQuickQuoteData vstrXMLRequest
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    SaveQuickQuoteData = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveQuickQuoteDataExit:
    
    Set objErrAssist = Nothing
    Set objXmlResponseNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objQQAppDetailsTxBO = Nothing

    Exit Function

SaveQuickQuoteDataVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlErrMsgDoc As New FreeThreadedDOMDocument40
        objXmlErrMsgDoc.validateOnParse = False
        objXmlErrMsgDoc.setProperty "NewParser", True
        objXmlErrMsgDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlResponseNode = objXmlResponseElem.appendChild(objXmlErrMsgDoc.documentElement)
        Set objXmlErrMsgDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    SaveQuickQuoteData = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume SaveQuickQuoteDataExit

End Function

Public Function FindList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       FindList         xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "FindList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    
    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    Dim strXMLData As String
    
    '   call Data Object FindList function
    If Not objXmlRequestNode Is Nothing Then
        strXMLData = objQuickQuoteApplicantDetailsDO.FindList(objXmlRequestNode.xml)
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Else
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    FindList = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindListExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing

    Exit Function

FindListVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindListExit

End Function

Public Function DeleteAll(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       DeleteAll       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "DeleteAll"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

    
' TODO ************************************************************************
'   replace QuickQuoteApplicantDetailsTxBO reference with actual TxBO required
' *****************************************************************************
    Dim objQuickQuoteApplicantDetailsTxBO As QuickQuoteApplicantDetailsTxBO

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsTxBO = New QuickQuoteApplicantDetailsTxBO
    Else
        Set objQuickQuoteApplicantDetailsTxBO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsTxBO")
    End If

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
        
    '   call Business Transaction Object DeleteAll function
    objQuickQuoteApplicantDetailsTxBO.DeleteAll (vstrXMLRequest)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    DeleteAll = objXmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteAllExit:

    '   clean-up section, release object references
    Set objXmlOut = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsTxBO = Nothing

    Exit Function

DeleteAllVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteAll = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume DeleteAllExit

End Function

Public Sub CheckIfQuickQuoteAllowed(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Check if QuickQuote allowed for this application fact find:
'   Check if application is a further advance or a transfer of equity. If it is either,
'   check OUTSTANDINGLOANAMOUNT and TOTALMONTHLYREPAYMENT are not null for the
'   application fact find.
'   Called by SaveQuickQuoteData.
'   APS 24/05/00 - SYS0779 Altered the processing to call Cross component BO instead of a DO
' pass:
'   vstrXMLRequest
' return:   n/a
' Raise Errors:
'   omiga4err192
'------------------------------------------------------------------------------------------
On Error GoTo CheckIfQuickQuoteAllowedVbErr
    
    Const strFunctionName As String = "CheckIfQuickQuoteAllowed"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objApplicationBO As ApplicationBO
    Dim objXmlAssist As New XMLAssist
    
    If objContext Is Nothing Then
        Set objApplicationBO = New ApplicationBO
    Else
        ' note use of base components
        Set objApplicationBO = objContext.CreateInstance( _
            gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
    
    Dim strResponse, strTypeOfApplication, _
        strOUTSTANDINGLOANAMOUNT, strTOTALMONTHLYREPAYMENT As String
    
    'JLD  read the outstanding loan amount and the total monthly repayment
    '     from the string passed in, not the database
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'APS 22/05/00 SYS0779 Removed Customer Registration component
    Dim xmlDOMApp As New FreeThreadedDOMDocument40
    xmlDOMApp.validateOnParse = False
    xmlDOMApp.setProperty "NewParser", True
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlApplicationFFNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlAppFF As IXMLDOMNode

    Set xmlNode = objXmlAssist.GetRequestNodeEx(objXmlIn.documentElement)
    Set xmlRequestNode = xmlDOMApp.appendChild(xmlNode)

    Set xmlNode = xmlDOMApp.createElement("APPLICATION")
    Set xmlAppFF = xmlRequestNode.appendChild(xmlNode)

    Set xmlApplicationNode = objXmlAssist.GetNode(objXmlIn.documentElement, ".//APPLICATIONNUMBER")
    Set xmlApplicationFFNode = objXmlAssist.GetNode(objXmlIn.documentElement, ".//APPLICATIONFACTFINDNUMBER")

    xmlAppFF.appendChild xmlApplicationNode.cloneNode(True)
    xmlAppFF.appendChild xmlApplicationFFNode.cloneNode(True)
    
    Dim strXMLResponse As String
    strXMLResponse = objApplicationBO.GetApplicationData(xmlDOMApp.xml)
    objErrAssist.CheckXMLResponse strXMLResponse, True
    Set xmlDoc = objXmlAssist.load(strXMLResponse, TypeName(Me), strFunctionName)
    
    strTypeOfApplication = _
        objXmlAssist.GetTagValue(xmlDoc.documentElement, "TYPEOFAPPLICATION")
        
    Dim blnIsFurtherAdvance, blnIsTransferOfEquity As Boolean
    
    blnIsFurtherAdvance = objApplicationBO.IsFurtherAdvance(strTypeOfApplication)
    blnIsTransferOfEquity = objApplicationBO.IsTransferOfEquity(strTypeOfApplication)
    
    If (blnIsFurtherAdvance = True) Or _
        (blnIsTransferOfEquity = True) Then
    
        strOUTSTANDINGLOANAMOUNT = _
            objXmlAssist.GetTagValue(objXmlIn.documentElement, "OUTSTANDINGLOANAMOUNT")
            'objXmlAssist.GetTagValue(xmlDoc.documentElement, "OUTSTANDINGLOANAMOUNT")  JLD 25/10/99
        strTOTALMONTHLYREPAYMENT = _
            objXmlAssist.GetTagValue(objXmlIn.documentElement, "TOTALMONTHLYREPAYMENT")
            'objXmlAssist.GetTagValue(xmlDoc.documentElement, "TOTALMONTHLYREPAYMENT")  JLD 25/10/99
    
        If Len(strOUTSTANDINGLOANAMOUNT) = 0 Or _
            Len(strTOTALMONTHLYREPAYMENT) = 0 Then
            
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err192
            
        End If
    End If

    Set objErrAssist = Nothing
    Set xmlDoc = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objApplicationBO = Nothing
    
    Set xmlDOMApp = Nothing
    Set xmlNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlApplicationFFNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlAppFF = Nothing

    Exit Sub

CheckIfQuickQuoteAllowedVbErr:

    Set xmlDoc = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objApplicationBO = Nothing
    
    Set xmlDOMApp = Nothing
    Set xmlNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlApplicationFFNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlAppFF = Nothing
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Public Function GetQuickQuoteData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get details of any existing Quick Quote info held for single and joint applicants.
'   Also get ApplicationFactFind data.
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   instance of the persistant data to be retrieved
' return:
'   GetQuickQuoteData         xml Response data stream containing results of operation
'                             either: TYPE="SUCCESS" and xml representation of data
'                             or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetQuickQuoteDataVbErr
    
    Const strFunctionName As String = "GetQuickQuoteData"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As New FreeThreadedDOMDocument40
    objXmlIn.validateOnParse = False
    objXmlIn.setProperty "NewParser", True
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    Dim objApplicationBO As ApplicationBO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)

    ' APS 22/05/00 Removed Customer Registration component
    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
        Set objApplicationBO = New ApplicationBO
    Else
        Set objQuickQuoteApplicantDetailsDO = _
            objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
        Set objApplicationBO = _
            objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If
                
    Dim strTagName As String
    strTagName = "CUSTOMERROLELIST"
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(strTagName).Item(0)
    
    If objXmlRequestNode Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, strTagName & " tag not found"
    End If
    
    '------------------------------------------------------------------------------------------
    ' check if we have customer info
    '------------------------------------------------------------------------------------------
    Dim xmlDOMApp As New FreeThreadedDOMDocument40
    xmlDOMApp.validateOnParse = False
    xmlDOMApp.setProperty "NewParser", True
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlApplicationFFNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlAppFF As IXMLDOMNode

    Set xmlNode = objXmlAssist.GetRequestNodeEx(objXmlIn.documentElement)
    Set xmlRequestNode = xmlDOMApp.appendChild(xmlNode)

    Set xmlNode = xmlDOMApp.createElement("APPLICATION")
    Set xmlAppFF = xmlRequestNode.appendChild(xmlNode)

    Set xmlApplicationNode = objXmlAssist.GetNode(objXmlRequestNode, ".//APPLICATIONNUMBER")
    Set xmlApplicationFFNode = objXmlAssist.GetNode(objXmlRequestNode, ".//APPLICATIONFACTFINDNUMBER")

    xmlAppFF.appendChild xmlApplicationNode.cloneNode(True)
    xmlAppFF.appendChild xmlApplicationFFNode.cloneNode(True)
        
    Dim strCUSTOMERNUMBER As String
    Dim strResponse As String
    
    Set xmlElem = objXmlRequestNode
    strCUSTOMERNUMBER = objXmlAssist.GetTagValue(xmlElem, "CUSTOMERNUMBER")
    If Len(strCUSTOMERNUMBER) = 0 Then
        ' we need to find the main applicants and replace the section of the
        ' input parameters
        
        ' APS 22/05/00 Removed Customer Registration component
        strResponse = objApplicationBO.FindMainApplicants(xmlDOMApp.xml)
        Set objXmlData = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
        Dim xmlCustomerList As IXMLDOMNode
        Set xmlCustomerList = objXmlAssist.GetNode(objXmlData.documentElement, "MAINAPPLICANTLIST")
        
        Dim xmlParent As IXMLDOMNode
        Set xmlParent = objXmlRequestNode.parentNode
        If (Not xmlCustomerList Is Nothing) And (Not xmlParent Is Nothing) Then
            xmlParent.replaceChild xmlCustomerList, objXmlRequestNode
        End If
        Set objXmlRequestNode = Nothing
        strTagName = "MAINAPPLICANTLIST"
        Set objXmlRequestNode = objXmlIn.getElementsByTagName(strTagName).Item(0)
    End If
    
    '------------------------------------------------------------------------------------------
    ' call Data Object GetQuickQuoteData function
    '------------------------------------------------------------------------------------------
    
    strResponse = objQuickQuoteApplicantDetailsDO.GetQuickQuoteData(objXmlRequestNode.xml)
    Set objXmlData = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    Set objXmlData = Nothing
    '------------------------------------------------------------------------------------------
    ' get ApplicationFactFind data
    '------------------------------------------------------------------------------------------
    
    ' APS 22/05/00 Removed Customer Registration component
    strResponse = objApplicationBO.GetApplicationData(xmlDOMApp.xml)
    objErrAssist.CheckXMLResponse strResponse, True
    Set objXmlData = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    
    '   if we get here, everything has completed OK
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    GetQuickQuoteData = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetQuickQuoteDataExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing
    Set objApplicationBO = Nothing
    
    Set xmlDOMApp = Nothing
    Set xmlNode = Nothing
    Set xmlApplicationFFNode = Nothing
    Set xmlAppFF = Nothing
    Set xmlApplicationNode = Nothing
    
    Exit Function

GetQuickQuoteDataVbErr:
    
    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlErrDoc As New FreeThreadedDOMDocument40
        objXmlErrDoc.validateOnParse = False
        objXmlErrDoc.setProperty "NewParser", True
        objXmlErrDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlErrDoc.documentElement)
        Set objXmlErrDoc = Nothing
        Resume Next
    End If

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetQuickQuoteData = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetQuickQuoteDataExit

End Function







