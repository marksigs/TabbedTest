VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuotationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'QuotationBO' with actual class name
' *****************************************************************************

'Workfile:      QuotationBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Quotation Business Object which 'supports transactions' only.
'               Any transactioned requests will
'               be delegated to QuotationTxBO

'Dependencies:  QuotationTxBO, QuotationDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    09/11/99    Created
'AY     21/03/00    Change to QuotationSummary XML comments
'JLD    28/03/00    Enabled FindSubquoteDetails fnc
'SR     06/04/00    Modified method 'ValidateQuotationForReinstatement' - Add tags for
'                   validBCSubquote and ValidPPSubQuote only when they exist in validated
'                   quotation details. Do not raise error, if they do not exist
'MS     04/07/00    SYS1048 GetAcceptedBCSubQuoteNumber Method Added
'BG     16/08/00    SYS1239 Changed IsQuotationAffordable so it was comparing the total
'                   outgoings with income.  Previously it compared total outgoings with the
'                   cost of the quote making it "not affordable" most of the time.
'PSC    05/03/01    SYS1879 Added GetQuotationLoanSummary
'PSC    19/03/01    SYS2053 Add GetQuotationDetails
'MDC    28/08/01    SYS2250 Add UpdateLinkedQuotations method
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5 - Add further information to Quotation returned
'MV     14/09/2005  MAR35 - New UpdateLastOfferedQuotation() , GetAllQuotationsForApplication()
'                   IQuotationBO_UpdateLastOfferedQuotation();IQuotationBO_GetAllQuotationsForApplication()
'------------------------------------------------------------------------------------------

Option Explicit
Implements IQuotationBO
Implements ObjectControl
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrROOT_NODE_NAME = "QUOTATION"

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist


Public Function UpdateLinkedQuotations(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update quotations linked to a remodelled MortgageSubQuote
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLinkedQuotationsVbErr

    Const strFunctionName As String = "UpdateLinkedQuotations"


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
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_UpdateLinkedQuotations(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLinkedQuotations = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateLinkedQuotationsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateLinkedQuotationsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    UpdateLinkedQuotations = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLinkedQuotationsExit

End Function

Private Function IQuotationBO_UpdateLastOfferedQuotation(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo UpdateLastOfferedQuotationVbErr
    
    Const strFunctionName As String = "IQuotationBO_UpdateLastOfferedQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objIQuotationDO As IQuotationDO
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim strQuotationNumber As String
    Dim strNewQuotationNo  As String
    Dim xmlLastOfferedQuoteNode  As IXMLDOMNode
    
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationTxBO = m_objContext.CreateInstance(App.Title & ".QuotationTxBO")
        Set objIQuotationDO = m_objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
    'Get Mandatory Request items
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "//APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "//APPLICATIONFACTFINDNUMBER")
    strQuotationNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "//QUOTATIONNUMBER")
    
    'Find all Quotations for the requested Application
    Set xmlTempResponseNode = objIQuotationDO.GetAllQuotationsForApplication(vxmlTableElement)
    Set xmlNodeList = xmlTempResponseNode.selectNodes("//QUOTATION")
    
    For Each xmlNode In xmlNodeList
    
        strNewQuotationNo = m_objXmlAssist.GetMandatoryElementText(xmlNode, "QUOTATIONNUMBER")
        
        Set xmlLastOfferedQuoteNode = m_objXmlAssist.GetNode(xmlNode, "LASTOFFERED")
        
        'If Node does not exist then create one
        If xmlLastOfferedQuoteNode Is Nothing Then
            Set xmlLastOfferedQuoteNode = xmlOut.createElement("LASTOFFERED")
            xmlNode.appendChild xmlLastOfferedQuoteNode
        End If
        
        'Set the values
        If strNewQuotationNo = strQuotationNumber Then
            xmlLastOfferedQuoteNode.Text = 1
        Else
            xmlLastOfferedQuoteNode.Text = 0
        End If
        
        'Update the node
        objIQuotationTxBO.Update xmlNode
        
    Next xmlNode
  
    Set IQuotationBO_UpdateLastOfferedQuotation = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateLastOfferedQuotationExit:
    
    Set objIQuotationDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateLastOfferedQuotationVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_UpdateLastOfferedQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLastOfferedQuotationExit

End Function

Private Function IQuotationBO_GetAllQuotationsForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetAllQuotationsForApplicationVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetAllQuotationsForApplication"

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
    Dim objIQuotationDO As IQuotationDO
    Dim xmlRequestClone As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
    ' Clone the input so that we can add default country and language
    Set xmlRequestClone = vxmlTableElement.cloneNode(True)
                
    If xmlRequestClone.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlTableElement
    Else
        Set xmlRequestNode = xmlRequestClone.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objIQuotationDO.GetAllQuotationsForApplication(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IQuotationBO_GetAllQuotationsForApplication = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetAllQuotationsForApplicationExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestClone = Nothing

    Exit Function

GetAllQuotationsForApplicationVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetAllQuotationsForApplication = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
End Function


Public Function GetQuotationSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get summary details of the active quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <QUOTATIONSUMMARY>
'           <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           <TOTALPPMONTHLYCOST>Value</TOTALPPMONTHLYCOST>
'           <PPCOVERTYPE>Value</PPCOVERTYPE>
'           <MORTGAGECOVERFORAPPLICANT1>Value</MORTGAGECOVERFORAPPLICANT1>
'           <MORTGAGECOVERFORAPPLICANT2>Value</MORTGAGECOVERFORAPPLICANT2>
'           <TOTALBCMONTHLYCOST>Value</TOTALBCMONTHLYCOST>
'           <BCCOVERTYPE>Value</BCCOVERTYPE>
'           <BCFREQUENCY>Value</BCFREQUENCY>
'           <ACCIDENTALBUILDCOVERREQUIRED>Value</ACCIDENTALBUILDCOVERREQUIRED>
'           <ACCIDENTALCONTENTCOVERREQUIRED>Value</ACCIDENTALCONTENTCOVERREQUIRED>
'           <LOANCOMPONENTLIST>
'               <LOANCOMPONENT>
'                   <LOANCOMPONENTSEQUENCENUMBER>Value</LOANCOMPONENTSEQUENCENUMBER>
'                   <REPAYMENTMETHOD>Value</REPAYMENTMETHOD>
'                   <TOTALLOANCOMPONENTAMOUNT>Value</TOTALLOANCOMPONENTAMOUNT>
'                   <TERMINMONTHS>Value</TERMINMONTHS>
'                   <TERMINYEARS>Value</TERMINYEARS>
'                   <NETMONTHLYCOST>Value</NETMONTHLYCOST>
'                   <MORTGAGEPRODUCTCODE>Value</MORTGAGEPRODUCTCODE>
'                   <PRODUCTNAME>Value</PRODUCTNAME>
'                   <RATETYPE>Value</RATETYPE>
'                   <RATE>Value</RATE>
'               </LOANCOMPONENT>
'               ...
'               ...
'           </LOANCOMPONENTLIST>
'           <LIFEBENEFITLIST>
'               <LIFEBENEFIT>
'                   <LIFEBENEFITSEQUENCENUMBER>Value</LIFEBENEFITSEQUENCENUMBER>
'                   <CUSTOMERNUMBER1>Value</CUSTOMERNUMBER1>
'                   <CUSTOMERNUMBER2>Value</CUSTOMERNUMBER2>
'                   <BENEFITTYPE>Value</BENEFITTYPE>
'                   <MONTHLYCOST>Value</MONTHLYCOST>
'               </LIFEBENEFIT>
'               ...
'               ...
'           </LIFEBENEFITLIST>
'       </QUOTATIONSUMMARY>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationSummaryVbErr

    Const strFunctionName As String = "GetQuotationSummary"

    
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
    Set xmlTempResponseNode = IQuotationBO_GetQuotationSummary(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetQuotationSummary = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetQuotationSummaryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetQuotationSummaryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetQuotationSummary = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetQuotationSummaryExit

End Function

Public Function SaveOneOffCostDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveOneOffCostDetailsVbErr

    Const strFunctionName As String = "SaveOneOffCostDetails"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_SaveOneOffCostDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveOneOffCostDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveOneOffCostDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveOneOffCostDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    SaveOneOffCostDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOneOffCostDetailsExit
    
End Function

Public Function GetValidatedQuotationDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets quotation details and whether they are valid
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ISMEMBEROFATAFF>Value</ISMEMBEROFATAFF>
'               <ISSELFEMPLOYED>Value</ISSELFEMPLOYED>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
 On Error GoTo GetValidatedQuotationDetailsVbErr

    Const strFunctionName As String = "GetValidatedQuotationDetails"

    
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
    Set xmlTempResponseNode = IQuotationBO_GetValidatedQuotationDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetValidatedQuotationDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetValidatedQuotationDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetValidatedQuotationDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetValidatedQuotationDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetValidatedQuotationDetailsExit

End Function

Public Function ValidateQuotationForReinstatement(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Validates the sub quotes on the quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ISMEMBEROFATAFF>Value</ISMEMBEROFATAFF>
'               <ISSELFEMPLOYED>Value</ISSELFEMPLOYED>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'               <LTV>Value</LTV>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ValidateQuotationForReinstatementVbErr

    Const strFunctionName As String = "ValidateQuotationForReinstatement"

    
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
    Set xmlTempResponseNode = IQuotationBO_ValidateQuotationForReinstatement(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    ValidateQuotationForReinstatement = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ValidateQuotationForReinstatementExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateQuotationForReinstatementVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ValidateQuotationForReinstatement = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateQuotationForReinstatementExit

End Function

Public Function FindStoredQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets details of quotations for a fact find
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <STOREDQUOTATIONLIST>
'           <STOREDQUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <QUOTATIONTYPE>Value<<QUOTATIONTYPE>>
'               <TOTALQUOTATIONCOST>Value</TOTALQUOTATIONCOST>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'               <LIFECOVER>Value</LIFECOVER>
'               <BCCOVERTYPE>Value</BCCOVERTYPE>
'               <PPCOVERTYPE>Value<PPCOVERTYPE>
'               <LOANCOMPONENTLIST>
'                   <LOANCOMPONENT>
'                       <LOANCOMPONENTSEQUENCENUMBER>Value</LOANCOMPONENTSEQUENCENUMBER>
'                       <REPAYMENTMETHOD>Value</REPAYMENTMETHOD>
'                   </LOANCOMPONENT>
'                   ...
'                   ...
'               </LOANCOMPONENTLIST>
'           </STOREDQUOTATION>
'       </STOREDQUOTATIONLIST>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo FindStoredQuoteDetailsVbErr

    Const strFunctionName As String = "FindStoredQuoteDetails"

    
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
    Set xmlTempResponseNode = IQuotationBO_FindStoredQuoteDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindStoredQuoteDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindStoredQuoteDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Public Function IsQuotationAffordable(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks whether a quotation is affordable
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo IsQuotationAffordableVbErr

    Const strFunctionName As String = "IsQuotationAffordable"

    
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
    Set xmlTempResponseNode = IQuotationBO_IsQuotationAffordable(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    IsQuotationAffordable = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

IsQuotationAffordableExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

IsQuotationAffordableVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    IsQuotationAffordable = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume IsQuotationAffordableExit

End Function

Public Function FindSubQuotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <BUSINESSTYPE>Value</BUSINESSTYPE>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
 On Error GoTo FindSubQuotesVbErr

    Const strFunctionName As String = "FindSubQuotes"

    
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
    Set xmlTempResponseNode = IQuotationBO_FindSubQuotes(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindSubQuotes = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindSubQuotesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    FindSubQuotes = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindSubQuotesExit

End Function

Public Function ReinstateSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Reinstates a particular sub quote to be the active sub quote
'               invokes appropriate transactioning (Tx) business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <UPDATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ACTIVESUBQUOTENUMBER>Value</ACTIVESUBQUOTENUMBER>
'               <SELECTEDSUBQUOTENUMBER>Value</SELECTEDSUBQUOTENUMBER>
'               <BUSINESSTYPE>Value</BUSINESSTYPE>
'           </QUOTATION>
'       </UPDATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateSubQuoteVbErr

    Const strFunctionName As String = "ReinstateSubQuote"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_ReinstateSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ReinstateSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

ReinstateSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ReinstateSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    ReinstateSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateSubQuoteExit
    
End Function

Public Function CreateFirstMortgageLifeSubQuotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <QUOTATIONTYPE>Value</QUOTATIONTYPE>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstMortgageLifeSubQuotesVbErr

    Const strFunctionName As String = "CreateFirstMortgageLifeSubQuotes"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateFirstMortgageLifeSubQuotes(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateFirstMortgageLifeSubQuotes = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateFirstMortgageLifeSubQuotesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateFirstMortgageLifeSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateFirstMortgageLifeSubQuotes = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstMortgageLifeSubQuotesExit

End Function

Public Function CreateNewMortgageLifeSubQuotes(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <QUOTATIONTYPE>Value</QUOTATIONTYPE>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'               <LIFESUBQUOTENUMBER>Value</LIFESUBQUOTENUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewMortgageLifeSubQuotesVbErr

    Const strFunctionName As String = "CreateNewMortgageLifeSubQuotes"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateNewMortgageLifeSubQuotes(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateNewMortgageLifeSubQuotes = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateNewMortgageLifeSubQuotesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewMortgageLifeSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateNewMortgageLifeSubQuotes = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewMortgageLifeSubQuotesExit
    
End Function

Public Function CreateFirstBCSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <CONTEXT>Value</CONTEXT>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstBCSubQuoteVbErr

    Const strFunctionName As String = "CreateFirstBCSubQuote"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateFirstBCSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateFirstBCSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateFirstBCSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateFirstBCSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateFirstBCSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstBCSubQuoteExit

End Function

Public Function CreateNewBCSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <CONTEXT>Value</CONTEXT>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <BCSUBQUOTENUMBER>Value</BCSUBQUOTENUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewBCSubQuoteVbErr

    Const strFunctionName As String = "CreateNewBCSubQuote"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateNewBCSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateNewBCSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateNewBCSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewBCSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateNewBCSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewBCSubQuoteExit

End Function

Public Function StoreQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'
'   Input XML
'
'   <REQUEST>
'       <UPDATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <TOTALQUOTATIONCOST>Value</TOTALQUOTATIONCOST>
'           </QUOTATION>
'       </UPDATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo StoreQuotationVbErr

    Const strFunctionName As String = "StoreQuotation"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_StoreQuotation(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    StoreQuotation = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

StoreQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

StoreQuotationVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    StoreQuotation = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume StoreQuotationExit
    
End Function

Public Function CreateNewQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a new quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewQuotationVbErr

    Const strFunctionName As String = "CreateNewQuotation"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateNewQuotation(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateNewQuotation = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateNewQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewQuotationVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateNewQuotation = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit

End Function

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
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
    Set xmlTempResponseNode = IQuotationBO_GetData(xmlIn.documentElement)
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
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetData = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Public Function Update(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"

    
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
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_Update(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Update = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Update = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
    
End Function

Public Function UpdateLastOfferedQuotation(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo UpdateLastOfferedQuotationVbErr

    Const strFunctionName As String = "UpdateLastOfferedQuotation"

    Dim xmlIn As FreeThreadedDOMDocument40
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IQuotationBO_Validate(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_UpdateLastOfferedQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateLastOfferedQuotation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateLastOfferedQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateLastOfferedQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    UpdateLastOfferedQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLastOfferedQuotationExit
    
End Function

Public Function IsMortgageSubQuoteAttachedToOneQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks to see if a mortgage sub quote is attached to more than one
'               quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo IsMortgageSubQuoteAttachedToOneQuoteVbErr

    Const strFunctionName As String = "IsMortgageSubQuoteAttachedToOneQuote"

    
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
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = _
                IQuotationBO_IsMortgageSubQuoteAttachedToOneQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    IsMortgageSubQuoteAttachedToOneQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

IsMortgageSubQuoteAttachedToOneQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

IsMortgageSubQuoteAttachedToOneQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    IsMortgageSubQuoteAttachedToOneQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume IsMortgageSubQuoteAttachedToOneQuoteExit

End Function

Public Function CreateFirstPPSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstPPSubQuoteVbErr

    Const strFunctionName As String = "CreateFirstPPSubQuote"

    
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
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateFirstPPSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateFirstPPSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateFirstPPSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateFirstPPSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateFirstPPSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstPPSubQuoteExit

End Function

Public Function CreateNewPPSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPPSubQuoteVbErr

    Const strFunctionName As String = "CreateNewPPSubQuote"

    
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
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuotationBO_CreateNewPPSubQuote(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CreateNewPPSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateNewPPSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewPPSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    CreateNewPPSubQuote = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPPSubQuoteExit

End Function

Private Function IQuotationBO_CreateFirstBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the first Buildings and Contents sub quote under a quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <CONTEXT>Value</CONTEXT>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstBCSubQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateFirstBCSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIQuotationTxBO.CreateFirstBCSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateFirstBCSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstBCSubQuoteExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateFirstBCSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateFirstBCSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstBCSubQuoteExit

End Function

Private Function IQuotationBO_CreateFirstMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the first Mortgage and Life sub quotes under a quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <QUOTATIONTYPE>Value</QUOTATIONTYPE>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>

'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstMortgageLifeSubQuotesVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateFirstMortgageLifeSubQuotes"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
                    objIQuotationTxBO.CreateFirstMortgageLifeSubQuotes(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateFirstMortgageLifeSubQuotes = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstMortgageLifeSubQuotesExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateFirstMortgageLifeSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateFirstMortgageLifeSubQuotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstMortgageLifeSubQuotesExit

End Function

Private Function IQuotationBO_CreateNewBCSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the first Buildings and Contents sub quote under a quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <CONTEXT>Value</CONTEXT>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <BCSUBQUOTENUMBER>Value</BCSUBQUOTENUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewBCSubQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateNewBCSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIQuotationTxBO.CreateNewBCSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateNewBCSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewBCSubQuoteExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewBCSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateNewBCSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewBCSubQuoteExit

End Function

Private Function IQuotationBO_CreateNewMortgageLifeSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates the first Mortgage and Life sub quotes under a quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <QUOTATIONTYPE>Value</QUOTATIONTYPE>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'               <LIFESUBQUOTENUMBER>Value</LIFESUBQUOTENUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewMortgageLifeSubQuotesVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateNewMortgageLifeSubQuotes"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = _
                    objIQuotationTxBO.CreateNewMortgageLifeSubQuotes(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateNewMortgageLifeSubQuotes = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewMortgageLifeSubQuotesExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewMortgageLifeSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateNewMortgageLifeSubQuotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewMortgageLifeSubQuotesExit

End Function

Private Function IQuotationBO_CreateNewQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a new Quotation
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <CREATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </CREATE>
'   </REQUEST>

'------------------------------------------------------------------------------------------
On Error GoTo CreateNewQuotationVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateNewQuotation"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIQuotationTxBO.CreateNewQuotation(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateNewQuotation = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewQuotationExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewQuotationVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateNewQuotation = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit

End Function

Private Function IQuotationBO_FindStoredQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets details of quotations for a fact find
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <STOREDQUOTATIONLIST>
'           <STOREDQUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <QUOTATIONTYPE>Value<<QUOTATIONTYPE>>
'               <TOTALQUOTATIONCOST>Value</TOTALQUOTATIONCOST>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'               <LIFECOVER>Value</LIFECOVER>
'               <TOTALNETMONTHLYCOST>Value</TOTALNETMONTHLYCOST>
'               <TOTALBCMONTHLYCOST>Value</TOTALBCMONTHLYCOST>
'               <TOTALPPMONTHLYCOST>Value</TOTALPPMONTHLYCOST>
'               <BCCOVERTYPE>Value</BCCOVERTYPE>
'               <ANNUALBUILDINGSPREMIUM>Value</ANNUALBUILDINGSPREMIUM>
'               <ANNUALCONTENTSPREMIUM>Value</ANNUALCONTENTSPREMIUM>
'               <PPCOVERTYPE>Value<PPCOVERTYPE>
'               <LOANCOMPONENTLIST>
'                   <LOANCOMPONENT>
'                       <LOANCOMPONENTSEQUENCENUMBER>Value</LOANCOMPONENTSEQUENCENUMBER>
'                       <REPAYMENTMETHOD>Value</REPAYMENTMETHOD>
'                       <APR>Value</APR>
'                       <FINALRATEMONTHLYCOST>Value</FINALRATEMONTHLYCOST>
'                       <LOANAMOUNT>Value</LOANAMOUNT>
'                       <CAPITALANDINTERESTELEMENT>Value</CAPITALANDINTERESTELEMENT>
'                       <INTERESTONLYELEMENT>Value</INTERESTONLYELEMENT>
'                       <PRODUCTNAME>Value</PRODUCTNAME>
'                       <GROUPTYPE>Value</GROUPTYPE>
'                   </LOANCOMPONENT>
'                   ...
'                   ...
'               </LOANCOMPONENTLIST>
'           </STOREDQUOTATION>
'       </STOREDQUOTATIONLIST>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo FindStoredQuoteDetailsVbErr
    
    Const strFunctionName As String = "IQuotationBO_FindStoredQuoteDetails"


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
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIQuotationDO.StoredQuoteFindList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IQuotationBO_FindStoredQuoteDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
FindStoredQuoteDetailsExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_FindStoredQuoteDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit


End Function

Private Function IQuotationBO_FindSubQuotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <BUSINESSTYPE>Value</BUSINESSTYPE>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo FindSubQuotesVbErr
    
    Const strFunctionName As String = "IQuotationBO_FindSubQuotes"


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
    Dim objIMortgageSubQuoteDO As IMortgageSubQuoteDO
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
    Dim xmlFindListDoc As New FreeThreadedDOMDocument40
    xmlFindListDoc.validateOnParse = False
    xmlFindListDoc.setProperty "NewParser", True
    Dim xmlFindListElem As IXMLDOMElement
    Dim xmlPPResponseDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
    If objContext Is Nothing Then
        Set objIMortgageSubQuoteDO = New MortgageSubQuoteDO
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objIMortgageSubQuoteDO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteDO")
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    Set xmlFindListElem = xmlRequestNode.cloneNode(True)
    xmlFindListDoc.appendChild xmlFindListElem
    
    Dim strBusinessType As String
    strBusinessType = objXmlAssist.GetTagValue(xmlFindListElem, "BUSINESSTYPE")

    If StrComp(strBusinessType, "Mortgage", vbTextCompare) = 0 Then
        objXmlAssist.ChangeNodeName xmlFindListElem, "QUOTATION", "MORTGAGESUBQUOTE"
        Set xmlDataNode = objIMortgageSubQuoteDO.FindSubQuoteDetails(xmlFindListElem)
    ElseIf StrComp(strBusinessType, "BC", vbTextCompare) = 0 Then
        objXmlAssist.ChangeNodeName xmlFindListElem, "QUOTATION", "BUILDINGSANDCONTENTSSUBQUOTE"
        Set xmlDataNode = objIBCSubQuoteDO.FindList(xmlFindListElem)
    ElseIf StrComp(strBusinessType, "PP", vbTextCompare) = 0 Then
        objXmlAssist.ChangeNodeName xmlFindListElem, "QUOTATION", "PAYMENTPROTECTIONSUBQUOTE"
        strResponse = objPPSubQuoteDO.FindList(xmlFindListElem.xml)
        Set xmlPPResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        Set xmlDataNode = xmlPPResponseDoc.documentElement
    Else
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "BUSINESSTYPE = " & strBusinessType
                                
    End If
        
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IQuotationBO_FindSubQuotes = xmlResponseElem.cloneNode(True)

FindSubQuotesExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageSubQuoteDO = Nothing
    Set objIBCSubQuoteDO = Nothing
    Set objPPSubQuoteDO = Nothing
    
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindSubQuotesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_FindSubQuotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindSubQuotesExit

End Function

Private Function IQuotationBO_GetData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetData"


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
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIQuotationDO.GetData(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IQuotationBO_GetData = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetDataExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
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
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetData = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetDataExit

End Function

Private Function IQuotationBO_GetQuotationSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get summary details of the active quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <QUOTATIONSUMMARY>
'           <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           <TOTALPPMONTHLYCOST>Value</TOTALPPMONTHLYCOST>
'           <PPCOVERTYPE>Value</PPCOVERTYPE>
'           <MORTGAGECOVERFORAPPLICANT1>Value</MORTGAGECOVERFORAPPLICANT1>
'           <MORTGAGECOVERFORAPPLICANT2>Value</MORTGAGECOVERFORAPPLICANT2>
'           <TOTALBCMONTHLYCOST>Value</TOTALBCMONTHLYCOST>
'           <BCCOVERTYPE>Value</BCCOVERTYPE>
'           <BCFREQUENCY>Value</BCFREQUENCY>
'           <ACCIDENTALBUILDCOVERREQUIRED>Value</ACCIDENTALBUILDCOVERREQUIRED>
'           <ACCIDENTALCONTENTCOVERREQUIRED>Value</ACCIDENTALCONTENTCOVERREQUIRED>
'           <LOANCOMPONENTLIST>
'               <LOANCOMPONENT>
'                   <LOANCOMPONENTSEQUENCENUMBER>Value</LOANCOMPONENTSEQUENCENUMBER>
'                   <REPAYMENTMETHOD>Value</REPAYMENTMETHOD>
'                   <TOTALLOANCOMPONENTAMOUNT>Value</TOTALLOANCOMPONENTAMOUNT>
'                   <TERMINMONTHS>Value</TERMINMONTHS>
'                   <TERMINYEARS>Value</TERMINYEARS>
'                   <NETMONTHLYCOST>Value</NETMONTHLYCOST>
'                   <MORTGAGEPRODUCTCODE>Value</MORTGAGEPRODUCTCODE>
'                   <PRODUCTNAME>Value</PRODUCTNAME>
'                   <RATETYPE>Value</RATETYPE>
'                   <RATE>Value</RATE>
'               </LOANCOMPONENT>
'               ...
'               ...
'           </LOANCOMPONENTLIST>
'           <LIFEBENEFITLIST>
'               <LIFEBENEFIT>
'                   <LIFEBENEFITSEQUENCENUMBER>Value</LIFEBENEFITSEQUENCENUMBER>
'                   <CUSTOMERNUMBER1>Value</CUSTOMERNUMBER1>
'                   <CUSTOMERNUMBER2>Value</CUSTOMERNUMBER2>
'                   <BENEFITTYPE>Value</BENEFITTYPE>
'                   <MONTHLYCOST>Value</MONTHLYCOST>
'               </LIFEBENEFIT>
'               ...
'               ...
'           </LIFEBENEFITLIST>
'       </QUOTATIONSUMMARY>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationSummaryVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetQuotationSummary"


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
    Dim objIQuotationDO As IQuotationDO
    Dim xmlRequestClone As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
    ' Clone the input so that we can add default country and language
    Set xmlRequestClone = vXMLRequest.cloneNode(True)
                
    If xmlRequestClone.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = xmlRequestClone.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
    ' Add in the default parameters for Country of origin and Language
    AddQuoteSummaryDefaults xmlRequestNode
    
    Set xmlDataNode = objIQuotationDO.GetQuotationSummary(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IQuotationBO_GetQuotationSummary = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetQuotationSummaryExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestClone = Nothing

    Exit Function

GetQuotationSummaryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetQuotationSummary = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetQuotationSummaryExit

End Function

Private Function IQuotationBO_IsMortgageSubQuoteAttachedToOneQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Checks to see if a mortgage sub quote is attached to more than one
'               quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo IsMortgageSubQuoteAttachedToOneQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_IsMortgageSubQuoteAttachedToOneQuote"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Dim intNoOfQuotes
    
    intNoOfQuotes = objIQuotationDO.CountQuotationsForMortgageSubQuote(xmlRequestNode)
    
    If intNoOfQuotes > 1 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 177
    End If
        
    Set IQuotationBO_IsMortgageSubQuoteAttachedToOneQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
IsMortgageSubQuoteAttachedToOneQuoteExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

IsMortgageSubQuoteAttachedToOneQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_IsMortgageSubQuoteAttachedToOneQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume IsMortgageSubQuoteAttachedToOneQuoteExit

End Function

Private Function IQuotationBO_CreateFirstPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstPPSubQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateFirstPPSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIQuotationTxBO.CreateFirstPPSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateFirstPPSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstPPSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    
    Exit Function

CreateFirstPPSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateFirstPPSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstPPSubQuoteExit


End Function

Private Function IQuotationBO_CreateNewPPSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPPSubQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_CreateNewPPSubQuote"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    Dim objXmlAssist As New XMLAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIQuotationTxBO.CreateNewPPSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
   
    Set IQuotationBO_CreateNewPPSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewPPSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objXmlAssist = Nothing
    
    Exit Function

CreateNewPPSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_CreateNewPPSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPPSubQuoteExit

End Function

Private Function IQuotationBO_GetValidatedQuotationDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets quotation details and whether they are valid
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ISMEMBEROFATAFF>Value</ISMEMBEROFATAFF>
'               <ISSELFEMPLOYED>Value</ISSELFEMPLOYED>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'               <LTV>Value</LTV>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo GetValidatedQuotationDetailsVbErr
    'GD BMIDS00256
    Const strFunctionName As String = "IQuotationBO_GetValidatedQuotationDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIQuotationDO As IQuotationDO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objIBCSubQuoteBO As IBuildingsAndContentsSubQuoteBO
    Dim objIPPSubQuoteBO As IPaymentProtectionSubQuoteBO

    Dim xmlAppDataRequest As IXMLDOMElement
    Dim strResponse As String
    Dim xmlRequestClone As IXMLDOMElement
    Dim xmlElement1 As IXMLDOMElement
    Dim xmlElement2 As IXMLDOMElement
    Dim xmlElement3 As IXMLDOMElement
    Dim xmlTempDoc As New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Dim xmlTempResponse As IXMLDOMElement
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlQuoteDetailsElem As IXMLDOMElement
    Dim xmlBCSubQuoteElem As IXMLDOMElement
    Dim xmlPPSubQuoteElem As IXMLDOMElement
    Dim xmlMortSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteResponseElem As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlRequestNodeClone As IXMLDOMNode
    Dim xmlMemberOfStaffElem As IXMLDOMElement
    ''BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
    'Dim xmlSelfEmployedElem As IXMLDOMElement
    Dim xmlPortedLoansElem As IXMLDOMElement
    Dim xmlAmountRequested As IXMLDOMElement
    Dim xmlLTVElem As IXMLDOMElement            'BMIDS00515 SA
    
    'MV - 18/07/2002 - BMIDS00179 - Core Upgrade
    'SG 28/05/02 SYS4767
    'Dim xmlIsQuickQuoteElem As IXMLDOMElement
    
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strActiveQuoteNumber As String
    Dim strMortgageSubQuoteNumber As String
    Dim strAmountRequested As String
    



    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If vxmlTableElement.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestElem = vxmlTableElement
    Else
        Set xmlRequestElem = vxmlTableElement.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
   
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If

    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objIBCSubQuoteBO = New BuildingsAndContentsSubQuoteBO
        Set objIPPSubQuoteBO = New PaymentProtectionSubQuoteBO

    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
        Set objIMortgageSubQuoteBO = objContext.CreateInstance(App.Title & ".MortgageSubQuoteBO")
        Set objIBCSubQuoteBO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteBO")
        Set objIPPSubQuoteBO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteBO")
    End If

    Dim strValidMortgageSubQuote As String
    Dim strValidBCSubQuote As String
    Dim strValidPPSubQuote As String
    Dim strTagName As String
    ' ** This must be done before we check the validity of the Mortgage, in case the LTV is out of range **
    ' Recalculate LTV using current purchase price, and loan amount from the active mortgagesubquote
    
    ' Get the application data and latest application fact find records
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, _
                                                       "APPLICATIONNUMBER")
    strActiveQuoteNumber = objXmlAssist.GetTagValue(vxmlTableElement, "ACTIVEQUOTENUMBER")
    Set xmlRequestClone = vxmlTableElement.cloneNode(True)
    Set xmlQuoteDetailsElem = objIQuotationDO.GetQuotationDetails(xmlRequestElem)
    strMortgageSubQuoteNumber = objXmlAssist.GetTagValue(xmlQuoteDetailsElem, "MORTGAGESUBQUOTENUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlQuoteDetailsElem, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlRequestNodeClone = objXmlAssist.GetRequestNodeEx(vxmlTableElement)

    xmlRequestNodeClone.appendChild xmlQuoteDetailsElem
    'DC AQR SYS0977 - Rolled back
 '   xmlRequestNodeClone.appendChild objXmlAssist.GetNode(vxmlTableElement, ".//LTV")
        
    Set xmlElement = xmlOut.createElement("MORTGAGESUBQUOTE")
    Set xmlElement1 = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement1.Text = strApplicationNumber
    Set xmlElement2 = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement2.Text = strApplicationFactFindNumber
    Set xmlElement3 = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElement3.Text = strMortgageSubQuoteNumber

    xmlElement.appendChild xmlElement1
    xmlElement.appendChild xmlElement2
    xmlElement.appendChild xmlElement3

    Set xmlTempResponse = objIMortgageSubQuoteBO.GetData(xmlElement)
    strAmountRequested = objXmlAssist.GetTagValue(xmlTempResponse, "AMOUNTREQUESTED")
    Set xmlAmountRequested = xmlTempResponse.selectSingleNode(".//AMOUNTREQUESTED")
    ' If mortgagesubquotenumber is not null validate the quote
    strValidMortgageSubQuote = objXmlAssist.GetTagValue(xmlQuoteDetailsElem, _
                                                       "MORTGAGESUBQUOTENUMBER")
                                                                           
    If Len(strValidMortgageSubQuote) > 0 Then
    
        strTagName = "MORTGAGESUBQUOTE"
        Set xmlMortSubQuoteElem = xmlQuoteDetailsElem.getElementsByTagName(strTagName).Item(0)
   
        ' Clone nodes to attach to the request to go down to the mortgage subquote validate
        Set xmlMemberOfStaffElem = xmlRequestElem.getElementsByTagName("ISMEMBEROFSTAFF").Item(0)
        
        If Not xmlMemberOfStaffElem Is Nothing Then
            Set xmlMemberOfStaffElem = xmlMemberOfStaffElem.cloneNode(True)
            xmlMortSubQuoteElem.appendChild xmlMemberOfStaffElem
        End If
        'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
        'Set xmlSelfEmployedElem = xmlRequestElem.getElementsByTagName("ISSELFEMPLOYED").Item(0)
        
        'If Not xmlSelfEmployedElem Is Nothing Then
        '    Set xmlSelfEmployedElem = xmlSelfEmployedElem.cloneNode(True)
        '    xmlMortSubQuoteElem.appendChild xmlSelfEmployedElem
        'End If
    
        Set xmlPortedLoansElem = xmlRequestElem.getElementsByTagName("LOANSTOBEPORTED").Item(0)
        
        If Not xmlPortedLoansElem Is Nothing Then
            Set xmlPortedLoansElem = xmlPortedLoansElem.cloneNode(True)
            xmlMortSubQuoteElem.appendChild xmlPortedLoansElem
        End If
        xmlMortSubQuoteElem.appendChild xmlAmountRequested
        'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
        'SG 28/05/02 SYS4767 START
        'JLD MSMS0034 send down the ISQUICKQUOTE node too
        'Set xmlIsQuickQuoteElem = xmlRequestElem.getElementsByTagName("ISQUICKQUOTE").Item(0)
        
        'If Not xmlIsQuickQuoteElem Is Nothing Then
        '    Set xmlIsQuickQuoteElem = xmlIsQuickQuoteElem.cloneNode(True)
        '    xmlMortSubQuoteElem.appendChild xmlIsQuickQuoteElem
        'End If
        'SG 28/05/02 SYS4767 END
        'End
        
        'BMIDS00515 SA need to check LTV {
        '--BMIDS00794 LTV is part of response xml
        'Set xmlLTVElem = xmlRequestElem.getElementsByTagName("LTV").Item(0)
        '--BMIDS00794.
        '--Set xmlLTVElem = xmlTempResponse.getElementsByTagName("LTV").Item(0)
        'SA 11/11/02 BMIDS00794/515 In all cases, if there is an existing mortgagesubquote
        ' with an amount requested, the LTV is passed in to this method.
        Set xmlLTVElem = vxmlTableElement.getElementsByTagName("LTV").Item(0)
        
        If Not xmlLTVElem Is Nothing Then
            Set xmlLTVElem = xmlLTVElem.cloneNode(True)
            xmlMortSubQuoteElem.appendChild xmlLTVElem
        End If
        'BMIDS00515 SA }
                
        Set xmlSubQuoteResponseElem = _
                            objIMortgageSubQuoteBO.ValidateSubQuote(xmlRequestNodeClone)
        objErrAssist.CheckResponse xmlSubQuoteResponseElem, xmlResponseElem, True

        strTagName = "VALIDMORTGAGESUBQUOTE"
        Set xmlElement = xmlSubQuoteResponseElem.getElementsByTagName(strTagName).Item(0)


        
        xmlMortSubQuoteElem.appendChild xmlElement
        
        ' Remove the nodes we added in for mortgage subquote validation
        If Not xmlMemberOfStaffElem Is Nothing Then
            xmlMortSubQuoteElem.removeChild xmlMemberOfStaffElem
        End If
        'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
        'If Not xmlSelfEmployedElem Is Nothing Then
        '    xmlMortSubQuoteElem.removeChild xmlSelfEmployedElem
        'End If
        
        If Not xmlPortedLoansElem Is Nothing Then
            xmlMortSubQuoteElem.removeChild xmlPortedLoansElem
        End If

    End If
           
    ' If bcsubquotenumber is not null validate the quote
    strValidBCSubQuote = objXmlAssist.GetTagValue(xmlQuoteDetailsElem, "BCSUBQUOTENUMBER")
    
    If Len(strValidBCSubQuote) > 0 Then

        strTagName = "BUILDINGSANDCONTENTSSUBQUOTE"
        Set xmlBCSubQuoteElem = xmlQuoteDetailsElem.getElementsByTagName(strTagName).Item(0)
        Set xmlSubQuoteResponseElem = objIBCSubQuoteBO.ValidateSubQuote(xmlRequestNodeClone)
        objErrAssist.CheckResponse xmlSubQuoteResponseElem, xmlResponseElem, True
        
        strTagName = "VALIDBCSUBQUOTE"
        Set xmlElement = xmlSubQuoteResponseElem.getElementsByTagName(strTagName).Item(0)

        xmlBCSubQuoteElem.appendChild xmlElement
    End If

    strValidPPSubQuote = objXmlAssist.GetTagValue(xmlQuoteDetailsElem, "PPSUBQUOTENUMBER")
    
    ' If ppsubquotenumber is not null validate the quote
    If Len(strValidPPSubQuote) > 0 Then

        strTagName = "PAYMENTPROTECTIONSUBQUOTE"
        Set xmlPPSubQuoteElem = xmlQuoteDetailsElem.getElementsByTagName(strTagName).Item(0)
        Set xmlSubQuoteResponseElem = objIPPSubQuoteBO.ValidateSubQuote(xmlRequestNodeClone)
        objErrAssist.CheckResponse xmlSubQuoteResponseElem, xmlResponseElem, True
        
        strTagName = "VALIDPPSUBQUOTE"
        Set xmlElement = xmlSubQuoteResponseElem.getElementsByTagName(strTagName).Item(0)

        xmlPPSubQuoteElem.appendChild xmlElement

    End If

    
    xmlResponseElem.appendChild xmlQuoteDetailsElem
    

    
    

    Set IQuotationBO_GetValidatedQuotationDetails = xmlResponseElem
    

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetValidatedQuotationDetailsExit:

    Set xmlOut = Nothing
    Set xmlData = Nothing
    Set xmlRequestElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIQuotationDO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set objIBCSubQuoteBO = Nothing
    Set objIPPSubQuoteBO = Nothing
    Set xmlElement = Nothing
    Set xmlResponseElem = Nothing
    Set xmlQuoteDetailsElem = Nothing
    Set xmlBCSubQuoteElem = Nothing
    Set xmlPPSubQuoteElem = Nothing
    Set xmlMortSubQuoteElem = Nothing
    Set xmlSubQuoteResponseElem = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlRequestNodeClone = Nothing
    Set xmlMemberOfStaffElem = Nothing
    'BMIDS00952 SA 15/11/2002 Self Employed no longer available as product option so remove it here.
    'Set xmlSelfEmployedElem = Nothing
    Set xmlPortedLoansElem = Nothing

    Exit Function

GetValidatedQuotationDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        'FIX ME
        'objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetValidatedQuotationDetails = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume GetValidatedQuotationDetailsExit

End Function

Private Function IQuotationBO_IsQuotationAffordable(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Checks whether a quotation is affordable
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo IsQuotationAffordableVbErr
    
    Const strFunctionName As String = "IQuotationBO_IsQuotationAffordable"


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
    Dim objIQuotationDO As IQuotationDO
    Dim objConvertAssist As New ConvertAssist
    Dim xmlElement As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIQuotationDO.GetData(xmlRequestNode)
    
    Dim lngLoansAndLiabilities As Long
    Dim lngMonthlyMortgagePayments As Long
    Dim lngMortgageRelatedIns As Long
    Dim lngOtherOutgoings As Long
    Dim lngTotalQuotationCost As Long
    Dim lngTotalOutgoings As Long
    Dim lngTotalMonthlyIncome As Long
    
    Dim strTagName As String
    Dim strTagValue As String
    Dim blnIsQuoteAffordable As Boolean
    
    strTagName = "LOANSANDLIABILITIES"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngLoansAndLiabilities = objConvertAssist.CSafeLng(strTagValue)
    
    strTagName = "MONTHLYMORTGAGEPAYMENTS"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngMonthlyMortgagePayments = objConvertAssist.CSafeLng(strTagValue)
    
    strTagName = "MORTGAGERELATEDINSURANCE"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngMortgageRelatedIns = objConvertAssist.CSafeLng(strTagValue)
    
    strTagName = "OTHEROUTGOINGS"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngOtherOutgoings = objConvertAssist.CSafeLng(strTagValue)
    
    strTagName = "TOTALQUOTATIONCOST"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngTotalQuotationCost = objConvertAssist.CSafeLng(strTagValue)
    
    strTagName = "TOTALMONTHLYINCOME"
    strTagValue = objXmlAssist.GetTagValue(xmlDataNode, strTagName)
    lngTotalMonthlyIncome = objConvertAssist.CSafeLng(strTagValue)
    
    lngTotalOutgoings = lngLoansAndLiabilities + lngMonthlyMortgagePayments + _
                        lngMortgageRelatedIns + lngOtherOutgoings
    'BG 16/08/00 SYS1239 See notes at top of page
    If lngTotalOutgoings <= lngTotalMonthlyIncome Then
        blnIsQuoteAffordable = True
    Else
        blnIsQuoteAffordable = False
    End If
    
    Set xmlElement = xmlOut.createElement("ISQUOTATIONAFFORDABLE")
    xmlElement.Text = Abs(blnIsQuoteAffordable)
    xmlResponseElem.appendChild xmlElement
        
    Set IQuotationBO_IsQuotationAffordable = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
IsQuotationAffordableExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objConvertAssist = Nothing
    Set xmlElement = Nothing

    Exit Function

IsQuotationAffordableVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_IsQuotationAffordable = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume IsQuotationAffordableExit

End Function

Private Function IQuotationBO_ReinstateSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Reinstates a particular sub quote to be the active sub quote
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <UPDATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ACTIVESUBQUOTENUMBER>Value</ACTIVESUBQUOTENUMBER>
'               <SELECTEDSUBQUOTENUMBER>Value</SELECTEDSUBQUOTENUMBER>
'               <BUSINESSTYPE>Value</BUSINESSTYPE>
'           </QUOTATION>
'       </UPDATE>
'   </REQUEST>
'
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateSubQuoteVbErr
    
    Const strFunctionName As String = "IQuotationBO_ReinstateSubQuote"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuotationTxBO.ReinstateSubQuote(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuotationBO_ReinstateSubQuote = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ReinstateSubQuoteExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ReinstateSubQuoteVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_ReinstateSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateSubQuoteExit

End Function

Private Function IQuotationBO_SaveOneOffCostDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveOneOffCostDetailsVbErr
    
    Const strFunctionName As String = "IQuotationBO_SaveOneOffCostDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuotationTxBO.SaveOneOffCostDetails(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuotationBO_SaveOneOffCostDetails = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveOneOffCostDetailsExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveOneOffCostDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_SaveOneOffCostDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveOneOffCostDetailsExit

End Function

Private Function IQuotationBO_StoreQuotation(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'
'   Input XML
'
'   <REQUEST>
'       <UPDATE>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <TOTALQUOTATIONCOST>Value</TOTALQUOTATIONCOST>
'           </QUOTATION>
'       </UPDATE>
'   </REQUEST>
'
'
'------------------------------------------------------------------------------------------
On Error GoTo StoreQuotationVbErr
    
    Const strFunctionName As String = "IQuotationBO_StoreQuotation"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuotationTxBO.StoreQuotation(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuotationBO_StoreQuotation = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
StoreQuotationExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

StoreQuotationVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_StoreQuotation = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume StoreQuotationExit

End Function

Private Function IQuotationBO_Update(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IQuotationBO_Update"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuotationTxBO As IQuotationTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objIQuotationTxBO = objContext.CreateInstance(App.Title & ".QuotationTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuotationTxBO.Update(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuotationBO_Update = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:

    Set objErrAssist = Nothing
    Set objIQuotationTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_Update = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IQuotationBO_UpdateLinkedQuotations(ByVal vXMLRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update quotations linked to a remodelled MortgageSubQuote
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLinkedComponentsVbErr
    
Const strFunctionName As String = "IQuotationBO_UpdateLinkedQuotations"

Dim objErrAssist As New ErrAssist
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim objIQuotationDO As IQuotationDO
Dim objConvAssist As ConvertAssist

Dim dblCurrentCost As Double
Dim dblOldCost As Double
Dim dblNewCost As Double


    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = m_objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    Set objConvAssist = New ConvertAssist
    
    'Get Mandatory Request items
    dblOldCost = objConvAssist.CSafeDbl(m_objXmlAssist.GetMandatoryElementText(vXMLRequest, "//OLDMORTGAGESUBQUOTECOST"))
    dblNewCost = objConvAssist.CSafeDbl(m_objXmlAssist.GetMandatoryElementText(vXMLRequest, "//NEWMORTGAGESUBQUOTECOST"))
    
    'Find all Quotations for the requested Application/MortgageSubQuote
    Set xmlTempResponseNode = objIQuotationDO.FindList(vXMLRequest)
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Set xmlNodeList = xmlTempResponseNode.selectNodes("//QUOTATION")
    
    'For each Quotation found, update with new cost
    For Each xmlNode In xmlNodeList
        Set xmlTempNode = m_objXmlAssist.GetMandatoryNode(xmlNode, "TOTALQUOTATIONCOST")
        dblCurrentCost = objConvAssist.CSafeDbl(xmlTempNode.Text)
        dblCurrentCost = dblCurrentCost - dblOldCost + dblNewCost
        xmlTempNode.Text = CStr(dblCurrentCost)
        Set xmlTempResponseNode = IQuotationBO_Update(xmlNode)
    Next
    
    Set IQuotationBO_UpdateLinkedQuotations = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateLinkedComponentsExit:
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateLinkedComponentsVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_UpdateLinkedQuotations = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLinkedComponentsExit

End Function

Private Function IQuotationBO_Validate(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IQuotationBO_Validate"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIQuotationDO As IQuotationDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

'    If objContext Is Nothing Then
'        Set objIQuotationDO = New QuotationDO
'    Else
'        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
'    End If
'
'    ' Get node to add derived data to and apply validation to
'    If vxmlRequest.nodeName = "UPDATE" Or vxmlRequest.nodeName = "CREATE" Then
'        Set xmlRequestTypeNode = vxmlRequest
'    ElseIf vxmlRequest.getElementsByTagName("UPDATE").length > 0 Then
'        Set xmlRequestTypeNode = vxmlRequest.getElementsByTagName("UPDATE").Item(0)
'    ElseIf vxmlRequest.getElementsByTagName("CREATE").length > 0 Then
'        Set xmlRequestTypeNode = vxmlRequest.getElementsByTagName("CREATE").Item(0)
'    End If
'
'    If Not xmlRequestTypeNode Is Nothing Then
'        Set xmlDataNode = xmlRequestTypeNode.firstChild
'
'        ' Add in derived data
'        If Not xmlDataNode Is Nothing Then
'            objIQuotationDO.AddDerivedData xmlDataNode
'        End If
'    End If
    
' TODO ************************************************************************
'   add hard coded validation here, & raise error if it fails
' *****************************************************************************

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IQuotationBO_Validate = xmlResponseElem.cloneNode(True)
    
ValidateExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_Validate = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Private Function IQuotationBO_ValidateQuotationForReinstatement(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Validates the sub quotes on the quotation
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'               <QUOTATIONNUMBER>Value</QUOTATIONNUMBER>
'               <ISMEMBEROFATAFF>Value</ISMEMBEROFATAFF>
'               <ISSELFEMPLOYED>Value</ISSELFEMPLOYED>
'               <LOANSTOBEPORTED>
'                   <MORTGAGELOAN>
'                       ...
'                       ...
'                   </MORTGAGELOAN>
'               </LOANSTOBEPORTED>
'               <LTV>Value</LTV>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ValidateQuotationForReinstatementVbErr
    
    Const strFunctionName As String = "IQuotationBO_ValidateQuotationForReinstatement"


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
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempResponseElem = IQuotationBO_GetValidatedQuotationDetails(vXMLRequest)
    objErrAssist.CheckResponse xmlTempResponseElem, xmlResponseElem, True
    
    Dim strTagName As String
    strTagName = "VALIDMORTGAGESUBQUOTE"
    
    Set xmlElement = xmlTempResponseElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strTagName & " tag not found"
    End If
    
    xmlResponseElem.appendChild xmlElement
    
    strTagName = "VALIDBCSUBQUOTE"
    Set xmlElement = xmlTempResponseElem.getElementsByTagName(strTagName).Item(0)
    'SR : 05/04/00 - Add the BCSubQuote details, only if they exist
    If Not xmlElement Is Nothing Then xmlResponseElem.appendChild xmlElement
    
    strTagName = "VALIDPPSUBQUOTE"
    Set xmlElement = xmlTempResponseElem.getElementsByTagName(strTagName).Item(0)
    If Not xmlElement Is Nothing Then xmlResponseElem.appendChild xmlElement
        
    Set IQuotationBO_ValidateQuotationForReinstatement = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
ValidateQuotationForReinstatementExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlElement = Nothing

    Exit Function

ValidateQuotationForReinstatementVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_ValidateQuotationForReinstatement = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateQuotationForReinstatementExit

End Function

Private Sub AddQuoteSummaryDefaults(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Add default values for Language and Country of Origin
'
' pass:         vxmlRequest     xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddQuoteSummaryDefaultsVbErr
    
    Const strFunctionName As String = "AddQuoteSummaryDefaults"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim objComboDO As ComboDO
    Dim xmlElement As IXMLDOMElement
    Dim strTagName As String
        
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' Add the country of origin  if not already set
    Dim strCountryOfOrigin As String
    
    strTagName = "COUNTRYOFORIGIN"
    Set xmlElement = vXMLRequest.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vXMLRequest.ownerDocument.createElement(strTagName)
        vXMLRequest.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        strCountryOfOrigin = objComboDO.GetFirstComboValueId("CountryOfOrigin", "UK")
        xmlElement.Text = strCountryOfOrigin
    End If
   
    ' Add the language  if not already set
    Dim strLanguage As String
    
    strTagName = "LANGUAGE"
    Set xmlElement = vXMLRequest.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vXMLRequest.ownerDocument.createElement(strTagName)
        vXMLRequest.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        strLanguage = objComboDO.GetFirstComboValueId("MortgageProductLanguage", "Eng")
        xmlElement.Text = strLanguage
    End If
        
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlElement = Nothing

    Exit Sub

AddQuoteSummaryDefaultsVbErr:
    
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set xmlElement = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Public Function GetAcceptedBCSubQuoteNumber(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get summary details of the active quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <QUOTATION>
'           <ACCEPTEDQUOTENUMBER>Value</ACCEPTEDQUOTENUMBER>
'       </QUOTATION>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetAcceptedBCSubQuoteNumberVbErr

    Const strFunctionName As String = "GetAcceptedBCSubQuoteNumber"

    
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
    Set xmlTempResponseNode = IQuotationBO_GetAcceptedBCSubQuoteNumber(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAcceptedBCSubQuoteNumber = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetAcceptedBCSubQuoteNumberExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAcceptedBCSubQuoteNumberVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetAcceptedBCSubQuoteNumber = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetAcceptedBCSubQuoteNumberExit

End Function

Public Function GetAllQuotationsForApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get summary details of the active quotation
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           </QUOTATION>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <QUOTATION>
'
'       </QUOTATION>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetAllQuotationsForApplicationVbErr

    Const strFunctionName As String = "GetAllQuotationsForApplication"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IQuotationBO_GetAllQuotationsForApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetAllQuotationsForApplication = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetAllQuotationsForApplicationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetAllQuotationsForApplicationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetAllQuotationsForApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetAllQuotationsForApplicationExit

End Function
Private Function IQuotationBO_GetAcceptedBCSubQuoteNumber(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get the BC Subquote number for the Accepted quote number
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'
'   Input XML
'
'   <REQUEST>
'       <SEARCH>
'           <QUOTATION>
'               <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'               <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'           </QUOTATION>
'       </SEARCH>
'   </REQUEST>
'
'   Output XML
'
'   <RESPONSE>
'       <QUOTATION>
'           <ACCEPTEDQUOTENUMBER>Value</ACCEPTEDQUOTENUMBER>
'       </QUOTATION>
'   </RESPONSE>
'------------------------------------------------------------------------------------------
On Error GoTo GetAcceptedBCSubQuoteNumberVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetAcceptedBCSubQuoteNumber"


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
    Dim objIQuotationDO As IQuotationDO
'    Dim xmlRequestClone As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = objContext.CreateInstance(App.Title & ".QuotationDO")
    End If
    
'    ' Clone the input so that we can add default country and language
'    Set xmlRequestClone = vxmlRequest.cloneNode(True)
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    
'    ' Add in the default parameters for Country of origin and Language
'    AddQuoteSummaryDefaults xmlRequestNode
    
    Set xmlDataNode = objIQuotationDO.GetAcceptedBCSubQuoteNumber(xmlRequestNode)
    If Not xmlDataNode Is Nothing Then
        xmlResponseElem.appendChild xmlDataNode
    
    End If
        
    Set IQuotationBO_GetAcceptedBCSubQuoteNumber = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
GetAcceptedBCSubQuoteNumberExit:
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
'    Set xmlRequestClone = Nothing

    Exit Function

GetAcceptedBCSubQuoteNumberVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    If objErrAssist.GetOmigaErrorNumber(Err.number) = oeRecordNotfound Then
            Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetAcceptedBCSubQuoteNumber = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetAcceptedBCSubQuoteNumberExit

End Function

Public Function GetQuotationLoanSummary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationLoanSummaryVbErr

    Const strFunctionName As String = "GetQuotationLoanSummary"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IQuotationBO_GetQuotationLoanSummary(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetQuotationLoanSummary = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetQuotationLoanSummaryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetQuotationLoanSummaryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetQuotationLoanSummary = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetQuotationLoanSummaryExit

End Function

Private Function IQuotationBO_GetQuotationLoanSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationLoanSummaryVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetQuotationLoanSummary"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIQuotationDO = m_objContext.CreateInstance(App.Title & ".QuotationDO")
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIQuotationDO.GetQuotationLoanSummary(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IQuotationBO_GetQuotationLoanSummary = xmlResponseElem
    
    m_objContext.SetComplete
    
GetQuotationLoanSummaryExit:
    
    Set xmlRequestNode = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetQuotationLoanSummaryVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetQuotationLoanSummary = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetQuotationLoanSummaryExit

End Function

Public Function GetQuotationDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo GetQuotationDetailsVbErr

    Const strFunctionName As String = "GetQuotationDetails"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IQuotationBO_GetQuotationDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetQuotationDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetQuotationDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

GetQuotationDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    GetQuotationDetails = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetQuotationDetailsExit

End Function

Private Function IQuotationBO_GetQuotationDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationDetailsVbErr
    
    Const strFunctionName As String = "IQuotationBO_GetQuotationDetails"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIQuotationDO As IQuotationDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIQuotationDO = m_objContext.CreateInstance(App.Title & ".QuotationDO")
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIQuotationDO.GetQuotationDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set IQuotationBO_GetQuotationDetails = xmlResponseElem

GetQuotationDetailsExit:

    Set xmlRequestNode = Nothing
    Set objIQuotationDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetQuotationDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_GetQuotationDetails = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetQuotationDetailsExit

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

'BMIDS624 GHun 22/10/2003
Public Function HaveRatesChanged(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks if rates have changed since the quote was last modelled
' pass:         vxmlRequest  xml Request data containing application number to check
' return:       xml Response Node
'   Input XML:
'
'   <REQUEST>
'       <APPLICATION>
'           <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'       </APPLICATION>
'   </REQUEST>
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "HaveRatesChanged"

    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
On Error GoTo HaveRatesChangedVbErr
    
    ' Create default response block
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Set xmlResponseElem = xmlIn.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IQuotationBO_HaveRatesChanged(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    HaveRatesChanged = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

HaveRatesChangedExit:

    ' clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

HaveRatesChangedVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    HaveRatesChanged = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume HaveRatesChangedExit

End Function
'BMIDS624 End

'BMIDS624 GHun 22/10/2003
Private Function IQuotationBO_HaveRatesChanged(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Checks if rates have changed since the quote was last modelled
' pass:         vxmlRequest  xml Request data containing application number to check
' return:       xml Response Node
'   Input XML:
'
'   <REQUEST>
'       <QUOTATION>
'           <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'       </QUOTATION>
'   </REQUEST>
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName As String = "IQuotationBO_HaveRatesChanged"

    Dim xmlDoc          As FreeThreadedDOMDocument40
    Dim xmlRequestNode  As IXMLDOMNode
    Dim objIQuotationDO As IQuotationDO
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDataNode     As IXMLDOMNode
    
On Error GoTo HaveRatesChangedVbErr
         
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If vXMLRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vXMLRequest
    Else
        Set xmlRequestNode = vXMLRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If

    If m_objContext Is Nothing Then
        Set objIQuotationDO = New QuotationDO
    Else
        Set objIQuotationDO = m_objContext.CreateInstance(App.Title & ".QuotationDO")
    End If

    Set xmlDataNode = objIQuotationDO.HaveRatesChanged(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
           
    Set IQuotationBO_HaveRatesChanged = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
HaveRatesChangedExit:
    
    Set xmlDoc = Nothing
    Set xmlRequestNode = Nothing
    Set objIQuotationDO = Nothing
    Set xmlResponseElem = Nothing
    Set xmlDataNode = Nothing

    Exit Function

HaveRatesChangedVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource cstrFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If

    Set IQuotationBO_HaveRatesChanged = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume HaveRatesChangedExit

End Function
'BMIDS624 End
