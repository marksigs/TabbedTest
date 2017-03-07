VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "BuildingsAndContentsSubQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BuildingsAndContentsSubQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for BuildingsAndContentsSubQuote Business Object which requires transactions
'
'Dependencies:  List any other dependent components
'               e.g. BuildingsAndContentsSubQuoteDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'PSC    01/10/99    Added new error processing
'PSC    27/10/99    Remove after image processing on Update
'PSC    30/11/99    Amend to use DOMDocuments instead of strings
'PSC    14/12/99    Amend to not use the default interface so that scripting cannot access
'                   methods but internal Omiga components can
'PSC    16/12/99    Amend interfaces
'APS    30/05/00    SYS0767 - Avoid duplicate key error by always creating the latest + 1 record
'DJP    07/09/00    SYS1460, don't default to 1 in CreateFirstSubQuoteXML
'SR     04/07/01    SYS2433 - modified method 'IBuildingsAndContentsSubQuoteTxBO_DeleteAllValuables'
'---------------------------------------------------------------------------------------------------

Option Explicit

Implements IBuildingsAndContentsSubQuoteTxBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
Private Const cstrTABLE_NAME = "TABLENAME"

'BMIDS History:
'
'Prog   Date        Description
'GD     21/08/02    BMIDS00312  - Keep TOTALBCMONTHLYCOST


Private Function CreateNewSubQuoteXML(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement

' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the new sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteXMLVbErr

    Const strFunctionName As String = "CreateNewSubQuoteXML"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteNoList As IXMLDOMElement
    Dim xmlDetailsElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    
    Dim strApplicationNo As String
    Dim ApplicationFFNo As String
    Dim strContext As String
    Dim strBCSubQuoteNumber As String
    Dim strTagName As String
    Dim xmlResponseElem As IXMLDOMElement
    
    'Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    'xmlOut.appendChild xmlResponseElem
    'xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    ApplicationFFNo = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    strTagName = "CONTEXT"
    strContext = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    strTagName = "BCSUBQUOTENUMBER"
    strBCSubQuoteNumber = objXmlAssist.GetTagValue(vXMLRequest, strTagName)

    If Len(strApplicationNo) = 0 Or _
       Len(ApplicationFFNo) = 0 Or _
       Len(strContext) = 0 Or _
       Len(strBCSubQuoteNumber) = 0 Then
        Dim strDetails As String
        strDetails = "APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & " APPLICATIONFACTFINDNUMBER = " & ApplicationFFNo
        strDetails = strDetails & " CONTEXT = " & strContext
        strDetails = strDetails & " BCSUBQUOTENUMBER = " & strBCSubQuoteNumber
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If
       
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
   
    Set xmlDataNode = objIBuildingsAndContentsSubQuoteDO.GetSubQuoteDetails(vXMLRequest)
    xmlOut.appendChild xmlDataNode
    
    Set xmlSubQuoteElem = xmlOut.getElementsByTagName("BUILDINGSANDCONTENTSSUBQUOTE").Item(0)
    Set xmlDetailsElem = xmlOut.getElementsByTagName("BUILDINGSANDCONTENTSDETAILS").Item(0)
    
    'BUILDINGSANDCONTENTSSUBQUOTE
    strTagName = "BCSUBQUOTENUMBER"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
  
    strTagName = "DATEANDTIMEGENERATED"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If
    
    strTagName = "QUOTATIONTYPE"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        If StrComp(strContext, "Quick Quote", vbTextCompare) = 0 Then
            xmlElement.Text = "1"
        Else
            xmlElement.Text = "2"
        End If
    End If
    
    strTagName = "TOTALBCMONTHLYCOST"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    'GD BMIDS00312 21/08/2002 - Keep TOTALBCMONTHLYCOST
    'If Not xmlElement Is Nothing Then
     '   xmlElement.Text = ""
    'End If
    
    'BUILDINGSANDCONTENTSDETAILS
    strTagName = "BCSUBQUOTENUMBER"
    Set xmlElement = xmlDetailsElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
  
    strTagName = "IPT"
    Set xmlElement = xmlDetailsElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
    
    strTagName = "PREMIUM"
    Set xmlElement = xmlDetailsElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
        
    Set CreateNewSubQuoteXML = xmlOut.documentElement
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlDetailsElem = Nothing
    Set xmlSubQuoteNoList = Nothing
    Set xmlElement = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing

    Exit Function

CreateNewSubQuoteXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlDetailsElem = Nothing
    Set xmlSubQuoteNoList = Nothing
    Set xmlElement = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function CreateFirstSubQuoteXML(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for the first sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       First sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteXMLVbErr

    Const strFunctionName As String = "CreateFirstSubQuoteXML"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteDetailsElem As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlDetailsElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim strApplicationNo As String
    Dim ApplicationFFNo As String
    Dim strContext As String
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    ApplicationFFNo = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    strTagName = "CONTEXT"
    strContext = objXmlAssist.GetTagValue(vXMLRequest, strTagName)
    
    If Len(strApplicationNo) = 0 Or _
       Len(ApplicationFFNo) = 0 Or Len(strContext) = 0 Then
        Dim strDetails As String
        strDetails = "APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & "APPLICATIONFACTFINDNUMBER = " & ApplicationFFNo
        strDetails = strDetails & "CONTEXT = " & strContext
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If
        
    Set xmlSubQuoteDetailsElem = xmlOut.createElement("BCSUBQUOTEDETAILS")
    xmlOut.appendChild xmlSubQuoteDetailsElem
    Set xmlSubQuoteElem = xmlOut.createElement("BUILDINGSANDCONTENTSSUBQUOTE")
    xmlSubQuoteDetailsElem.appendChild xmlSubQuoteElem
    Set xmlDetailsElem = xmlOut.createElement("BUILDINGSANDCONTENTSDETAILS")
    xmlSubQuoteDetailsElem.appendChild xmlDetailsElem

    ' BUILDINGSANDCONTENTSSUBQUOTE
    strTagName = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = strApplicationNo
    xmlSubQuoteElem.appendChild xmlElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = ApplicationFFNo
    xmlSubQuoteElem.appendChild xmlElement
    
    strTagName = "BCSUBQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    ' APS SYS0767 - Avoid duplicate key error by always creating the latest + 1 record
    xmlElement.Text = ""
    xmlSubQuoteElem.appendChild xmlElement
    
    strTagName = "DATEANDTIMEGENERATED"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    xmlSubQuoteElem.appendChild xmlElement
   
    strTagName = "QUOTATIONTYPE"
    Set xmlElement = xmlOut.createElement(strTagName)
    
    If StrComp(strContext, "Quick Quote", vbTextCompare) = 0 Then
        xmlElement.Text = "1"
    Else
        xmlElement.Text = "2"
    End If
    xmlSubQuoteElem.appendChild xmlElement
      
    'BUILDINGSANDCONTENTSDETAILS
    strTagName = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = strApplicationNo
    xmlDetailsElem.appendChild xmlElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = ApplicationFFNo
    xmlDetailsElem.appendChild xmlElement
    
    strTagName = "BCSUBQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    
    ' DJP SYS1460, don't default to 1.
    xmlElement.Text = ""
    xmlDetailsElem.appendChild xmlElement
    
    Set CreateFirstSubQuoteXML = xmlSubQuoteDetailsElem
    
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlElement = Nothing

    Exit Function

CreateFirstSubQuoteXMLVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlElement = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub CreateNewValuablesList(ByVal vxmlOldValuables As IXMLDOMElement, vxmlOldSubquote As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  copies a valuables list to a new subquotenumber for the same application/factfind
'               number ( taken from OldSubquote passed in) and saves them
' pass:         vxmlOldValuables  xml Request data stream containing data to be copied
'               vxmlOldSubQuote   xml containing the BCSubQuote number to copy to.
' return:       valuables list xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewValuablesListVbErr

    Const strFunctionName As String = "CreateNewValuablesList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNewElement As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlListNode As IXMLDOMNode
    Dim xmlValuablesNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim strBCSubQuoteNumber As String
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
   
    Set xmlDataNode = vxmlOldSubquote.selectSingleNode(".//BCSUBQUOTENUMBER")
    strBCSubQuoteNumber = xmlDataNode.Text
    Set xmlDataNode = vxmlOldSubquote.selectSingleNode(".//APPLICATIONNUMBER")
    strApplicationNumber = xmlDataNode.Text
    Set xmlDataNode = vxmlOldSubquote.selectSingleNode(".//APPLICATIONFACTFINDNUMBER")
    strApplicationFFNumber = xmlDataNode.Text
    
    Set xmlNodeList = vxmlOldValuables.selectNodes(".//VALUABLESOVERLIMIT")
    'traverse the list and add each valuable to the new list
    For Each xmlElement In xmlNodeList
    
        Set xmlNewElement = xmlOut.createElement("VALUABLESOVERLIMIT")
        Set xmlValuablesNode = xmlOut.appendChild(xmlNewElement)
        Set xmlNewElement = xmlOut.createElement("APPLICATIONNUMBER")
        xmlNewElement.Text = strApplicationNumber
        xmlValuablesNode.appendChild xmlNewElement
        Set xmlNewElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNewElement.Text = strApplicationFFNumber
        xmlValuablesNode.appendChild xmlNewElement
        Set xmlNewElement = xmlOut.createElement("BCSUBQUOTENUMBER")
        xmlNewElement.Text = strBCSubQuoteNumber
        xmlValuablesNode.appendChild xmlNewElement
        Set xmlNewElement = xmlOut.createElement("VALUABLESLIMITSEQUENCENUMBER")
        xmlNewElement.Text = objXmlAssist.GetTagValue(xmlElement, "VALUABLESLIMITSEQUENCENUMBER")
        xmlValuablesNode.appendChild xmlNewElement
        Set xmlNewElement = xmlOut.createElement("DESCRIPTION")
        xmlNewElement.Text = objXmlAssist.GetTagValue(xmlElement, "DESCRIPTION")
        xmlValuablesNode.appendChild xmlNewElement
        Set xmlNewElement = xmlOut.createElement("INSURABLEVALUE")
        xmlNewElement.Text = objXmlAssist.GetTagValue(xmlElement, "INSURABLEVALUE")
        xmlValuablesNode.appendChild xmlNewElement
       
        objIBuildingsAndContentsSubQuoteDO.CreateValuablesOverLimit xmlValuablesNode.cloneNode(True)
        Set xmlOut = Nothing
    Next
   
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlNewElement = Nothing
    Set xmlDataNode = Nothing
    Set xmlListNode = Nothing
    Set xmlValuablesNode = Nothing
    Set xmlNodeList = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing

    Exit Sub

CreateNewValuablesListVbErr:
    
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlNewElement = Nothing
    Set xmlDataNode = Nothing
    Set xmlListNode = Nothing
    Set xmlValuablesNode = Nothing
    Set xmlNodeList = Nothing
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IBuildingsAndContentsSubQuoteTxBO_CreateBCNotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateBCNotesVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_CreateBCNotes"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BUILDINGSANDCONTENTSNOTES" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BUILDINGSANDCONTENTSNOTES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BUILDINGSANDCONTENTSNOTES tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIBCSubQuoteDO.CreateBCNotes xmlElement
    Else
        ' For each relevant node call the data object create
        For Each xmlElement In xmlNodeList
            objIBCSubQuoteDO.CreateBCNotes xmlElement
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_CreateBCNotes = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateBCNotesExit:

    Set objIBCSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateBCNotesVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_CreateBCNotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateBCNotesExit


End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_CreateFirstSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_CreateFirstSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubQuote As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement


    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILS" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BCSUBQUOTEDETAILS tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    Set xmlTempElem = xmlOut.createElement("BCSUBQUOTEDETAILS")
    Set xmlTempNode = xmlResponseElem.appendChild(xmlTempElem)
    Dim strBCSubQuoteNumber As String
    
    If Not xmlElement Is Nothing Then
        Set xmlSubQuote = CreateFirstSubQuoteXML(xmlElement)
        Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlSubQuote)
        'extract the bcsubquote number only and return it
        strBCSubQuoteNumber = objXmlAssist.GetTagValue(xmlTempResponseNode, "BCSUBQUOTENUMBER")
        Set xmlTempElem = xmlOut.createElement("BCSUBQUOTENUMBER")
        xmlTempElem.Text = strBCSubQuoteNumber
        xmlTempNode.appendChild xmlTempElem

    Else
        ' For each relevant node call the data object create
        Dim iCount As Integer
        iCount = 0
        For Each xmlElement In xmlNodeList
            Set xmlSubQuote = CreateFirstSubQuoteXML(xmlElement)
            Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlSubQuote)
            If iCount = 0 Then
                'extract the bcsubquote number only and return it
                strBCSubQuoteNumber = objXmlAssist.GetTagValue(xmlTempResponseNode, "BCSUBQUOTENUMBER")
                Set xmlTempElem = xmlOut.createElement("BCSUBQUOTENUMBER")
                xmlTempElem.Text = strBCSubQuoteNumber
                xmlTempNode.appendChild xmlTempElem
            End If
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_CreateFirstSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFirstExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuote = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOut = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempElem = Nothing
    
    Exit Function

CreateFirstVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_CreateFirstSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFirstExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_CreateNewSubQuote(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_CreateNewSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubQuote As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILS" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BCSUBQUOTEDETAILS tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlForValuablesNode As IXMLDOMNode
    Dim xmlForNotesNode As IXMLDOMNode
    Dim strBCSubQuoteNumber As String
    
    Set xmlTempElem = xmlOut.createElement("BCSUBQUOTEDETAILS")
    Set xmlTempNode = xmlResponseElem.appendChild(xmlTempElem)
    If Not xmlElement Is Nothing Then
        'we need to remember the current bcsubquotenumber to get the valuables and notes
        Set xmlForValuablesNode = xmlElement.cloneNode(True)
        Set xmlForNotesNode = xmlElement.cloneNode(True)
        
        Set xmlSubQuote = CreateNewSubQuoteXML(xmlElement)
        Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlSubQuote)
        'extract the bcsubquote number only and return it
        strBCSubQuoteNumber = objXmlAssist.GetTagValue(xmlTempResponseNode, "BCSUBQUOTENUMBER")
        Set xmlTempElem = xmlOut.createElement("BCSUBQUOTENUMBER")
        xmlTempElem.Text = strBCSubQuoteNumber
        xmlTempNode.appendChild xmlTempElem
        
        objXmlAssist.ChangeNodeName xmlForValuablesNode, "BCSUBQUOTEDETAILS", "VALUABLESOVERLIMIT"
        Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.FindValuablesOverLimitList(xmlForValuablesNode)
        If Not xmlTempResponseNode Is Nothing Then
            ' If any valuables have been found then we need to change the subquote
            ' number to the new subquote and remove the valuableslimit sequence
            ' number then create them
            CreateNewValuablesList xmlTempResponseNode, xmlSubQuote
        End If
        
        objXmlAssist.ChangeNodeName xmlForNotesNode, "BCSUBQUOTEDETAILS", "BUILDINGSANDCONTENTSNOTES"
        Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.GetBCNotes(xmlForNotesNode)
        If Not xmlTempResponseNode Is Nothing Then
            ' If any notes have been found then we need to change the subquote
            ' number to the new subquote and then create them
            Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//BCSUBQUOTENUMBER")
            xmlTempNode.Text = strBCSubQuoteNumber
            objIBuildingsAndContentsSubQuoteDO.CreateBCNotes xmlTempResponseNode
        End If
        
    Else
        ' For each relevant node call the data object create
        Dim iCount As Integer
        iCount = 0
        For Each xmlElement In xmlNodeList
            'we need to remember the current bcsubquotenumber to get the valuables.
            Set xmlForValuablesNode = xmlElement.cloneNode(True)
            Set xmlForNotesNode = xmlElement.cloneNode(True)
            
            Set xmlSubQuote = CreateNewSubQuoteXML(xmlElement)
            Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlSubQuote)
            If iCount = 0 Then
                'extract the bcsubquote number only and return it
                strBCSubQuoteNumber = objXmlAssist.GetTagValue(xmlTempResponseNode, "BCSUBQUOTENUMBER")
                Set xmlTempElem = xmlOut.createElement("BCSUBQUOTENUMBER")
                xmlTempElem.Text = strBCSubQuoteNumber
                xmlTempNode.appendChild xmlTempElem
            End If
            
            objXmlAssist.ChangeNodeName xmlForValuablesNode, "BCSUBQUOTEDETAILS", "VALUABLESOVERLIMIT"
            Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.FindValuablesOverLimitList(xmlForValuablesNode)
            ' need to change the subquote number to the new subquote and remove
            ' the valuableslimit sequence number then create them
            If Not xmlTempResponseNode Is Nothing Then
                CreateNewValuablesList xmlTempResponseNode, xmlSubQuote
            End If
            
            'objXmlAssist.ChangeNodeName xmlForNotesNode, "BCSUBQUOTEDETAILS", "BUILDINGSANDCONTENTSNOTES"
            'Set xmlTempResponseNode = objIBuildingsAndContentsSubQuoteDO.GetBCNotes(xmlForNotesNode)
            'If Not xmlTempResponseNode Is Nothing Then
                ' If any notes have been found then we need to change the subquote
                ' number to the new subquote and then create them
            '    Set xmlTempNode = xmlTempResponseNode.selectSingleNode(".//BCSUBQUOTENUMBER")
            '    xmlTempNode.Text = strBCSubQuoteNumber
            '    objIBuildingsAndContentsSubQuoteDO.CreateBCNotes xmlTempResponseNode
            'End If
            
            iCount = iCount + 1
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_CreateNewSubQuote = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateNewExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuote = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlForValuablesNode = Nothing
    Set xmlForNotesNode = Nothing
    
    Exit Function

CreateNewVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlTempResponseNode
        Resume Next
    End If

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_CreateNewSubQuote = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_CreateSubQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_CreateSubQuoteDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILS" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BCSUBQUOTEDETAILS tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        Set xmlResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlElement)
    Else
        ' For each relevant node call the data object create
        For Each xmlElement In xmlNodeList
            Set xmlResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateSubQuoteDetails(xmlElement)
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_CreateSubQuoteDetails = xmlResponseNode.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_CreateSubQuoteDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_CreateValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_CreateValuablesOverLimit"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("VALUABLESOVERLIMIT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VALUABLESOVERLIMIT tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If Not xmlElement Is Nothing Then
        Set xmlResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateValuablesOverLimit(xmlElement)
        xmlResponseElem.appendChild xmlResponseNode
    Else
        ' For each relevant node call the data object create
        For Each xmlElement In xmlNodeList
            Set xmlResponseNode = objIBuildingsAndContentsSubQuoteDO.CreateValuablesOverLimit(xmlElement)
            xmlResponseElem.appendChild xmlResponseNode
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_CreateValuablesOverLimit = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlResponseNode = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseNode
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_CreateValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit


End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_DeleteValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_DeleteValuablesOverLimit"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "VALUABLESOVERLIMIT" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("VALUABLESOVERLIMIT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VALUABLESOVERLIMIT tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIBuildingsAndContentsSubQuoteDO.DeleteValuablesOverLimit xmlElement
    Else
        ' For each relevant node call the data object Delete
        For Each xmlElement In xmlNodeList
            objIBuildingsAndContentsSubQuoteDO.DeleteValuablesOverLimit xmlElement
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_DeleteValuablesOverLimit = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_DeleteValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_DeleteAllValuables(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllValuablesVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_DeleteAllValuables"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
        
    Dim objXmlAssist As New XMLAssist
    
    Dim xmlValuableOverLimit As IXMLDOMNode
    'SR 04/07/01 : SYS2433 - Check the name of the top-level node and set the search pattern accordingly
    If vXMLRequest.nodeName = "VALUABLESOVERLIMITLIST" Then
        Set xmlValuableOverLimit = objXmlAssist.GetNode(vXMLRequest, "./VALUABLESOVERLIMIT")
    Else
        Set xmlValuableOverLimit = objXmlAssist.GetNode(vXMLRequest, ".//VALUABLESOVERLIMITLIST/VALUABLESOVERLIMIT")
    End If
        
    Dim xmlTable As New FreeThreadedDOMDocument40
    xmlTable.validateOnParse = False
    xmlTable.setProperty "NewParser", True
    xmlTable.appendChild xmlValuableOverLimit.cloneNode(True)
    
    Dim xmlValSeqNoNode As IXMLDOMNode
    Set xmlValSeqNoNode = objXmlAssist.GetNode(xmlTable.documentElement, "//VALUABLESOVERLIMIT/VALUABLESLIMITSEQUENCENUMBER")
    xmlValSeqNoNode.parentNode.removeChild xmlValSeqNoNode
    
    ' Pass through to the data object
    objIBuildingsAndContentsSubQuoteDO.DeleteAllValuables xmlTable.documentElement
    
    Set IBuildingsAndContentsSubQuoteTxBO_DeleteAllValuables = xmlResponseElem.cloneNode(True)
    
    Set objXmlAssist = Nothing
    Set xmlTable = Nothing
    Set xmlValSeqNoNode = Nothing
    Set xmlValuableOverLimit = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteAllValuablesExit:
    
    Set objXmlAssist = Nothing
    Set xmlTable = Nothing
    Set xmlValSeqNoNode = Nothing
    Set xmlValuableOverLimit = Nothing
    
    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function
    
DeleteAllValuablesVbErr:
    
    If Err.Number = omiga4NoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_DeleteAllValuables = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteAllValuablesExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_DeleteAll(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_DeleteAll"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

' TODO ************************************************************************
'   Remove oeNotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented

' TODO ************************************************************************
'   replace BuildingsAndContentsSubQuoteDO reference with actual DO required
' *****************************************************************************
    Dim objIBuildingsAndContentsSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIBuildingsAndContentsSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBuildingsAndContentsSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
        
    ' Pass through to the data object
    objIBuildingsAndContentsSubQuoteDO.DeleteAll vXMLRequest
    
    Set IBuildingsAndContentsSubQuoteTxBO_DeleteAll = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteAllExit:

    Set objIBuildingsAndContentsSubQuoteDO = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function
    
DeleteAllVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_DeleteAll = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteAllExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_UpdateBCNotes(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_UpdateBCNotes"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BUILDINGSANDCONTENTSNOTES" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BUILDINGSANDCONTENTSNOTES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BUILDINGSANDCONTENTSNOTES tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIBCSubQuoteDO.UpdateBCNotes xmlElement
    Else
        ' For each relevant node call the data object Update
        For Each xmlElement In xmlNodeList
            objIBCSubQuoteDO.UpdateBCNotes xmlElement
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_UpdateBCNotes = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIBCSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_UpdateBCNotes = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_UpdateSubQuoteDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_UpdateSubQuoteDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "BCSUBQUOTEDETAILSNOTES" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("BCSUBQUOTEDETAILSNOTES")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "BCSUBQUOTEDETAILSNOTES tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIBCSubQuoteDO.UpdateSubQuoteDetails xmlElement
    Else
        ' For each relevant node call the data object Update
        For Each xmlElement In xmlNodeList
            objIBCSubQuoteDO.UpdateSubQuoteDetails xmlElement
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_UpdateSubQuoteDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIBCSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_UpdateSubQuoteDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IBuildingsAndContentsSubQuoteTxBO_UpdateValuablesOverLimit(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteTxBO_UpdateValuablesOverLimit"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIBCSubQuoteDO As IBuildingsAndContentsSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vXMLRequest.nodeName = "VALUABLESOVERLIMIT" Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlNodeList = vXMLRequest.getElementsByTagName("VALUABLESOVERLIMIT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VALUABLESOVERLIMIT tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIBCSubQuoteDO = New BuildingsAndContentsSubQuoteDO
    Else
        Set objIBCSubQuoteDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsSubQuoteDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIBCSubQuoteDO.UpdateValuablesOverLimit xmlElement
    Else
        ' For each relevant node call the data object Update
        For Each xmlElement In xmlNodeList
            objIBCSubQuoteDO.UpdateValuablesOverLimit xmlElement
        Next
    End If
    
    Set IBuildingsAndContentsSubQuoteTxBO_UpdateValuablesOverLimit = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIBCSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsSubQuoteTxBO_UpdateValuablesOverLimit = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit


End Function

