VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PaymentProtectionSubQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProtectionSubQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Payment Protection Sub Quote Business Object which requires
'               transactions

'Dependencies:  PaymentProtectionSubQuoteDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    03/11/99    Created
'APS    30/05/00    SYS0767 - Avoid duplicate key error by always creating the latest + 1 record
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************

Private Const cstrSUBQUOTE_NAME = "PAYMENTPROTECTIONSUBQUOTE"
 
'BMIDS History:
'
'Prog   Date        Description
'GD     21/08/02    BMIDS00312  - Keep TOTALPPMONTHLYCOST

Public Function CreateFirstSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateFirstSubQuoteVbErr

    Const strFunctionName As String = "CreateFirstSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
    Dim xmlListElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    ' If we have more than one in the list then we need a list of keys
    If xmlNodeList.length > 1 Then
        Set xmlListElem = xmlOut.createElement(cstrSUBQUOTE_NAME & "LIST")
        xmlOut.appendChild xmlListElem
    End If
   
    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
    
    Dim intNodeIndex As Integer
    Dim strSubQuoteXML As String
    Dim strData As String
    
    ' For each relevant node call the data object create and add the key into the
    ' output xml
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        strSubQuoteXML = CreateFirstSubQuoteXML(objElement.xml)
        strData = objPPSubQuoteDO.CreateSubQuoteDetails(strSubQuoteXML)
        Set xmlDoc = objXmlAssist.load(strData, TypeName(Me), strFunctionName)
        
        If xmlListElem Is Nothing Then
            xmlOut.appendChild xmlDoc.documentElement
        Else
            xmlListElem.appendChild xmlDoc.documentElement
        End If
    Next
 
    CreateFirstSubQuote = xmlOut.xml
 
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

CreateFirstSubQuoteVbErr:
   
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function CreateNewSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteVbErr

    Const strFunctionName As String = "CreateNewSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
    Dim xmlListElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    ' If we have more than one in the list then we need a list of keys
    If xmlNodeList.length > 1 Then
        Set xmlListElem = xmlOut.createElement(cstrSUBQUOTE_NAME & "LIST")
        xmlOut.appendChild xmlListElem
    End If
   
    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
    
    Dim intNodeIndex As Integer
    Dim strSubQuoteXML As String
    Dim strData As String
    
    ' For each relevant node call the data object create and add the key into the
    ' output xml
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        strSubQuoteXML = CreateNewSubQuoteXML(objElement.xml)
        strData = objPPSubQuoteDO.CreateSubQuoteDetails(strSubQuoteXML)
        Set xmlDoc = objXmlAssist.load(strData, TypeName(Me), strFunctionName)
        
        If xmlListElem Is Nothing Then
            xmlOut.appendChild xmlDoc.documentElement
        Else
            xmlListElem.appendChild xmlDoc.documentElement
        End If
    Next
 
    CreateNewSubQuote = xmlOut.xml
 
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

CreateNewSubQuoteVbErr:
   
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Sub SaveDetails(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing update data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo SaveDetailsVbErr

    Const strFunctionName As String = "SaveDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer
    
    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
    
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        CalculateAndAddTotals objElement
        objPPSubQuoteDO.UpdateSubQuoteDetails objElement.xml
    Next
 
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
    Exit Sub

SaveDetailsVbErr:
   
    Set objPPSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function CreateFirstSubQuoteXML(vstrXMLRequest) As String
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
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Dim strApplicationNo As String
    Dim ApplicationFFNo As String
    Dim strContext As String
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    ApplicationFFNo = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    strTagName = "CONTEXT"
    strContext = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
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
        
    Set xmlSubQuoteElem = xmlOut.createElement(cstrSUBQUOTE_NAME)
    xmlOut.appendChild xmlSubQuoteElem
        
    strTagName = "APPLICATIONNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = strApplicationNo
    xmlSubQuoteElem.appendChild xmlElement
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    xmlElement.Text = ApplicationFFNo
    xmlSubQuoteElem.appendChild xmlElement
    
    strTagName = "PPSUBQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTagName)
    ' APS SYS0767 - Avoid duplicate key error by always creating the latest + 1 record
    'xmlElement.Text = "1"
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
      
    CreateFirstSubQuoteXML = xmlOut.xml
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlElement = Nothing

    Exit Function

CreateFirstSubQuoteXMLVbErr:
    
    Set xmlIn = Nothing
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


Private Function CreateNewSubQuoteXML(vstrXMLRequest) As String

' header ----------------------------------------------------------------------------------
' description:  Creates the xml structure for a new sub quote

' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       New sub quote xml
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewSubQuoteXMLVbErr

    Const strFunctionName As String = "CreateNewSubQuoteXML"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubQuoteNoList As IXMLDOMNodeList
    Dim objPPSubQuoteDO As PaymentProtectionSubQuoteDO
        
    Dim strApplicationNo As String
    Dim ApplicationFFNo As String
    Dim strContext As String
    Dim strPPSubQuoteNo As String
    Dim strTagName As String
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    If objContext Is Nothing Then
        Set objPPSubQuoteDO = New PaymentProtectionSubQuoteDO
    Else
        Set objPPSubQuoteDO = objContext.CreateInstance(App.Title & ".PaymentProtectionSubQuoteDO")
    End If
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    ApplicationFFNo = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    strTagName = "PPSUBQUOTENEMBER"
    strPPSubQuoteNo = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    strTagName = "CONTEXT"
    strContext = objXmlAssist.GetTagValue(xmlIn.documentElement, strTagName)
    
    If Len(strApplicationNo) = 0 Or _
       Len(ApplicationFFNo) = 0 Or Len(strContext) = 0 Then
        Dim strDetails As String
        strDetails = "APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & "APPLICATIONFACTFINDNUMBER = " & ApplicationFFNo
        strDetails = strDetails & "PPSUBQUOTENEMBER = " & strPPSubQuoteNo
        strDetails = strDetails & "CONTEXT = " & strContext
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                strDetails
    End If
    
    Dim strResponse As String
    strResponse = objPPSubQuoteDO.GetSubQuoteDetails(vstrXMLRequest)
    Set xmlOut = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    
    Set xmlSubQuoteElem = xmlOut.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
    
    strTagName = "PPSUBQUOTENUMBER"
    Set xmlSubQuoteNoList = xmlSubQuoteElem.getElementsByTagName(strTagName)
    
    For Each xmlElement In xmlSubQuoteNoList
        xmlElement.Text = ""
    Next
  
    strTagName = "DATEANDTIMEGENERATED"
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
    
    strTagName = "TOTALPPMONTHLYCOST"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    'GD BMIDS00312 21/08/2002 Keep TOTALPPMONTHLYCOST
    'If Not xmlElement Is Nothing Then
    '    xmlElement.Text = ""
    'End If
    
    strTagName = "IPT"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
    
    strTagName = "PREMIUM"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If Not xmlElement Is Nothing Then
        xmlElement.Text = ""
    End If
        
    CreateNewSubQuoteXML = xmlOut.xml
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlElement = Nothing
    Set objPPSubQuoteDO = Nothing
    Set xmlSubQuoteNoList = Nothing

    Exit Function

CreateNewSubQuoteXMLVbErr:
    
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlElement = Nothing
    Set objPPSubQuoteDO = Nothing
    Set xmlSubQuoteNoList = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Sub CalculateAndAddTotals(ByVal vxmlSubQuoteElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Calculates totals and adds them into the xml

' pass:         vxmlSubQuoteElem  xml Element containing sub quote details
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndAddTotalsVbErr

    Const strFunctionName As String = "CalculateAndAddTotals"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objConvAssist As New ConvertAssist
    Dim xmlElement As IXMLDOMElement
    
    
    Dim dblTotalMonthlyCost As Double
    Dim dblTotalPremium As Double
    Dim dblTotalIPT As Double
    Dim dblApp1Premium As Double
    Dim dblApp1IPT As Double
    Dim dblApp2Premium As Double
    Dim dblApp2IPT As Double
    
    Dim strTagName As String
    Dim strTagValue As String
    
    strTagName = "APPLICANT1PREMIUM"
    strTagValue = objXmlAssist.GetTagValue(vxmlSubQuoteElem, strTagName)
    dblApp1Premium = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "APPLICANT1IPT"
    strTagValue = objXmlAssist.GetTagValue(vxmlSubQuoteElem, strTagName)
    dblApp1IPT = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "APPLICANT2PREMIUM"
    strTagValue = objXmlAssist.GetTagValue(vxmlSubQuoteElem, strTagName)
    dblApp2Premium = objConvAssist.CSafeDbl(strTagValue)
    
    strTagName = "APPLICANT2IPT"
    strTagValue = objXmlAssist.GetTagValue(vxmlSubQuoteElem, strTagName)
    dblApp2IPT = objConvAssist.CSafeDbl(strTagValue)
    
    dblTotalPremium = dblApp1Premium + dblApp2Premium
    dblTotalIPT = dblApp1IPT + dblApp2IPT
    dblTotalMonthlyCost = dblTotalPremium + dblTotalIPT

    strTagName = "TOTALPPMONTHLYCOST"
    Set xmlElement = vxmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlSubQuoteElem.ownerDocument.createElement(strTagName)
        vxmlSubQuoteElem.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        xmlElement.Text = Format(dblTotalMonthlyCost, "0.00")
    End If

    strTagName = "IPT"
    Set xmlElement = vxmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlSubQuoteElem.ownerDocument.createElement(strTagName)
        vxmlSubQuoteElem.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        xmlElement.Text = Format(dblTotalIPT, "0.00")
    End If

    strTagName = "PREMIUM"
    Set xmlElement = vxmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlSubQuoteElem.ownerDocument.createElement(strTagName)
        vxmlSubQuoteElem.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        xmlElement.Text = Format(dblTotalPremium, "0.00")
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set objConvAssist = Nothing
    
    Exit Sub

CalculateAndAddTotalsVbErr:
    
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set objConvAssist = Nothing
  
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

