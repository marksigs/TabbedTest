VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmKFIHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************
'** Module:         OmKFIHelper
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    This class provides the functionality to extract the
'**                 appropriate KFI data and optionally format it into the form
'**                 required by the KFI document template.
'********************************************************************************
'********************************************************************************
'BBG Specific History:
'
'Prog   Date        Description
'MV     09/03/2005  BBG1496 - Amended GetFormattedData() and GetDebugFileNameFromGlobalParam()
'TW     09/11/2005  MAR442  - A number of changes to allow omCK.CreateKFI to work
'                             Note that these changes prevent errors where certain data is missing
'                             See the specific methods/functions for the changes
'INR    20/03/2007  EP2_1977  Deal with REGULATIONINDICATOR_VALIDID
'********************************************************************************
Option Explicit
Option Compare Text

Private Function GetDebugFileNameFromGlobalParam(Optional ByVal strAppendToFileName As String = "") As String
    
    On Error GoTo GetDebugFileNameFromGlobalParamVbErr
    
    Dim strDebugPath  As String
    
    strDebugPath = GetGlobalParamString("KFIDebugPath")
    
    If Len(strDebugPath) > 0 Then
        If Right$(strDebugPath, 1) <> "\" Then
            strDebugPath = strDebugPath & "\"
        End If
        If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
            'Create the folder if it does not exist
            MkDir strDebugPath
            If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
                ' The output folder does not exist and cannot be created
                strDebugPath = vbNullString
            End If
        End If
    End If
    
    If Len(strDebugPath) > 0 Then
        'Add a random number so the filename will hopefully still be unique if multiple
        'requests are made in the same second
        Randomize
        Dim strRandom As String
        strRandom = Format$(Rnd * 10000, "0000")
    
        strDebugPath = strDebugPath & Format$(Now(), "YYYYMMDD_HHNNSS") & "_" & strRandom & strAppendToFileName & ".txt"
    End If
    
    GetDebugFileNameFromGlobalParam = strDebugPath

    Exit Function

GetDebugFileNameFromGlobalParamVbErr:
    
    GetDebugFileNameFromGlobalParam = vbNullString
    
    'Clear any errors that may have occurred as they can be ignored
    Err.Clear
                                                    
End Function

'********************************************************************************
'** Function:       GetFormattedData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Extracts the data required for the KFI document from the
'**                 Omiga4 database using the request broker component omRB.dll
'**                 and converts it into the format required for the KFI document
'**                 template.
'** Parameters:     vstrXMLRequest - the request XML.
'**                 vstrKFIDocType - the KFI document type to format the data for.
'** Returns:        The formatted XML data.
'** Errors:         None Expected
'********************************************************************************
Public Function GetFormattedData(ByVal vstrXMLRequest As String, _
        ByVal vstrKFIDocType As String) As String
    Dim strXML As String, strDebugFileName As String
    Dim blnDebug As Boolean
    
    On Error GoTo ErrHandler
    
    '*-get the data
    strXML = GetData(vstrXMLRequest)
    
    strDebugFileName = GetDebugFileNameFromGlobalParam()
    
    If GetGlobalParamBoolean("IsKFIDebugEnabled") Then
        If Len(strDebugFileName) > 0 Then
            blnDebug = True
            SaveDebugInfo strDebugFileName, "Original Request", vstrXMLRequest
            SaveDebugInfo strDebugFileName, "Response from Request Broker", strXML
        End If
    End If
    
    errCheckXMLResponse strXML, True
 
    '*-and return the formatted data
    GetFormattedData = FormatData(strXML, vstrKFIDocType)
    
    If blnDebug Then
        SaveDebugInfo strDebugFileName, "Formatted Date from KFI", GetFormattedData 'SR 29/10/2004 : BBG1734
    End If
    
CleanUpSection:

Exit Function
ErrHandler:
    If errIsSystemError() Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    GetFormattedData = errCreateErrorResponse()
    'go to clean-up section
    Resume CleanUpSection
End Function


'********************************************************************************
'** Function:       GetData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Extracts the data required for the KFI document from the
'**                 Omiga4 database using the request broker component omRB.dll.
'** Parameters:     vstrXMLRequest - the request XML.
'** Returns:        The extracted XML data.
'** Errors:         None Expected
'********************************************************************************
Public Function GetData(ByVal vstrXMLRequest As String) As String
    Dim objRequest As OmRequestDO
    Dim strResponse As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    
    On Error GoTo ErrHandler

    Set objRequest = New OmRequestDO
    strResponse = objRequest.OmDataRequest(vstrXMLRequest)
            
    '*-postprocess the data by converting combo id values to
    '*-validation type id's instead
    Set xmlOut = xmlLoad(strResponse, "GetData")
    Set xmlNode = xmlGetMandatoryNode(xmlOut, "RESPONSE")
    Call ConvertIDToValidationID(xmlNode)
    
    '*-return the data
    GetData = xmlOut.xml

CleanUpSection:
    Set objRequest = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
Exit Function
ErrHandler:
    If errIsSystemError() Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    GetData = errCreateErrorResponse()
    'go to clean-up section
    Resume CleanUpSection
End Function

'********************************************************************************
'** Function:       FormatData
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Takes data in the format returned by the GetData method of
'**                 this class and converts it into the XML format required for
'**                 the KFI document template.
'** Parameters:     vstrXMLRequest - the request XML.
'**                 vstrKFIDocType - the KFI document type to format the data for.
'** Returns:        The formatted XML data.
'** Errors:         None Expected
'********************************************************************************
Public Function FormatData(ByVal vstrXMLRequest As String, _
        ByVal vstrKFIDocType As String) As String
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objCommon As CommonDataHelper
    Dim objMortgage As IMortgage
    Dim xmlData As IXMLDOMNode
    
    On Error GoTo ErrHandler
    
    Set xmlIn = xmlLoad(vstrXMLRequest, "FormatData")
    xmlIn.setProperty "SelectionLanguage", "XPath"
    
    Set xmlData = xmlIn.selectSingleNode("//APPLICATION")
    If Not xmlData Is Nothing Then
        '*-create the common data helper
        Set objCommon = New CommonDataHelper
        Call objCommon.Initialise(xmlData, vstrKFIDocType) 'SR 10/09/2004 : CORE82
        
        If vstrKFIDocType = "STANDARD" Then
            Set objMortgage = New StandardMortgage
            Call objMortgage.FormatData(objCommon)
        ElseIf vstrKFIDocType = "OFFER" Then
            Set objMortgage = New MortgageOffer
            Call objMortgage.FormatData(objCommon)
        ElseIf vstrKFIDocType = "LIFETIME" Then
            Set objMortgage = New LifetimeMortgage
            Call objMortgage.FormatData(objCommon)
        ElseIf vstrKFIDocType = "LIFETIMEOFFER" Then
            Set objMortgage = New LifetimeOffer
            Call objMortgage.FormatData(objCommon)
        ElseIf vstrKFIDocType = "TRANSFEROFFER" Then
            Set objMortgage = New TransferOffer
            Call objMortgage.FormatData(objCommon)
        ElseIf vstrKFIDocType = "NONREG" Then
            Set objMortgage = New NonRegMortgage
            Call objMortgage.FormatData(objCommon)
        End If
        
        '*-return the formatted data
        FormatData = objCommon.Document.xml
    End If

CleanUpSection:
    Set xmlIn = Nothing
    Set objCommon = Nothing
    Set objMortgage = Nothing
    Set xmlData = Nothing

Exit Function
ErrHandler:
    If errIsSystemError() Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    FormatData = errCreateErrorResponse()
    'go to clean-up section
    Resume CleanUpSection
End Function

'********************************************************************************
'** Function:       ConvertIDToValidationID
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Converts combo ID's in the XML to combo validation ID's. It
'**                 does so by copying the XXX_VALIDID value (which is a comma
'**                 separated list of validation ID's) over the ID value.  NB a
'**                 comma is appended to the list of validation ID's to enable
'**                 the use of Instr to identify if it contains a particular
'**                 validation ID e.g. If the list is E,EOC,EN, then to test if
'**                 the list contains E you can use:
'**                 If Instr(sList, "E,") > 0 Then
'** Parameters:     vxmlNode - the XML to convert.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub ConvertIDToValidationID(ByVal vxmlNode As IXMLDOMNode)
    Const cstrFunctionName As String = "ConvertIDToValidationID"
    Dim xmlList As IXMLDOMNodeList
    Dim xmlItem As IXMLDOMNode
    Dim strValue As String
        
    On Error GoTo ErrHandler

    '*-select the APPLICATION node in the node and convert the
    '*-TYPEOFBUYER values to their appropriate validation type ID's
    Set xmlItem = vxmlNode.selectSingleNode("//APPLICATION")
    If xmlItem Is Nothing Then
        '*-we have not retrieved any data at all, so quit
        Exit Sub
    End If
    
    If xmlAttributeValueExists(xmlItem, "TYPEOFBUYER_VALIDID") Then
        strValue = xmlGetAttributeText(xmlItem, "TYPEOFBUYER_VALIDID") & ","
        Call xmlSetAttributeValue(xmlItem, "TYPEOFBUYER", strValue)
        Call xmlItem.Attributes.removeNamedItem("TYPEOFBUYER_VALIDID")
    End If
    
    '*-select all the APPLICATIONFACTFIND nodes in the node and convert the
    '*-LEVELOFADVICE values to their appropriate validation type ID's
    '*-TYPEOFAPPLICATION values to their appropriate validation type ID's
    '*-DIRECTINDIRECTBUSINESS values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//APPLICATIONFACTFIND")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "LEVELOFADVICE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "LEVELOFADVICE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "LEVELOFADVICE", strValue)
            Call xmlItem.Attributes.removeNamedItem("LEVELOFADVICE_VALIDID")
        End If
        If xmlAttributeValueExists(xmlItem, "TYPEOFAPPLICATION_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "TYPEOFAPPLICATION", strValue)
            Call xmlItem.Attributes.removeNamedItem("TYPEOFAPPLICATION_VALIDID")
        End If
        If xmlAttributeValueExists(xmlItem, "DIRECTINDIRECTBUSINESS_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "DIRECTINDIRECTBUSINESS_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "DIRECTINDIRECTBUSINESS", strValue)
            Call xmlItem.Attributes.removeNamedItem("DIRECTINDIRECTBUSINESS_VALIDID")
        End If
        'SR 13/10/2004 : BBG1592
        If xmlAttributeValueExists(xmlItem, "SPECIALGROUP_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "SPECIALGROUP_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "SPECIALGROUP", strValue)
            Call xmlItem.Attributes.removeNamedItem("SPECIALGROUP_VALIDID")
        End If
        'SR 13/10/2004 : BBG1592 - End
        'EP2_1977
        If xmlAttributeValueExists(xmlItem, "REGULATIONINDICATOR_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "REGULATIONINDICATOR_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "REGULATIONINDICATOR", strValue)
            Call xmlItem.Attributes.removeNamedItem("REGULATIONINDICATOR_VALIDID")
        End If
    Next xmlItem
    
    '*-select all the MORTGAGEONEOFFCOST nodes in the node and convert the
    '*-MORTGAGEONECOSTTYPE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//MORTGAGEONEOFFCOST")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "MORTGAGEONEOFFCOSTTYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "MORTGAGEONEOFFCOSTTYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "MORTGAGEONEOFFCOSTTYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("MORTGAGEONEOFFCOSTTYPE_VALIDID")
        End If
    Next xmlItem
    
    '*-select all the LOANCOMPONENT nodes in the node and convert the
    '*-REPAYMENTMETHOD values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//LOANCOMPONENT")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "REPAYMENTMETHOD_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "REPAYMENTMETHOD", strValue)
            Call xmlItem.Attributes.removeNamedItem("REPAYMENTMETHOD_VALIDID")
        End If
        
        If xmlAttributeValueExists(xmlItem, "PURPOSEOFLOAN_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "PURPOSEOFLOAN_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "PURPOSEOFLOAN", strValue)
            Call xmlItem.Attributes.removeNamedItem("PURPOSEOFLOAN_VALIDID")
        End If
    Next xmlItem

    '*-select all the FEEPAYMENT nodes in the node and convert the
    '*-PAYMENTEVENT values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//FEEPAYMENT")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "PAYMENTEVENT_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "PAYMENTEVENT_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "PAYMENTEVENT", strValue)
            Call xmlItem.Attributes.removeNamedItem("PAYMENTEVENT_VALIDID")
        End If
    Next xmlItem
    
    '*-select all the BUILDINGSANDCONTENTSDETAILS nodes in the node and convert the
    '*-COVERTYPE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//BUILDINGSANDCONTENTSDETAILS")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "COVERTYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "COVERTYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "COVERTYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("COVERTYPE_VALIDID")
        End If
    Next xmlItem

    '*-select all the APPLICATIONINTERMEDIARY nodes in the node and convert the
    '*-INTERMEDIARYTYPE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//APPLICATIONINTERMEDIARY")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "INTERMEDIARYTYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "INTERMEDIARYTYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "INTERMEDIARYTYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("INTERMEDIARYTYPE_VALIDID")
        End If
    Next xmlItem
    
    '*-select the NEWPROPERTY nodes in the node APPLICATIONFACTFIND and convert the
    '*-TENURE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//NEWPROPERTY")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "TENURETYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "TENURETYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "TENURETYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("TENURETYPE_VALIDID")
        End If
        'SR 08/10/2004 : BBG1545
        If xmlAttributeValueExists(xmlItem, "VALUATIONTYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "VALUATIONTYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "VALUATIONTYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("VALUATIONTYPE_VALIDID")
        End If
        'SR 08/10/2004 : BBG1545 - End
    Next xmlItem
    
    '*-select all the CUSTOMERROLE nodes in the APPLICATIONFACTFIND node and convert the
    '*- CUSTOMERROLE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes(gcstrCUSTOMERROLE_PATH)
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "CUSTOMERROLETYPE_VALIDID") Then
            strValue = xmlGetAttributeText(xmlItem, "CUSTOMERROLETYPE_VALIDID") & ","
            Call xmlSetAttributeValue(xmlItem, "CUSTOMERROLETYPE", strValue)
            Call xmlItem.Attributes.removeNamedItem("CUSTOMERROLETYPE_VALIDID")
        End If
    Next xmlItem

    Set xmlList = Nothing
    Set xmlItem = Nothing
Exit Sub
ErrHandler:
    Set xmlList = Nothing
    Set xmlItem = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub SaveDebugInfo(ByVal vstrFileName As String, ByVal vstrMessage As String, ByVal vstrXML As String)
    
    Dim objFSO          As Object
    Dim objTextStream   As Object
    
    On Error Resume Next
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextStream = objFSO.OpenTextFile(vstrFileName, 8, True)
    
    objTextStream.WriteLine "------------------------------------------------------------------------------------------"
    objTextStream.WriteLine Format$(Now(), "DD/MM/YYYY HH:NN:SS") & ": " & vstrMessage
    objTextStream.WriteLine "------------------------------------------------------------------------------------------"
    
    objTextStream.WriteLine vstrXML & vbCrLf
    
    objTextStream.Close
    
    Set objFSO = Nothing
    Set objTextStream = Nothing
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    
End Sub
