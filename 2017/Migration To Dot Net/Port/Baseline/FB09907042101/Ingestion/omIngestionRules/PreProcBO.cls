VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "PreProcBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     18/10/2005  created for Project MARS
'IK     22/10/2005  MAR271 - Add APPLICATIONPRIORITY when application created
'IK     25/10/2005  MAR232 - cannot remove APPLICATIONFACTFIND node
'                       required by post-proc
'                   modify CompareNode to match xpath references
'                   update LASTAMENDEDDATE on current CUSTOMERADDRESS as required
'                   compare ACCEPTEDQUOTENUMBER and ACTIVEQUOTENUMBER on input,
'                        error if different
'                   compare exiting ACCEPTEDQUOTENUMBER and new ACCEPTEDQUOTENUMBER,
'                       set postproc flag if different
'                   re-stucturing, RemoveRedundantNodes moved to IngestionManagerBO
'IK     25/11/2005  MAR695 gstrComponentId & gstrComponentResponse now private member variables
'IK     30/11/2005  MAR707 - use omiga4 components for new APPLICATIONNUMBER / CUSTOMERNUMBER
'IK     01/12/2005  MAR767 - modify OutputRules, no UPDATE on CUSTOMERADDRESS when CREATE
'IK     05/12/2005  MAR666 - loadtest deadlock, add _LOCKHINT to search in AddCustomerNumbers
'RF     30/01/2006  MAR1124 - in CompareData need to transform GetApplicationData response before use
'                             as AccountRelationship needs converting to OtherCustomerAccountRelationship
'IK     14/02/2006  MAR1264 - create PAYEEHISTORY record on SubmitFMA if not present
'IK     23/03/2006  EP2 - add espom default values to request
'-----------------------------------------------------------------------------
Option Explicit

Private gstrComponentId As String
Private gstrComponentResponse As String

Public Function OmRequest(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "OmRequest"
    On Error GoTo OmRequestVbErr
    
    GetObjectContext.SetComplete
    
    LoadConfigFile
    
    Dim xmlInDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMElement
    
    'IK 25/10/2005 MAR232
    Dim blnIsCreate As Boolean
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    Set xmlRequestNode = xmlInDoc.selectSingleNode("REQUEST")
    If xmlRequestNode Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no REQUEST node"
    End If
    
    ValidateRequest xmlRequestNode
    
    InputRules xmlRequestNode
    
    'IK 25/10/2005 MAR232
    blnIsCreate = IsCreate(xmlRequestNode)
    
    If blnIsCreate Then
        AddPrimaryKeys xmlRequestNode
        'IK 22/10/2005 MAR271
        AddApplicationPriority xmlRequestNode
        'IK 25/10/2005 MAR232
        CreateAcceptedQuoteTest xmlRequestNode
    Else
        CompareData xmlRequestNode
    End If
    
    If AnyCrudOp(xmlRequestNode) Then
        xmlRequestNode.setAttribute "CRUD_OP", "iUPDATE"
        xmlRequestNode.setAttribute "SCHEMA_NAME", "WebServiceSchema"
        xmlRequestNode.setAttribute "ENTITY_REF", "AIPINGESTION"
    End If
    
    'IK 25/10/2005 MAR232
    OutputRules xmlRequestNode, blnIsCreate
    
    OmRequest = xmlInDoc.xml
    
OmRequestVbErr:

    Set gxmldocConfig = Nothing
    Set xmlInDoc = Nothing
    
    If Err.Number <> 0 Then
        If Err.Source = "OMIGAERROR" Then
            OmRequest = Err.Description
        Else
            If Err.Source <> cstrMethodName Then
                If Err.Source = App.EXEName Then
                    Err.Source = cstrMethodName
                Else
                    Err.Source = cstrMethodName & "." & Err.Source
                End If
            End If
            Err.Source = App.EXEName & "." & TypeName(Me) & "." & Err.Source
            OmRequest = FormatError(vstrRequest, gstrComponentId, gstrComponentResponse)
        End If
    End If

End Function

Private Sub ValidateRequest(ByVal vxmlRequestNode As IXMLDOMNode)

    Const cstrMethodName As String = "ValidateRequest"
    On Error GoTo ValidateRequestExit
    
    If vxmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no OPERATION on pre-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("USERID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no USERID on pre-processing REQUEST node"
    End If
    
    If vxmlRequestNode.Attributes.getNamedItem("UNITID") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no UNITID on pre-processing REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("APPLICATION") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION node on pre-processing REQUEST node"
    End If
    
    If vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND") Is Nothing Then
        Err.Raise oeMissingPrimaryTag, cstrMethodName, "no APPLICATION/APPLICATIONFACTFIND node on pre-processing REQUEST node"
    End If
    
    'ik EP2 23/03/2006
    If vxmlRequestNode.Attributes.getNamedItem("USERID").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERID]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("USERID").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERID]").Attributes.getNamedItem("USERID").Text
        End If
    End If
    If vxmlRequestNode.Attributes.getNamedItem("UNITID").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@UNITID]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("UNITID").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@UNITID]").Attributes.getNamedItem("UNITID").Text
        End If
    End If
    If vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text = "epsom" Then
        If Not gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERAUTHORITYLEVEL]") Is Nothing Then
            vxmlRequestNode.Attributes.getNamedItem("USERAUTHORITYLEVEL").Text = _
                gxmldocConfig.selectSingleNode("omIngestion/default[@entity='REQUEST'][@USERAUTHORITYLEVEL]").Attributes.getNamedItem("USERAUTHORITYLEVEL").Text
        End If
    End If
    'ik EP2 23/03/2006 ends
    
ValidateRequestExit:
    
    CheckError cstrMethodName

End Sub

Private Function IsCreate(ByVal vxmlRequestNode As IXMLDOMElement) As Boolean

    Const cstrMethodName As String = "IsCreate"
    On Error GoTo IsCreateExit

    If vxmlRequestNode.selectSingleNode("APPLICATION[@APPLICATIONNUMBER]") Is Nothing Then
        IsCreate = True
    End If
    
IsCreateExit:
    
    CheckError cstrMethodName

End Function

Private Sub InputRules(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "InputRules"
    On Error GoTo InputRulesExit
    
    Dim xmlTarget As IXMLDOMElement
    Dim xmlSource As IXMLDOMElement
    
    For Each xmlTarget In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE='3'][not(@DATEMOVEDOUT)]")
    
        Set xmlSource = xmlTarget.nextSibling
        
        If Not xmlSource Is Nothing Then
            If xmlSource.nodeName <> "CUSTOMERADDRESS" Then
                Set xmlSource = Nothing
            ElseIf xmlSource.getAttribute("ADDRESSTYPE") <> "3" Then
                Set xmlSource = Nothing
            End If
        End If
    
        If xmlSource Is Nothing Then
            Set xmlSource = xmlTarget.selectSingleNode("../CUSTOMERADDRESS[@ADDRESSTYPE='1']")
        End If
            
        If Not xmlSource Is Nothing Then
            If Not xmlSource.Attributes.getNamedItem("DATEMOVEDIN") Is Nothing Then
                xmlTarget.setAttribute "DATEMOVEDOUT", xmlSource.getAttribute("DATEMOVEDIN")
            End If
        End If
    
    Next
    
InputRulesExit:
    
    Set xmlSource = Nothing
    Set xmlTarget = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub OutputRules(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vblnIsCreate As Boolean)

    Const cstrMethodName As String = "OutputRules"
    On Error GoTo OutputRulesExit
    
    Dim xmlCustAddress As IXMLDOMElement
    Dim xmlAddress As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strNodeOp As String
    Dim blnNodeOp As Boolean
    
    For Each xmlCustAddress In vxmlRequestNode.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE='1']")
    
        blnNodeOp = vblnIsCreate
        
        If Not blnNodeOp Then
        
            Set xmlAttrib = xmlCustAddress.Attributes.getNamedItem("CRUD_OP")
            If Not xmlAttrib Is Nothing Then
                
                strNodeOp = xmlAttrib.Text
                If strNodeOp = "CREATE" Or strNodeOp = "UPDATE" Then
                    blnNodeOp = True
                End If
                
            End If
            
        End If
                
        If Not blnNodeOp Then
                
            Set xmlAddress = xmlCustAddress.selectSingleNode("ADDRESS[@CRUD_OP]")
            If Not xmlAddress Is Nothing Then
                
                strNodeOp = xmlAddress.Attributes.getNamedItem("CRUD_OP").Text
                If strNodeOp = "CREATE" Or strNodeOp = "UPDATE" Then
                    blnNodeOp = True
                End If
                
            End If
        
        End If
        
        If blnNodeOp Then
            xmlCustAddress.setAttribute "LASTAMENDEDDATE", FormatNow
            'IK_MAR767_20051201
            If Not vblnIsCreate Then
                If xmlCustAddress.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
                    If xmlCustAddress.parentNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
                        xmlCustAddress.setAttribute "CRUD_OP", "UPDATE"
                    End If
                End If
            End If
        End If
    
    Next
    
OutputRulesExit:
    
    Set xmlCustAddress = Nothing
    
    CheckError cstrMethodName

End Sub

Private Sub AddPrimaryKeys(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "AddPrimaryKeys"
    On Error GoTo AddPrimaryKeysExit
    
    Dim xmlApplicationNode As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    
    'ik_mar707_20051130
    Dim strChannelId As String
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    xmlApplicationNode.setAttribute "CRUD_OP", "CREATE"
    
    'ik_mar707_20051130
    If vxmlRequestNode.Attributes.getNamedItem("CHANNELID") Is Nothing Then
        strChannelId = "1"
    Else
        strChannelId = vxmlRequestNode.Attributes.getNamedItem("CHANNELID").Text
    End If
    AddApplicationNumber xmlApplicationNode, strChannelId
    'ik_mar707_20051130_ends
    
    Set xmlNodeList = xmlApplicationNode.selectNodes("APPLICATIONFACTFIND/CUSTOMERROLE[not(@CUSTOMERNUMBER)]")
    
    If xmlNodeList.length <> 0 Then
        AddCustomerNumbers xmlNodeList
    End If
    
AddPrimaryKeysExit:
    
    CheckError cstrMethodName

End Sub

'ik_mar707_20051130
Private Sub AddApplicationNumber( _
    ByVal vxmlApplicationNode As IXMLDOMElement, _
    ByVal vstrChannelId As String)
    
    Const cstrMethodName As String = "AddCustomerNumbers"
    On Error GoTo AddApplicationNumberExit
    
    Dim strNewApplicationNumber As String
    Dim strComponentResponse As String
    
    Dim xmlFFNode As IXMLDOMElement
    
    'ik_mar707_20051130
    Dim omObj As Object
    Set omObj = GetObjectContext.CreateInstance("omAPP.ApplicationBO")
    strComponentResponse = omObj.GetNextApplicationNumber(vstrChannelId)
    Set omObj = Nothing
    
    If Len(strComponentResponse) = 0 Then
        gstrComponentId = "omAPP.ApplicationBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving next APPLICATIONNUMBER"
    Else
        strNewApplicationNumber = strComponentResponse
    End If
    
    If Not gxmldocConfig.selectSingleNode("omIngestion/default[@APPLICATIONNUMBERPREFIX]") Is Nothing Then
        strNewApplicationNumber = gxmldocConfig.selectSingleNode("omIngestion/default/@APPLICATIONNUMBERPREFIX").Text & strNewApplicationNumber
    End If
    
    vxmlApplicationNode.setAttribute "APPLICATIONNUMBER", strNewApplicationNumber
    
    Set xmlFFNode = vxmlApplicationNode.selectSingleNode("APPLICATIONFACTFIND")
    xmlFFNode.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Set xmlFFNode = Nothing
    
AddApplicationNumberExit:

    Set xmlFFNode = Nothing

    CheckError cstrMethodName

End Sub

Private Sub AddCustomerNumbers(ByVal vxmlCustomerNodes As IXMLDOMNodeList)

    Const cstrMethodName As String = "AddCustomerNumbers"
    On Error GoTo AddCustomerNumbersExit
    
    Dim strNewCustomerNumber As String
    Dim strComponentResponse As String
    Dim xmlChildNode As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    
    Dim xmlDoc As DOMDocument40
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False

    Dim crudObj As Object
    Dim omObj As Object
    
    For Each xmlNode In vxmlCustomerNodes
    
        xmlDoc.loadXML ""
    
        If Not xmlNode.selectSingleNode("CUSTOMER[@OTHERSYSTEMCUSTOMERNUMBER]") Is Nothing Then
                
            Set xmlElem = xmlDoc.createElement("REQUEST")
            xmlElem.setAttribute "CRUD_OP", "READ"
            xmlElem.setAttribute "ENTITY_REF", "CUSTOMER"
            'IK_05/12/2005_MAR666
            xmlElem.setAttribute "_LOCKHINT", "NOLOCK"
            Set xmlChildNode = xmlDoc.appendChild(xmlElem)
            
            Set xmlElem = xmlDoc.createElement("CUSTOMER")
            xmlElem.setAttribute "OTHERSYSTEMCUSTOMERNUMBER", xmlNode.selectSingleNode("CUSTOMER/@OTHERSYSTEMCUSTOMERNUMBER").Text
            xmlChildNode.appendChild xmlElem
            
            Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
            strComponentResponse = crudObj.OmRequest(xmlDoc.xml)
            Set crudObj = Nothing
            xmlDoc.loadXML strComponentResponse
        
            If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
                ChainResponseDescription xmlDoc, "error retrieving CUSTOMER for OTHERSYSTEMCUSTOMERNUMBER"
                Err.Raise oeUnspecifiedError, "OMIGAERROR", xmlDoc.xml
            End If
            
            If Not xmlDoc.selectSingleNode("RESPONSE/CUSTOMER[@CUSTOMERNUMBER]") Is Nothing Then
                
                xmlNode.setAttribute "CUSTOMERNUMBER", xmlDoc.selectSingleNode("RESPONSE/CUSTOMER/@CUSTOMERNUMBER").Text
                
                Set xmlChildNode = xmlNode.selectSingleNode("CUSTOMER")
                xmlChildNode.setAttribute "CRUD_OP", "UPDATE"
                
                Set xmlChildNode = xmlNode.selectSingleNode("CUSTOMER/CUSTOMERVERSION")
                If Not xmlChildNode Is Nothing Then
                    xmlChildNode.setAttribute "CRUD_OP", "CREATE"
                End If
            
            End If
            
        End If
        
        If xmlNode.Attributes.getNamedItem("CUSTOMERNUMBER") Is Nothing Then
            
            Set omObj = GetObjectContext.CreateInstance("omCust.CustomerBO")
            strComponentResponse = omObj.GetNextCustomerNumber()
            Set omObj = Nothing
            xmlDoc.loadXML strComponentResponse
            
            If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/CUSTOMER[@CUSTOMERNUMBER]") Is Nothing Then
                ChainResponseDescription xmlDoc, "error retrieving next CUSTOMERNUMBER"
                Err.Raise oeUnspecifiedError, "OMIGAERROR", xmlDoc.xml
            End If
            
            strNewCustomerNumber = xmlDoc.selectSingleNode("RESPONSE/CUSTOMER/@CUSTOMERNUMBER").Text
            
            xmlNode.setAttribute "CUSTOMERNUMBER", strNewCustomerNumber
            
        End If
    
    Next
    
AddCustomerNumbersExit:
    
    Set crudObj = Nothing
    Set omObj = Nothing
    Set xmlChildNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing

    CheckError cstrMethodName

End Sub

'IK 22/10/2005 MAR271
Private Sub AddApplicationPriority(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "AddApplicationPriority"
    On Error GoTo AddApplicationPriorityExit
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNode = gxmldocConfig.selectSingleNode("omIngestion/default[@entity='APPLICATIONPRIORITY']/@APPLICATIONPRIORITYVALUE")
    If xmlNode Is Nothing Then
        Err.Raise oeXMLMissingAttribute, cstrMethodName, "no default APPLICATIONPRIORITY in ingestionConfig.xml"
    End If
    
    Set xmlElem = vxmlRequestNode.ownerDocument.createElement("APPLICATIONPRIORITY")
    xmlElem.setAttribute "APPLICATIONPRIORITYVALUE", xmlNode.Text
    xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
    xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
    xmlElem.setAttribute "DATETIMEASSIGNED", FormatNow()
    
    vxmlRequestNode.selectSingleNode("APPLICATION").appendChild xmlElem
    
AddApplicationPriorityExit:

    Set xmlNode = Nothing
    Set xmlElem = Nothing
    
    CheckError cstrMethodName

End Sub
'IK 22/10/2005 MAR271 ends

' IK 14/02/2006 MAR1264
Private Sub AddPayeeHistory( _
    ByVal vxmlRequestApplicationNode As IXMLDOMElement, _
    ByVal vxmlLegalRep As IXMLDOMElement)

    Const cstrMethodName As String = "AddPayeeHistory"
    On Error GoTo AddPayeeHistoryExit

    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlSrc As IXMLDOMElement
    
    Set xmlElem = vxmlRequestApplicationNode.ownerDocument.createElement("PAYEEHISTORY")
    CopyAttribute vxmlLegalRep, xmlElem, "DXID"
    CopyAttribute vxmlLegalRep, xmlElem, "DXLOCATION"
    xmlElem.setAttribute "CRUD_OP", "CREATE"
    Set xmlNode = vxmlRequestApplicationNode.appendChild(xmlElem)
    
    If Not vxmlLegalRep.selectSingleNode("NAMEANDADDRESSDIRECTORY") Is Nothing Then
        Set xmlElem = vxmlRequestApplicationNode.ownerDocument.createElement("THIRDPARTY")
        Set xmlSrc = vxmlLegalRep.selectSingleNode("NAMEANDADDRESSDIRECTORY")
        CopyThirdPartyAttributes xmlSrc, xmlElem
        Set xmlNode = xmlNode.appendChild(xmlElem)
    ElseIf Not vxmlLegalRep.selectSingleNode("THIRDPARTY") Is Nothing Then
        Set xmlElem = vxmlRequestApplicationNode.ownerDocument.createElement("THIRDPARTY")
        Set xmlSrc = vxmlLegalRep.selectSingleNode("THIRDPARTY")
        CopyThirdPartyAttributes xmlSrc, xmlElem
        Set xmlNode = xmlNode.appendChild(xmlElem)
    End If
    
    If xmlNode.nodeName = "THIRDPARTY" Then
    
        If xmlNode.Attributes.getNamedItem("THIRDPARTYTYPE") Is Nothing Then
            xmlNode.setAttribute "THIRDPARTYTYPE", "10"
        End If
        
        If Not xmlSrc.selectSingleNode("ADDRESS") Is Nothing Then
            Set xmlElem = vxmlRequestApplicationNode.ownerDocument.createElement("ADDRESS")
            Set xmlSrc = xmlSrc.selectSingleNode("ADDRESS")
            CopyThirdPartyAttributes xmlSrc, xmlElem
            xmlNode.appendChild xmlElem
            
        End If
    
    End If
    
AddPayeeHistoryExit:
    
    Set xmlSrc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    CheckError cstrMethodName

End Sub
' IK 14/02/2006 MAR1264 ends

'IK 25/10/2005 MAR232
Private Sub CreateAcceptedQuoteTest(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "CreateAcceptedQuoteTest"
    On Error GoTo CreateAcceptedQuoteTestExit
    
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND[@ACCEPTEDQUOTENUMBER]")
    
    If Not xmlNode Is Nothing Then
        If IsNewAcceptedQuote(xmlNode) Then
            vxmlRequestNode.setAttribute "ACCEPTQUOTE", "true"
        End If
    End If
    
CreateAcceptedQuoteTestExit:

    Set xmlNode = Nothing
    
    CheckError cstrMethodName

End Sub
'IK 25/10/2005 MAR232 ends

Private Sub CompareData(ByVal vxmlRequestNode As IXMLDOMElement)

    Const cstrMethodName As String = "CompareData"
    On Error GoTo CompareDataExit

    ' omCRUD call to retrieve existing data
    Dim crudObj As Object
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMElement
    Dim xmlNewApplicationFactFindNode As IXMLDOMElement
    
    Dim strComponentResponse As String, _
        strApplicationNumber As String, _
        strApplicationFactFindNumber As String
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    
    strApplicationNumber = xmlApplicationNode.Attributes.getNamedItem("APPLICATIONNUMBER").Text
    
    If Not xmlApplicationNode.selectSingleNode("APPLICATIONFACTFIND[@APPLICATIONFACTFINDNUMBER]") Is Nothing Then
        strApplicationFactFindNumber = xmlApplicationNode.selectSingleNode("APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    Else
        strApplicationFactFindNumber = "1"
    End If
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.loadXML GetApplicationData(vxmlRequestNode)
    
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/APPLICATION") Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving comparison data "
    End If
    
    Set xmlNode = xmlDoc.selectSingleNode("RESPONSE")
    If xmlNode Is Nothing Then
        gstrComponentId = "omCRUD.omCRUDBO"
        gstrComponentResponse = strComponentResponse
        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving comparison data "
    End If
    
    'RF 30/01/2006 MAR1124 Start
    ' transform response as AccountRelationship needs converting to OtherCustomerAccountRelationship ...
    xmlTransform xmlDoc, "transformPreProcApplicationData.xslt"
    'RF 30/01/2006 MAR1124 End
        
    CompareNode _
        vxmlRequestNode.selectSingleNode("APPLICATION"), _
        xmlDoc.selectSingleNode("RESPONSE/APPLICATION"), _
        gxmldocConfig.selectSingleNode("omIngestion/match[@entity='APPLICATION']")
        
    ' IK 25/10/2005 MAR232 - new accepted quote?
    Set xmlNewApplicationFactFindNode = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND[@ACCEPTEDQUOTENUMBER]")
    If Not xmlNewApplicationFactFindNode Is Nothing Then
        If IsNewAcceptedQuote(xmlNewApplicationFactFindNode, xmlDoc.selectSingleNode("RESPONSE/APPLICATION/APPLICATIONFACTFIND")) Then
            vxmlRequestNode.setAttribute "ACCEPTQUOTE", "true"
        End If
    End If
    ' IK 25/10/2005 MAR232 - ends
    
    ' IK 14/02/2006 MAR1264
    If vxmlRequestNode.getAttribute("OPERATION") = "SubmitFMA" Then
        If vxmlRequestNode.selectSingleNode("APPLICATION/PAYEEHISTORY") Is Nothing Then
            If Not xmlDoc.selectSingleNode("RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP") Is Nothing Then
                AddPayeeHistory vxmlRequestNode.selectSingleNode("APPLICATION"), xmlDoc.selectSingleNode("RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP")
            End If
        End If
    End If
    ' IK 14/02/2006 MAR1264 ends
    
CompareDataExit:
    
    Set crudObj = Nothing
    Set xmlRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlNewApplicationFactFindNode = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing

    CheckError cstrMethodName

End Sub

Private Sub CompareNode( _
    ByVal vxmlNewNode As IXMLDOMElement, _
    ByVal vxmlOldNode As IXMLDOMElement, _
    ByVal vxmlConfigNode As IXMLDOMElement)

    Const cstrMethodName As String = "CompareNode"
    On Error GoTo CompareNodeExit
    
    Dim xmlNewAttrib As IXMLDOMAttribute
    Dim xmlOldAttrib As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMElement
    
    Dim strName As String
    
    If Not vxmlConfigNode Is Nothing Then
    
        For Each xmlNode In vxmlConfigNode.selectNodes("clone")
        
            strName = xmlNode.getAttribute("attribute")
        
            If vxmlNewNode.Attributes.getNamedItem(strName) Is Nothing Then
                If Not vxmlOldNode.Attributes.getNamedItem(strName) Is Nothing Then
                    vxmlNewNode.setAttribute strName, vxmlOldNode.getAttribute(strName)
                End If
            End If
        
        Next
        
    End If
    
    If vxmlNewNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
        
        For Each xmlNewAttrib In vxmlNewNode.Attributes
                
            'IK 25/10/2005 MAR232
            If Left(xmlNewAttrib.Text, 6) = "xpath:" Then
            
                Set xmlOldAttrib = vxmlOldNode.selectSingleNode(Mid(xmlNewAttrib.Text, 7))
                
                If Not xmlOldAttrib Is Nothing Then
                
                    If vxmlOldNode.Attributes.getNamedItem(xmlNewAttrib.Name) Is Nothing Then
                        xmlNewAttrib.Text = xmlOldAttrib.Text
                        vxmlNewNode.setAttribute "CRUD_OP", "UPDATE"
                    ElseIf vxmlOldNode.Attributes.getNamedItem(xmlNewAttrib.Name).Text <> xmlOldAttrib.Text Then
                        xmlNewAttrib.Text = xmlOldAttrib.Text
                        vxmlNewNode.setAttribute "CRUD_OP", "UPDATE"
                    Else
                        vxmlNewNode.removeAttribute xmlNewAttrib.Name
                    End If
                    
                Else
                
                    vxmlNewNode.removeAttribute xmlNewAttrib.Name
                
                End If
                'IK 25/10/2005 MAR232 ends
                
            Else
    
                If vxmlOldNode.Attributes.getNamedItem(xmlNewAttrib.Name) Is Nothing Then
                    vxmlNewNode.setAttribute "CRUD_OP", "UPDATE"
                ElseIf vxmlOldNode.Attributes.getNamedItem(xmlNewAttrib.Name).Text <> xmlNewAttrib.Text Then
                    vxmlNewNode.setAttribute "CRUD_OP", "UPDATE"
                End If
            
            End If
            
            If Not vxmlNewNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
                Exit For
            End If
    
        Next
                
    End If

    CompareChildNodes vxmlNewNode, vxmlOldNode
    
CompareNodeExit:

    Set xmlOldAttrib = Nothing
    Set xmlNewAttrib = Nothing
    Set xmlNode = Nothing

    CheckError cstrMethodName

End Sub
    
Private Sub CompareChildNodes(ByVal vxmlNewNode As IXMLDOMElement, ByVal vxmlOldNode As IXMLDOMElement)

    Const cstrMethodName As String = "CompareChildNodes"
    On Error GoTo CompareChildNodesExit
    
    Dim xmlNewNode As IXMLDOMNode, _
        xmlOldNode As IXMLDOMNode, _
        xmlNewAttrib As IXMLDOMAttribute, _
        xmlNewElem As IXMLDOMElement, _
        xmlNode As IXMLDOMNode
        
    Dim strNodeName As String
    
    For Each xmlNewNode In vxmlNewNode.childNodes
        
        If xmlNewNode.nodeName <> strNodeName Then
        
            strNodeName = xmlNewNode.nodeName
            
            If IsValidEdit(strNodeName) Then
            
                If gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & strNodeName & "']/primary[@attribute]") Is Nothing Then
                
                    CompareChildNodesByPosition vxmlNewNode, vxmlOldNode, strNodeName
                
                Else
                
                    CompareChildNodesByMatch vxmlNewNode, vxmlOldNode, strNodeName
                
                End If
                
            End If
            
        End If
    
    Next
    
    strNodeName = ""
    For Each xmlOldNode In vxmlOldNode.childNodes
        
        If xmlOldNode.nodeName <> strNodeName Then
        
            strNodeName = xmlOldNode.nodeName
            
            If IsValidEdit(strNodeName) Then
            
                If vxmlNewNode.selectNodes(strNodeName).length = 0 Then
                
                    If gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & strNodeName & "'][@noDelete]") Is Nothing Then
                    
                        CompareChildNodesByPosition vxmlNewNode, vxmlOldNode, strNodeName
                        
                    End If
                    
                End If
            
            End If
        
        End If
    
    Next
    
CompareChildNodesExit:

    CheckError cstrMethodName

End Sub
    
Private Sub CompareChildNodesByPosition( _
    ByVal vxmlNewNode As IXMLDOMElement, _
    ByVal vxmlOldNode As IXMLDOMElement, _
    ByVal vstrNodeName As String)

    Const cstrMethodName As String = "CompareChildNodesByPosition"
    On Error GoTo CompareChildNodesByPositionExit
    
    Dim xmlNewNode As IXMLDOMNode, _
        xmlOldNode As IXMLDOMNode, _
        xmlConfigNode As IXMLDOMNode, _
        xmlNewAttrib As IXMLDOMAttribute, _
        xmlNewElem As IXMLDOMElement, _
        xmlNode As IXMLDOMNode
        
    Dim lngIndex As Long
        
    Set xmlConfigNode = gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & vstrNodeName & "']")
            
    lngIndex = 0
    For Each xmlNewElem In vxmlNewNode.selectNodes(vstrNodeName)

        If vxmlOldNode.selectNodes(vstrNodeName).Item(lngIndex) Is Nothing Then
            xmlNewElem.setAttribute "CRUD_OP", "CREATE"
        Else
            CompareNode xmlNewElem, vxmlOldNode.selectNodes(vstrNodeName).Item(lngIndex), xmlConfigNode
        End If
        
        lngIndex = lngIndex + 1
        
    Next
            
    lngIndex = 0
    For Each xmlNode In vxmlOldNode.selectNodes(vstrNodeName)
    
        If vxmlNewNode.selectNodes(vstrNodeName).Item(lngIndex) Is Nothing Then
            
            If gxmldocConfig.selectSingleNode("omIngestion/default[@noDelete='true']") Is Nothing Then
        
                If gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & vstrNodeName & "'][@noDelete]") Is Nothing Then
            
                    Set xmlNewElem = vxmlNewNode.appendChild(xmlNode.cloneNode(True))
                    xmlNewElem.setAttribute "CRUD_OP", "DELETE"
                    xmlNewElem.setAttribute "_IGNORENOWROWS", "1"
                    
                End If
                
            End If
        
        End If
        
        lngIndex = lngIndex + 1
        
    Next
    
CompareChildNodesByPositionExit:
    
    CheckError cstrMethodName

End Sub
    
Private Sub CompareChildNodesByMatch( _
    ByVal vxmlNewNode As IXMLDOMElement, _
    ByVal vxmlOldNode As IXMLDOMElement, _
    ByVal vstrNodeName As String)
    
    Const cstrMethodName As String = "CompareChildNodesByMatch"
    On Error GoTo CompareChildNodesByMatchExit
    
    Dim xmlConfigNode As IXMLDOMNode, _
        xmlNewNode As IXMLDOMNode, _
        xmlOldNode As IXMLDOMNode, _
        xmlNewAttrib As IXMLDOMAttribute, _
        xmlNewElem As IXMLDOMElement, _
        xmlNode As IXMLDOMNode, _
        xmlAttrib As IXMLDOMAttribute
    
    Dim strMatchAttrib As String, _
        strMatchPath As String, _
        strMatchSecondaryAttrib As String, _
        strXpath As String
        
    Set xmlConfigNode = gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & vstrNodeName & "']")
    
    strMatchAttrib = xmlConfigNode.selectSingleNode("primary/@attribute").Text
            
    For Each xmlNewElem In vxmlNewNode.selectNodes(vstrNodeName)
    
        ' no "PRIMARY" key (typically sequence number)
        If xmlNewElem.Attributes.getNamedItem(strMatchAttrib) Is Nothing Then
            xmlNewElem.setAttribute "CRUD_OP", "CREATE"
        Else
            
            strMatchPath = "[@" & strMatchAttrib & "='" & xmlNewElem.Attributes.getNamedItem(strMatchAttrib).Text & "']"
    
            For Each xmlNode In xmlConfigNode.selectNodes("secondary[@attribute]")
            
                strMatchSecondaryAttrib = xmlNode.Attributes.getNamedItem("attribute").Text
                
                If xmlNewElem.Attributes.getNamedItem(strMatchSecondaryAttrib) Is Nothing Then
                    Err.Raise _
                        oeXMLMissingAttribute, _
                        cstrMethodName, _
                        "missing " & strMatchSecondaryAttrib & " on REQUEST node: " & xmlNewElem.xml
                End If
            
                strMatchPath = _
                    strMatchPath & _
                    "[@" & strMatchSecondaryAttrib & "='" & _
                    xmlNewElem.Attributes.getNamedItem(strMatchSecondaryAttrib).Text & "']"
            Next
        
            strXpath = vstrNodeName & strMatchPath
        
            If vxmlOldNode.selectSingleNode(strXpath) Is Nothing Then
                xmlNewElem.setAttribute "CRUD_OP", "CREATE"
            Else
                CompareNode xmlNewElem, vxmlOldNode.selectSingleNode(strXpath), xmlConfigNode
            End If
        
        End If
    
    Next
    
    For Each xmlOldNode In vxmlOldNode.selectNodes(vstrNodeName)
    
        Set xmlAttrib = xmlOldNode.Attributes.getNamedItem(strMatchAttrib)
    
        If xmlAttrib Is Nothing Then
                
            Err.Raise _
                oeNoFieldItemFound, _
                cstrMethodName, _
                "no match value for " & strMatchAttrib & " on existing data node: " & xmlOldNode.xml
        
        End If
            
        strMatchPath = _
            "[@" & strMatchAttrib & "='" & xmlOldNode.Attributes.getNamedItem(strMatchAttrib).Text & "']"
        
        For Each xmlNode In xmlConfigNode.selectNodes("secondary/@attribute")
        
            strMatchSecondaryAttrib = xmlNode.Text
        
            Set xmlAttrib = xmlOldNode.Attributes.getNamedItem(strMatchSecondaryAttrib)
            
            If xmlAttrib Is Nothing Then
                
                Err.Raise _
                    oeNoFieldItemFound, _
                    cstrMethodName, _
                    "no match value for " & strMatchSecondaryAttrib & " on existing data node: " & xmlOldNode.xml
                    
            End If
            
            strMatchPath = _
                strMatchPath & _
                "[@" & strMatchSecondaryAttrib & "='" & _
                xmlOldNode.Attributes.getNamedItem(strMatchSecondaryAttrib).Text & "']"
        
        Next
            
        strXpath = vstrNodeName & strMatchPath
    
        If vxmlNewNode.selectSingleNode(strXpath) Is Nothing Then
            
            If gxmldocConfig.selectSingleNode("omIngestion/default[@noDelete='true']") Is Nothing Then
        
                If gxmldocConfig.selectSingleNode("omIngestion/match[@entity='" & vstrNodeName & "'][@noDelete]") Is Nothing Then
        
                    Set xmlNewElem = vxmlNewNode.appendChild(xmlOldNode.cloneNode(True))
                    xmlNewElem.setAttribute "CRUD_OP", "DELETE"
                    xmlNewElem.setAttribute "_IGNORENOWROWS", "1"
                    
                End If
                
            End If
        
        End If
        
    Next
    
CompareChildNodesByMatchExit:
    
    CheckError cstrMethodName

End Sub

Private Function IsNewAcceptedQuote( _
    ByVal vxmlNewApplicationFactFindNode As IXMLDOMElement, _
    Optional ByVal vxmlOldApplicationFactFindNode As IXMLDOMElement = Null) _
    As Boolean

    Const cstrMethodName As String = "IsNewAcceptedQuote"
    On Error GoTo IsNewAcceptedQuoteExit
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    If vxmlNewApplicationFactFindNode Is Nothing Then
        Exit Function ' unlikely?
    End If
    
    Dim strNewActive As String
    Dim strNewAccepted As String
    Dim strOldAccepted As String
    
    strNewAccepted = vxmlNewApplicationFactFindNode.Attributes.getNamedItem("ACCEPTEDQUOTENUMBER").Text
    
    Set xmlAttrib = vxmlNewApplicationFactFindNode.Attributes.getNamedItem("ACTIVEQUOTENUMBER")
    If Not xmlAttrib Is Nothing Then
        strNewActive = xmlAttrib.Text
    End If

    If Not vxmlOldApplicationFactFindNode Is Nothing Then
    
        Set xmlAttrib = vxmlOldApplicationFactFindNode.Attributes.getNamedItem("ACCEPTEDQUOTENUMBER")
        If Not xmlAttrib Is Nothing Then
            strOldAccepted = xmlAttrib.Text
        End If
        
    End If
    
    If strNewAccepted <> strOldAccepted Then
        If strNewAccepted = strNewActive Then
            IsNewAcceptedQuote = True
            vxmlNewApplicationFactFindNode.Attributes.removeNamedItem ("ACCEPTEDQUOTENUMBER")
        Else
            Err.Raise oeUnspecifiedError, cstrMethodName, "new ACCEPTEDQUOTENUMBER not equal ACTIVEQUOTENUMBER on pre-processing APPLICATIONFACTFIND node: " & vxmlNewApplicationFactFindNode.xml
        End If
    End If
    
IsNewAcceptedQuoteExit:
    
    CheckError cstrMethodName

End Function

Private Function AnyCrudOp(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    Const cstrMethodName As String = "AnyCrudOp"
    On Error GoTo AnyCrudOpExit

    Dim xmlChildNode As IXMLDOMNode

    If Not vxmlRequestNode.Attributes.getNamedItem("CRUD_OP") Is Nothing Then
        AnyCrudOp = True
    Else
        For Each xmlChildNode In vxmlRequestNode.childNodes
            If AnyCrudOp(xmlChildNode) Then
                AnyCrudOp = True
            End If
        Next
    End If
    
AnyCrudOpExit:
    
    CheckError cstrMethodName

End Function

Private Function IsValidEdit(ByVal vstrNodeName As String) As Boolean
    Const cstrMethodName As String = "IsValidEdit"
    On Error GoTo IsValidEditExit

    Select Case vstrNodeName
    
        ' IK MAR1124
        ' Case "APPLICATIONSTAGE", "CASEACTIVITY", "APPLICATIONPRIORITY", "OTHERCUSTOMERACCOUNTRELATIONSHIP"
        Case "APPLICATIONSTAGE", "CASEACTIVITY", "APPLICATIONPRIORITY"
            IsValidEdit = False
            
        Case Else
            IsValidEdit = True
            
    End Select
    
IsValidEditExit:

    CheckError cstrMethodName

End Function

' IK 14/02/2006 MAR1264
Private Sub CopyThirdPartyAttributes( _
    ByVal vxmlSource As IXMLDOMElement, _
    ByVal vxmlTarget As IXMLDOMElement)
    
    Dim xmlAttrib As IXMLDOMAttribute
    
    For Each xmlAttrib In vxmlSource.Attributes
    
        If xmlAttrib.Name <> "DIRECTORYGUID" _
        And xmlAttrib.Name <> "THIRDPARTYGUID" _
        And xmlAttrib.Name <> "ADDRESSGUID" _
        And xmlAttrib.Name <> "CONTACTDETAILSGUID" _
        And xmlAttrib.Name <> "TELEPHONESEQNUM" _
        Then
        
        vxmlTarget.Attributes.setNamedItem xmlAttrib.cloneNode(True)
        
        End If

    Next
        
End Sub


Private Sub CopyAttribute( _
    ByVal vxmlSource As IXMLDOMElement, _
    ByVal vxmlTarget As IXMLDOMElement, _
    ByVal vstrAttName As String)
    
    If Not vxmlSource.Attributes.getNamedItem(vstrAttName) Is Nothing Then
        vxmlTarget.setAttribute vstrAttName, vxmlSource.getAttribute(vstrAttName)
    End If

End Sub
' IK 14/02/2006 MAR1264 ends
