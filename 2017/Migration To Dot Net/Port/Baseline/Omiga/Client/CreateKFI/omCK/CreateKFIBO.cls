VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CreateKFIBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'header ----------------------------------------------------------------------------------
'Workfile:      CreateKFIBO.cls
'Copyright:     Copyright © 2005 Vertex Financial Services
'
'Description:   Routines for creating KFI's
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'GHun   04/11/2005  MAR301 Created
'TW     09/11/2005  MAR442
'GHun   14/11/2005  MAR490
'GHun   16/11/2005  MAR544 Add OneOffCosts to the response
'JD     15/12/2005  MAR846 add in attributes to the loancomponents to enable call to CalcArrangementFee
'GHun   22/12/2005  MAR846 backed out the previous MAR846 changes as they are no longer required
'PSC    16/01/2006  MAR1054 Use filenet user globalparameter
'BC     22/01/2006  MAR1091 Disposable KFI remove element NONDISPOSABLE, which contains attribute REFERENCE
'GHun   23/01/2006  MAR1062 Added GetFileNetGuidAndUpdateDMS
'GHun   27/01/2006  MAR1021 Added support for running CriticalDataCheck and CreditCheck
'GHun   29/01/2006  MAR1130 Changed CreateNonDisposableKFI to also call StoreQuotation
'BC     01/02/2006  MAR1106 Include QuotationNumber and Creation Date in DOCUMENTNAME
'GHun   03/02/2006  MAR1194 Changed GetGlobalParameters
'RF     06/02/2006  MAR1191 Add document name to fulfilment request
'PSC    07/02/2006  MAR1197 Add ContentType to filenet request
'PE     11/02/2006  MAR1254 Modify ConvertCombos to create MORTGAGEONEOFFCOSTTYPE attribute
'GHun   02/03/2006  MAR1143 Changed CreateNonDisposableKFI and GetDecision
'BC     08/03/2006  MAR1367 Retrieve MortgageLender Contact Details
'GHun   20/03/2006  MAR1407 Changed CreateDisposableKFI to get PDFFontMode from a globalparameter
'PE     18/04/2006  MAR1405 CR 141 - ENHANCEMENT KFI document location in FileNet CE
'PE     19/04/2006  MAR1633 RE: CC141 / 23 Create KFI change to use new GP for Disposable KFI.
'PSC    19/04/2006  MAR1649 Performance improvements
'BC     10/05/2006  MAR1738 For Disposable KFI, retrieve Additional Features and add to MortgageProduct
'bc     15/05/2006  MAR1792 Use omCK to retrieve PROVIDERCONTACTDETAILS and MORTGAGEPRODUCTADDITIONALFEATURES
'BC     15/05/2006  MAR1789 Remove attribute REFERENCE from TEMPLATEDATA fro disposable KFIs
'GHun   10/08/2006  MAR1937 Pass QuotationNumber to GetFileNetGUIDAndUpdateDMS
'PSC    07/11/2006  EP2_41  Add functionality to calculate KFI fees
'IK     29/11/2006  EP2_202 Product Cascade, do not create document
'GHun   11/12/2006  EP2_412 Fix DocumentContents encoding
'GHun   12/12/2006  EP2_445 Changed CreateDisposableKFI to pass correct attribute to CRUD
'PSC    14/12/2006  EP2_507 Changed SaveNonDisposableKFIData to save drawdown
'SR     02/03/2007  EP2_1753 - modified SaveNonDisposableKFIData
'PSC    05/03/2007  EP2_1792 Make one off costs optional in SaveNonDisposableKFIData
'PSC    05/03/2007  EP2_1347 Set CANBEREFUNDED attribute
'INR    20/03/2007  EP2_1977 return REGULATIONINDICATOR for Disposable KFI
'SR     26/03/2007  EP2_1208 modified SaveNonDisposableKFIData.
'INR    02/04/2007  EP2_2192 check it exists before trying to remove it
'PE     09/04/2007  EP2_2177 CreateNonDisposableKFI - Use different template if called from web service.
'INR    12/04/2007  EP2_1994 Various to retrieve introducer info & procfees for disposable KFI
'INR    17/04/2007  EP2_2448 Get some correct values
'INR    18/04/2007  EP2_2486 check it exists before trying to use it
'INR    18/04/2007  EP2_2478 need all validation types for "TypeOfMortgage"
'INR    19/04/2007  EP2_2478 MORTGAGEPRODUCTPORTABLEIND passed thru from storedproc
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private m_objContext As ObjectContext
Private m_xmldocConfig As FreeThreadedDOMDocument40 'MAR1021 GHun

Public Function CreateKFI(ByVal strXmlRequest As String) As String
    Const cstrFunctionName  As String = "CreateKFI"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    
    Dim xmlApp              As IXMLDOMElement
    Dim xmlFileNet          As IXMLDOMElement
    Dim xmlResponse         As IXMLDOMElement
    
    Dim isDisposable        As Boolean
        
    Dim strFileNetGUID      As String

On Error GoTo CreateKFIVbErr

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    xmlResponse.setAttribute "TYPE", "SUCCESS"
    
    xmlDoc.loadXML strXmlRequest
    Set xmlApp = xmlGetMandatoryNode(xmlDoc, "/REQUEST/APPLICATION")
    isDisposable = xmlGetAttributeAsBoolean(xmlApp, "DISPOSABLEKFI")
    
    If isDisposable Then
        ' PSC 08/11/2006 EP2_41
        CreateDisposableKFI xmlDoc.documentElement, xmlResponse
    Else
        ' PSC 08/11/2006 EP2_41
        CreateNonDisposableKFI xmlDoc.documentElement, xmlResponse
    End If
    
    errCheckXMLResponseNode xmlResponse, , True
    
'    Set xmlFileNet = xmlDoc.createElement("DOCUMENTLOCATION")
'    xmlFileNet.setAttribute "DOCUMENTID", strFileNetGUID
'    xmlResponse.appendChild xmlFileNet
    
    'Header inserted in case any of the data contains £ signs
    CreateKFI = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlResponse.xml
    Exit Function
    
CreateKFIVbErr:
    Set xmlApp = Nothing
    Set xmlDoc = Nothing
    Set xmlFileNet = Nothing
    Set xmlResponse = Nothing
    
    CreateKFI = errCreateErrorResponse()
End Function

' PSC 07/11/2006 EP2_41 - Start
Public Function CalculateKFIFees(ByVal strXmlRequest As String) As String
    Const cstrFunctionName  As String = "CalculateKFIFees"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    
    Dim xmlApp              As IXMLDOMElement
    Dim xmlFileNet          As IXMLDOMElement
    Dim xmlResponse         As IXMLDOMElement
    Dim xmlOneOffCostList    As IXMLDOMNode
    
    Dim isDisposable        As Boolean
        
    Dim strFees      As String

On Error GoTo CalculateKFIFeesVbErr

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    xmlResponse.setAttribute "TYPE", "SUCCESS"
    
    xmlDoc.loadXML strXmlRequest
    Set xmlApp = xmlGetMandatoryNode(xmlDoc, "/REQUEST/APPLICATION")
    isDisposable = xmlGetAttributeAsBoolean(xmlApp, "DISPOSABLEKFI")
    
    If isDisposable Then
        Set xmlOneOffCostList = CalculateDisposableKFIFees(xmlDoc.documentElement, xmlResponse)
    Else
        Set xmlOneOffCostList = CalculateNonDisposableKFIFees(xmlDoc.documentElement, xmlResponse)
    End If
    
    xmlResponse.appendChild xmlOneOffCostList
    
    'Header inserted in case any of the data contains £ signs
    CalculateKFIFees = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlResponse.xml
    Exit Function
    
CalculateKFIFeesVbErr:
    Set xmlApp = Nothing
    Set xmlDoc = Nothing
    Set xmlFileNet = Nothing
    Set xmlResponse = Nothing
    
    CalculateKFIFees = errCreateErrorResponse()
End Function
' PSC 07/11/2006 EP2_41 - End

Private Sub CreateDisposableKFI( _
    ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlResponse As IXMLDOMElement)
    
    Const cstrFunctionName      As String = "CreateDisposableKFI"
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim xmlKFIDocIn             As FreeThreadedDOMDocument40
    Dim xmlKFIResponseDoc       As FreeThreadedDOMDocument40
    Dim xmlKFITransformedXML    As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc      As FreeThreadedDOMDocument40
    Dim xmlTransformXSL         As FreeThreadedDOMDocument40
    
    Dim xmlAttribute            As IXMLDOMAttribute
    
    Dim xmlMortgageLenderList   As IXMLDOMNodeList 'MAR1367 BC 08/03/2006
    Dim xmlLoancomponentList    As IXMLDOMNodeList 'BC MAR1738 15/05/2006
    Dim xmlAdditionalFeaturesList As IXMLDOMNodeList 'BC MAR1738 15/05/2006
    
    Dim xmlElem                 As IXMLDOMElement
    Dim xmlKFIElem              As IXMLDOMElement
    
    Dim xmlHelperResponseNode   As IXMLDOMNode
    Dim xmlLTV                  As IXMLDOMNode
    Dim xmlNode                 As IXMLDOMNode
    Dim xmlNodeTemp             As IXMLDOMNode
    Dim xmlNodeTempID           As IXMLDOMNode
    Dim xmlPrintAttributes      As IXMLDOMNode
    Dim xmlTempRequest          As IXMLDOMNode
    Dim xmlItem                 As IXMLDOMNode 'BC MAR1091
    Dim xmlnonDisposable        As IXMLDOMNode 'BC MAR1091
    Dim xmlSchemaNode           As IXMLDOMNode 'BC MAR1367 08/03/2006
    Dim xmlRequestDoc           As FreeThreadedDOMDocument40 'BC MAR1367 08/03/2006
    Dim xmlRequestNode          As IXMLDOMNode 'BC MAR1367 08/03/2006
    Dim xmlMortgageLender       As IXMLDOMNode 'BC MAR1367 08/03/2006
    Dim xmlMortgageProduct      As IXMLDOMNode 'BC MAR1367 08/03/2006
    
    Dim xmlLoanComponent        As IXMLDOMNode 'BC MAR1738 15/05/2006
    Dim xmlAdditionalFeature    As IXMLDOMNode 'BC MAR1738 15/05/2006
    
    ' PSC 08/11/2006 EP2_41 - Start
    Dim xmlCalcs As IXMLDOMNode
    Dim xmlOneOffCosts As IXMLDOMNode
    Dim xmlDocumentNode As IXMLDOMNode
    Dim xmlDocContents As IXMLDOMAttribute
    Dim strContentType As String
    Dim strDocumentId As String
    'EP2_412 GHun
    'Dim arrByteDocumentContents() As Byte
    Dim strDocumentContents As String
    'EP2_412 End
    ' PSC 08/11/2006 EP2_41 - End

    
    Dim strAceXSL               As String
    Dim strAceXML               As String
    Dim strApplicationNumber    As String
    Dim strDeliveryType         As String
    Dim strDPSTemplateId        As String
    Dim strKFIDocType           As String
    Dim strKFIResponse          As String
    Dim strKFIOut               As String
    Dim strOutputType           As String
    Dim strResponse             As String
    Dim strTemplate             As String
    Dim strTemplateData         As String
    Dim strTemplateId           As String
    Dim strUserId               As String
    Dim strXSLFilename          As String
    Dim strDocumentName         As String
    Dim strCreationDate         As String
    Dim strPermanency           As String
    Dim strDocumentClass        As String
    
    Dim objCalcsBO              As omCM.CalcsBO
    Dim objKFI                  As eKFIControl
    Dim objKFIHelp              As OmKFIHelp
    Dim objParamDO              As GlobalParameterDO
    Dim objPrintBO              As omPrintBO
    Dim objTemplateHandlerBO    As TemplateHandlerBO
    Dim omCRUD                  As Object 'BC MAR1367 08/03/2006
    'EP2_1994
    Dim xmlIntroducersList As IXMLDOMNodeList
    Dim xmlIntroducerFeeDetails As IXMLDOMNode
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlMortgageSubquote As IXMLDOMNode
    Dim xmlProcFees As IXMLDOMNode
    Dim xmlMortgageIntroducerFee As IXMLDOMNode
    Dim xmlApplnIntroducer As IXMLDOMNode
    Dim xmlApplnIntroducerList As IXMLDOMNodeList
    Dim xmlApplnIntroducerDetails As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim principalFirmId As String
    Dim arFirmId As String
    Dim ClubNetworkAssocId As String
    Dim IntroducerId As String
    Dim selectString As String
     
On Error GoTo CreateDisposableKFIVbErr
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    'Calculate LTV
    Set xmlLTV = CalcLTV(vXMLRequest)
    
    'Calculate Mortgage Costs
    Set xmlNode = vXMLRequest.cloneNode(False)
    xmlNode.appendChild xmlMakeNodeElementBased(vXMLRequest.firstChild, True, "CALCS")
    Set xmlElem = xmlGetMandatoryNode(xmlNode, "//MORTGAGESUBQUOTE")
    xmlElem.appendChild xmlLTV
    
    Set xmlOneOffCosts = xmlGetNode(xmlElem, "ONEOFFCOSTLIST")
    
    If Not xmlOneOffCosts Is Nothing Then
        Set xmlCalcs = xmlGetMandatoryNode(xmlNode, "CALCS")
        xmlOneOffCosts.parentNode.removeChild xmlOneOffCosts
        xmlCalcs.appendChild xmlOneOffCosts
    End If
    
    If m_objContext Is Nothing Then
        Set objCalcsBO = New omCM.CalcsBO
    Else
        Set objCalcsBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".CalcsBO")
    End If
    
    strResponse = objCalcsBO.CalcMortCalcCosts(xmlNode.xml)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponse, True
    'Set xmlNode = xml
    For Each xmlNode In xmlDoc.documentElement.childNodes()
        vxmlResponse.appendChild xmlNode.cloneNode(True)
        vXMLRequest.appendChild xmlNode.cloneNode(True)
    Next
    Set xmlIntroducersList = vXMLRequest.selectNodes("//APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCERLIST/APPLICATIONINTRODUCER")
    'EP2_1994 Need to get the IntroducerFeeDetails for the MortgageIntroducerFee calc
    If xmlIntroducersList.length > 0 Then
        Set xmlCrudRequest = xmlGetRequestNode(vXMLRequest)
        Set xmlIntroducerFeeDetails = GetDisposableKFIProcFees(xmlIntroducersList, xmlCrudRequest)
        If Not xmlIntroducerFeeDetails Is Nothing Then
            Set xmlMortgageSubquote = vXMLRequest.selectSingleNode("//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE")
            xmlMortgageSubquote.appendChild xmlIntroducerFeeDetails.selectSingleNode("INTRODUCERPROCFEEDETAILS").cloneNode(True)
            'save our ApplicationIntroducer stuff to be used later
            Set xmlApplnIntroducerDetails = xmlIntroducerFeeDetails.selectSingleNode("RESPONSE").cloneNode(True)
            Set xmlProcFees = CalcProcFees(vXMLRequest)
            Set xmlMortgageIntroducerFee = vXMLRequest.selectSingleNode("//CALCS")
            If Not xmlProcFees Is Nothing And Not xmlMortgageIntroducerFee Is Nothing Then
                xmlMortgageIntroducerFee.appendChild xmlProcFees.cloneNode(True)
            End If
        End If
    End If
    
    'Add the Details to our ApplicationIntroducer stuff from Web
    If Not xmlApplnIntroducerDetails Is Nothing Then
        Set xmlApplnIntroducerList = vXMLRequest.selectNodes("//APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCERLIST/APPLICATIONINTRODUCER")
        For Each xmlApplnIntroducer In xmlApplnIntroducerList
            principalFirmId = xmlGetAttributeText(xmlApplnIntroducer, "PRINCIPALFIRMID")
            If Len(principalFirmId) > 0 Then
                'principalfirm could be either a PACKAGER or a NETWORK node
                selectString = "PACKAGER[@PACKAGERID = '" & principalFirmId & "']"
                Set xmlTemp = xmlApplnIntroducerDetails.selectSingleNode(selectString)
                If xmlTemp Is Nothing Then
                    selectString = "NETWORK[@PRINCIPALFIRMNETWORKID = '" & principalFirmId & "']"
                    Set xmlTemp = xmlApplnIntroducerDetails.selectSingleNode(selectString)
                End If
                If Not xmlTemp Is Nothing Then
                    xmlApplnIntroducer.appendChild xmlTemp.cloneNode(True)
                End If
            End If
            arFirmId = xmlGetAttributeText(xmlApplnIntroducer, "ARFIRMID")
            If Len(arFirmId) > 0 Then
                selectString = "ARFIRM[@ARFIRMID = '" & arFirmId & "']"
                Set xmlTemp = xmlApplnIntroducerDetails.selectSingleNode(selectString)
                If Not xmlTemp Is Nothing Then
                    xmlApplnIntroducer.appendChild xmlTemp.cloneNode(True)
                End If
            End If
            ClubNetworkAssocId = xmlGetAttributeText(xmlApplnIntroducer, "CLUBNETWORKASSOCID")
            If Len(ClubNetworkAssocId) > 0 Then
                selectString = "MORTGAGECLUBNETWORKASSOCIATION[@CLUBNETWORKASSOCIATIONID = '" & ClubNetworkAssocId & "']"
                Set xmlTemp = xmlApplnIntroducerDetails.selectSingleNode(selectString)
                If Not xmlTemp Is Nothing Then
                    xmlApplnIntroducer.appendChild xmlTemp.cloneNode(True)
                End If
            End If
            IntroducerId = xmlGetAttributeText(xmlApplnIntroducer, "INTRODUCERID")
            If Len(IntroducerId) > 0 Then
                selectString = "INTRODUCER[@INTRODUCERID = '" & IntroducerId & "']"
                Set xmlTemp = xmlApplnIntroducerDetails.selectSingleNode(selectString)
                If Not xmlTemp Is Nothing Then
                    xmlApplnIntroducer.appendChild xmlTemp.cloneNode(True)
                End If
            End If
        Next xmlApplnIntroducer
    End If
    
    'Create XML for printing KFI
    If m_objContext Is Nothing Then
        Set objParamDO = New GlobalParameterDO
        Set objPrintBO = New omPrintBO
    Else
        Set objParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objPrintBO = m_objContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    End If
    
    ' MAR 1633
    ' Peter Edney - 19/04/2006
    'strTemplateId = CStr(objParamDO.GetCurrentParameterByType("STKFIQuotationTemplateID", "AMOUNT"))
    ' PSC 08/11/2006 EP2_41
    strTemplateId = CStr(objParamDO.GetCurrentParameterByType("DisposableSTKFITemplateID", "STRING"))
    
    'Obtain print attributes for particular KFI
    
    Set xmlTempRequest = vXMLRequest.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
    Set xmlPrintAttributes = xmlDoc.createElement("FINDATTRIBUTES")
    xmlTempRequest.appendChild xmlPrintAttributes
    xmlSetAttributeValue xmlPrintAttributes, "HOSTTEMPLATEID", strTemplateId
    
    strResponse = objPrintBO.omRequest(xmlTempRequest.xml)

    Set xmlTempResponseDoc = xmlLoad(strResponse, cstrFunctionName)

    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, , True
    Set xmlPrintAttributes = Nothing
    Set xmlPrintAttributes = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
    strDeliveryType = xmlGetMandatoryAttributeText(xmlPrintAttributes, "DELIVERYTYPE")
    strDPSTemplateId = xmlGetMandatoryAttributeText(xmlPrintAttributes, "DPSTEMPLATEID")
    
    ' MAR1405
    ' Peter Edney - 18/04/2006
    'strPermanency = xmlGetMandatoryAttributeText(xmlPrintAttributes, "PERMANENCY")
    ' If the document permanency has not been set for the template, default it to
    ' Disposable (2)
    strPermanency = xmlGetAttributeText(xmlPrintAttributes, "PERMANENCY", 2)
    ' PSC 19/04/2006 MAR1649
    strDocumentClass = GetValidationTypeForValueID("DocumentPermanency", strPermanency)
    'strDocumentClass = objCombo.GetFirstComboValidation("DocumentPermanency", strPermanency)
    If strDocumentClass = "" Then
        strDocumentClass = "OutboundMortgageApplicationDocument"
    End If
    
    If Len(strDPSTemplateId) = 0 Then
        errThrowError TypeName(Me), cstrFunctionName, 502
    End If
    
    strOutputType = GetValidationTypeForValueID("DocumentDeliveryType", strDeliveryType)
    
    'Obtain print template for particular KFI

    If m_objContext Is Nothing Then
        Set objTemplateHandlerBO = New TemplateHandlerBO
    Else
        Set objTemplateHandlerBO = m_objContext.CreateInstance(gstrDPS_COMPONENT & ".TemplateHandlerBO")
    End If

    Set xmlTempRequest = vXMLRequest.cloneNode(False)
    Set xmlNodeTemp = xmlTempRequest.appendChild(xmlDoc.createElement("TEMPLATE"))
    Set xmlNodeTempID = xmlNodeTemp.appendChild(xmlDoc.createElement("TEMPLATEID"))
    xmlNodeTempID.Text = strDPSTemplateId
    strResponse = objTemplateHandlerBO.GetTemplate(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, , True

    strAceXML = xmlGetNodeText(xmlTempResponseDoc.documentElement, "//TEMPLATE/ACEXML")
    strAceXSL = xmlGetNodeText(xmlTempResponseDoc.documentElement, "//TEMPLATE/ACEXSL")
    strTemplate = xmlGetNodeText(xmlTempResponseDoc.documentElement, "//TEMPLATE/FILENAME")
        
    vXMLRequest.setAttribute "NOW", Format$(Now(), "DD/MM/YYYY")
    
    Set xmlNode = GetGlobalParameters()
    vXMLRequest.appendChild xmlNode.selectSingleNode(".//GLOBALPARAMETERLIST")
    
    Set xmlTransformXSL = New FreeThreadedDOMDocument40
    xmlTransformXSL.async = False
    xmlTransformXSL.setProperty "NewParser", True
        
' First Transformation
    strXSLFilename = Replace(App.Path, "DLL", "XML", 1, 1, vbTextCompare) & "\" & "BuildKFIData.xslt"

    'Load the xsl file
    If xmlTransformXSL.Load(strXSLFilename) Then
        Set xmlKFITransformedXML = New FreeThreadedDOMDocument40
        xmlKFITransformedXML.async = False
        xmlKFITransformedXML.setProperty "NewParser", True
        
        Set xmlNode = vXMLRequest.cloneNode(True)
        xmlNode.transformNodeToObject xmlTransformXSL.documentElement, xmlKFITransformedXML
    Else
        errThrowError cstrFunctionName, oeSchemaNotLoaded, "Failed to load stylesheet"
    End If
    
    If Not xmlKFITransformedXML Is Nothing Then
        
        xmlTempResponseDoc.setProperty "NewParser", True
        xmlTempResponseDoc.validateOnParse = False
        
        strKFIDocType = "Standard"
        
        If m_objContext Is Nothing Then
            Set objKFIHelp = New OmKFIHelp
        Else
            Set objKFIHelp = m_objContext.CreateInstance("omKFIHelper.OmKFIHelp")
        End If
        ConvertCombos xmlKFITransformedXML.documentElement
        
        'BC MAR1367 08/03/2006 Begin
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
            
        Set xmlRequestNode = xmlRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
        xmlRequestDoc.appendChild xmlRequestNode
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("USERID")
        xmlAttribute.Text = vXMLRequest.getAttribute("USERID")
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("CRUD_OP")
        xmlAttribute.Text = "READ"
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("ENTITY_REF")
        xmlAttribute.Text = "PROVIDERCONTACTDETAILS"
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("SCHEMA_NAME")
        xmlAttribute.Text = "omCK" 'BC MAR1792
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOLOOKUP")
        xmlAttribute.Text = "y"
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOTYPELOOKUP")
        xmlAttribute.Text = "Y"
        xmlRequestNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlSchemaNode = xmlRequestDoc.createNode(NODE_ELEMENT, "PROVIDERCONTACTDETAILS", "")
        xmlRequestNode.appendChild xmlSchemaNode
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("MORTGAGEPRODUCTCODE") 'EP2_445 GHun
        xmlAttribute.Text = vXMLRequest.selectSingleNode("//LOANCOMPONENT/@MORTGAGEPRODUCTCODE").Text
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
        
        Set xmlAttribute = xmlRequestDoc.createAttribute("STARTDATE")
        xmlAttribute.Text = vXMLRequest.selectSingleNode("//LOANCOMPONENT/@STARTDATE").Text
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
        
        Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        strResponse = omCRUD.omRequest(xmlRequestNode.xml)
        
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        
        errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponse, True

        'add Mortgagelender element to xmlKFITransformedXML
        Set xmlMortgageLenderList = xmlDoc.selectNodes("//MORTGAGELENDER")
             
        For Each xmlMortgageLender In xmlMortgageLenderList
            Set xmlNodeTemp = xmlKFITransformedXML.selectSingleNode("//MORTGAGELENDER")
            Set xmlMortgageProduct = xmlKFITransformedXML.selectSingleNode(".//MORTGAGEPRODUCT")
            xmlMortgageProduct.replaceChild xmlMortgageLender, xmlNodeTemp
        Next
        'BC MAR1367 08/03/2006 End
        
        'BC MAR1738 15/05/2006 Begin
        Set xmlLoancomponentList = xmlKFITransformedXML.selectNodes("//LOANCOMPONENT")
        
        For Each xmlLoanComponent In xmlLoancomponentList
            
            Set xmlRequestDoc = New FreeThreadedDOMDocument40
            
            Set xmlRequestNode = xmlRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
            xmlRequestDoc.appendChild xmlRequestNode
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("USERID")
            xmlAttribute.Text = vXMLRequest.getAttribute("USERID")
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("CRUD_OP")
            xmlAttribute.Text = "READ"
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("ENTITY_REF")
            xmlAttribute.Text = "MORTGAGEPRODUCTADDITIONALFEATURES"
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("SCHEMA_NAME")
            xmlAttribute.Text = "omCK" 'BC MAR1792
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOLOOKUP")
            xmlAttribute.Text = "y"
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOTYPELOOKUP")
            xmlAttribute.Text = "Y"
            xmlRequestNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlSchemaNode = xmlRequestDoc.createNode(NODE_ELEMENT, "MORTGAGEPRODUCTADDITIONALFEATURES", "")
            xmlRequestNode.appendChild xmlSchemaNode
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("MORTGAGEPRODUCT")
            xmlAttribute.Text = xmlGetAttributeText(xmlLoanComponent, "MORTGAGEPRODUCTCODE")
            xmlSchemaNode.Attributes.setNamedItem xmlAttribute
            
            Set xmlAttribute = xmlRequestDoc.createAttribute("STARTDATE")
            xmlAttribute.Text = xmlGetAttributeText(xmlLoanComponent, "STARTDATE")
            xmlSchemaNode.Attributes.setNamedItem xmlAttribute
            
            Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
            strResponse = omCRUD.omRequest(xmlRequestNode.xml)
            
            Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
            
            errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponse, True
'EP2_2478 MORTGAGEPRODUCTPORTABLEIND passed thru from storedproc
'            Set xmlNodeTemp = xmlLoanComponent.selectSingleNode(".//MORTGAGEPRODUCT")
'            Set xmlAttribute = xmlRequestDoc.createAttribute("MORTGAGEPRODUCTPORTABLEIND")
'            xmlAttribute.Text = xmlDoc.selectSingleNode _
'                    ("//MORTGAGEPRODUCTADDITIONALFEATURES/MORTGAGEPRODUCT/@MORTGAGEPRODUCTPORTABLEIND").Text
'            xmlNodeTemp.Attributes.setNamedItem xmlAttribute
            
            'add Mortgagelender element to xmlKFITransformedXML
            Set xmlAdditionalFeaturesList = xmlDoc.selectNodes("//ADDITIONALFEATURES")
            
            For Each xmlAdditionalFeature In xmlAdditionalFeaturesList
                Set xmlNodeTemp = xmlLoanComponent.selectSingleNode(".//MORTGAGEPRODUCT")
                xmlNodeTemp.appendChild xmlAdditionalFeature
            Next
        Next
        'BC MAR1738 15/05/2006 End
        
        strKFIOut = objKFIHelp.FormatData(xmlKFITransformedXML.xml, strKFIDocType)
        Set xmlTempResponseDoc = xmlLoad(strKFIOut, cstrFunctionName)
       
        'BC MAR1091 and MAR1367 Begin
        Set xmlItem = xmlTempResponseDoc.selectSingleNode("//TEMPLATEDATA")
        'EP2_2486
        If Not xmlItem Is Nothing Then
            Set xmlnonDisposable = xmlItem.selectSingleNode("//NONDISPOSABLE")
            'EP2_2192
            If Not xmlnonDisposable Is Nothing Then
                xmlItem.removeChild xmlnonDisposable
            End If
        End If
        
        'BC MAR1789 Begin
        Set xmlElem = xmlTempResponseDoc.selectSingleNode("//TEMPLATEDATA")
        'EP2_2486
        If Not xmlElem Is Nothing Then
            xmlElem.removeAttribute "REFERENCE"
        End If
        
        Set xmlItem = xmlTempResponseDoc.selectSingleNode("//SECTION14")
        'EP2_2486
        If Not xmlItem Is Nothing Then
            Set xmlnonDisposable = xmlItem.selectSingleNode("//NONDISPOSABLE")
            'EP2_2192
            If Not xmlnonDisposable Is Nothing Then
                xmlItem.removeChild xmlnonDisposable
            End If
        End If
        
        'KFIHelp.FormatData does not return a response node when successful
        'Consequently errCheckXMLResponseNode will fail. Still need to report
        'The error when it fails.
        Set xmlHelperResponseNode = xmlTempResponseDoc.selectSingleNode("//RESPONSE")
        If Not xmlHelperResponseNode Is Nothing Then
            errCheckXMLResponseNode xmlTempResponseDoc.documentElement, , True
        End If
        Set objKFIHelp = Nothing

        'Direct call to eKFI (this negates the need for normal Omiga printing code)

        If m_objContext Is Nothing Then
            Set objKFI = New eKFIControl
        Else
            Set objKFI = m_objContext.CreateInstance("eKFI.eKFIControl")
        End If

        strTemplateData = xmlTempResponseDoc.selectSingleNode("//TEMPLATEDATA").xml

        Set xmlKFIDocIn = New FreeThreadedDOMDocument40
        xmlKFIDocIn.async = False
        xmlKFIDocIn.setProperty "NewParser", True

        Set xmlKFIElem = xmlKFIDocIn.createElement("REQUEST")
        xmlKFIElem.setAttribute "Template", strTemplate
        xmlKFIElem.setAttribute "Data", strTemplateData
        xmlKFIElem.setAttribute "Output", ""
        xmlKFIElem.setAttribute "AlphaCalcs", "No"
        xmlKFIElem.setAttribute "XSL", strAceXSL
        xmlKFIElem.setAttribute "XML", strAceXML
        xmlKFIElem.setAttribute "OutputType", strOutputType
        xmlKFIElem.setAttribute "DateFormat", "dd/mm/yyyy"
        xmlKFIElem.setAttribute "IntegerFormat", "#,##0"
        xmlKFIElem.setAttribute "NumericFormat", "#,##0.00"
        xmlKFIElem.setAttribute "PDFFontMode", CStr(GetGlobalParamAmount("PDFFontMode"))    'MAR1407 GHun
        xmlKFIElem.setAttribute "PDFSecurity", "3"
        xmlKFIElem.setAttribute "PDFUserPassword", ""
        xmlKFIElem.setAttribute "PDFOwnerPassword", ""
        xmlKFIDocIn.appendChild xmlKFIElem

        strKFIResponse = objKFI.ProcessKFI(xmlKFIDocIn.documentElement.xml)

        Set xmlKFIResponseDoc = New FreeThreadedDOMDocument40
        Set xmlKFIResponseDoc = xmlLoad(strKFIResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlKFIResponseDoc.documentElement, , True

        Set xmlNode = xmlKFIResponseDoc.selectSingleNode("//DOCUMENTCONTENTS")
        Set xmlAttribute = xmlNode.Attributes.getNamedItem("FILECONTENTS")
        'EP2_412 GHun The FILECONTENTS is already bin base64 encoded even though the datatype is not set correctly
        'arrByteDocumentContents = xmlAttribute.nodeTypedValue
        strDocumentContents = xmlAttribute.Text
        'EP2_412 End
        strContentType = xmlGetAttributeText(xmlNode, "CONTENTTYPE")

        ' PSC 16/01/2006 MAR1054
        ' PSC 08/11/2006 EP2_41 - Start
        ' Only call if filenet is the DMS
        If objParamDO.GetCurrentParameterByType("DocumentMangementSystemType", "AMOUNT") = 2 Then

            strUserId = GetMandatoryGlobalParamString("FileNetUserId")
            strApplicationNumber = vXMLRequest.selectSingleNode("//APPLICATION/@APPLICATIONNUMBER").Text
    
            ' MAR1191 Start
            
            strCreationDate = Format$(Now, "yyyymmdd_hhmmss")
            
            strDocumentName = strTemplate & "_" & _
                strApplicationNumber & "__" & _
                strCreationDate ' quotation number is unknown so leave blank
        
        'CreateDisposableKFI = GetFileNetGUID(strApplicationNumber, strUserId, xmlAttribute.Text)
        ' PSC 07/02/2006 MAR1197
            strDocumentId = GetFileNetGUID( _
                strApplicationNumber, strUserId, xmlAttribute.Text, strDocumentName, _
                GetValidationTypeForValueID("DocumentDeliveryType", CSafeInt(strDeliveryType)), _
                strDocumentClass)
        End If
        ' PSC 08/11/2006 EP2_41 - End
        ' MAR1191 End
    
' Second Transformation
        'MAR544 GHun Get OneOffCost descriptions and types
        Set xmlNode = vXMLRequest.selectSingleNode("/REQUEST/ONEOFFCOSTLIST")
        Set xmlNodeTemp = GetDisposableKFIOneOffCosts(xmlNode.xml)
        vXMLRequest.replaceChild xmlNodeTemp, xmlNode
        'MAR544 End
    
        strXSLFilename = Replace(App.Path, "DLL", "XML", 1, 1, vbTextCompare) & "\" & "CreateDisposableKFIResponse.xslt"
    
        'Load the xsl file
        If xmlTransformXSL.Load(strXSLFilename) Then
            Set xmlKFITransformedXML = New FreeThreadedDOMDocument40
            xmlKFITransformedXML.async = False
            xmlKFITransformedXML.setProperty "NewParser", True
            
            Set xmlNode = vXMLRequest.cloneNode(True)
            xmlNode.transformNodeToObject xmlTransformXSL.documentElement, xmlKFITransformedXML
        Else
            errThrowError cstrFunctionName, oeSchemaNotLoaded, "Failed to load stylesheet"
        End If
        
        If Not xmlKFITransformedXML Is Nothing Then
            'Fix date formats
            For Each xmlElem In xmlKFITransformedXML.selectNodes("//LOANCOMPONENTPAYMENTSCHEDULE")
                If Not xmlElem.Attributes.getNamedItem("STARTDATE") Is Nothing Then
                    xmlElem.setAttribute "STARTDATE", Format$(CSafeDate(xmlElem.getAttribute("STARTDATE")), "DD/MM/YYYY HH:NN:SS")
                End If
            Next
            
            For Each xmlElem In xmlKFITransformedXML.selectNodes("//LOANCOMPONENTREDEMPTIONFEE")
                If Not xmlElem.Attributes.getNamedItem("REDEMPTIONFEEPERIODENDDATE") Is Nothing Then
                    xmlElem.setAttribute "REDEMPTIONFEEPERIODENDDATE", Format$(CSafeDate(xmlElem.getAttribute("REDEMPTIONFEEPERIODENDDATE")), "DD/MM/YYYY HH:NN:SS")
                End If
            Next
    
            For Each xmlNode In vxmlResponse.childNodes
                vxmlResponse.removeChild xmlNode
            Next
             
            vxmlResponse.appendChild xmlKFITransformedXML.documentElement.firstChild
        End If
        
        ' PSC 08/11/2006 EP2_41 - Start
        If Len(strDocumentId) > 0 Then
            Set xmlDocumentNode = xmlDoc.createElement("DOCUMENTLOCATION")
            xmlSetAttributeValue xmlDocumentNode, "DOCUMENTID", strDocumentId
            vxmlResponse.appendChild xmlDocumentNode
        End If
        
        'EP2_412 GHun
        'If UBound(arrByteDocumentContents) > 0 Then
        If Len(strDocumentContents) > 0 Then
        'EP2_412 End
            Set xmlDocumentNode = xmlDoc.createElement("DOCUMENTCONTENT")
            Set xmlDocContents = xmlDoc.createAttribute("CONTENT")
            xmlDocContents.Text = strDocumentContents 'EP2_412 GHun xmlAttribute
            xmlDocContents.dataType = "bin.base64"
            'xmlDocContents.nodeTypedValue = arrByteDocumentContents    'EP2_412 GHun
            xmlDocumentNode.Attributes.setNamedItem xmlDocContents
            xmlSetAttributeValue xmlDocumentNode, "CONTENTTYPE", strContentType
            vxmlResponse.appendChild xmlDocumentNode
            
        End If
        ' PSC 08/11/2006 EP2_41 - End

    End If
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateDisposableKFIVbErr:
    Set xmlDoc = Nothing
    Set xmlKFIDocIn = Nothing
    Set xmlKFIResponseDoc = Nothing
    Set xmlKFITransformedXML = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTransformXSL = Nothing

    Set xmlAttribute = Nothing

    Set xmlElem = Nothing
    Set xmlItem = Nothing
    Set xmlnonDisposable = Nothing 'BC MAR1091
    Set xmlKFIElem = Nothing 'BC MAR1091
    Set omCRUD = Nothing 'BC MAR1367 08/03/2006
    Set xmlSchemaNode = Nothing 'BC MAR1367 08/03/2006
    
    ' PSC 08/11/2006 EP2_41 - Start
    Set xmlCalcs = Nothing
    Set xmlOneOffCosts = Nothing
    Set xmlDocumentNode = Nothing
    Set xmlDocContents = Nothing
    ' PSC 08/11/2006 EP2_41 - End
    

    Set xmlHelperResponseNode = Nothing
    Set xmlLTV = Nothing
    Set xmlNode = Nothing
    Set xmlNodeTemp = Nothing
    Set xmlNodeTempID = Nothing
    Set xmlPrintAttributes = Nothing
    Set xmlTempRequest = Nothing
    Set xmlMortgageProduct = Nothing 'BC MAR1367 08/03/2006
    Set xmlMortgageLender = Nothing 'BC MAR1367 08/03/2006
    Set xmlMortgageLenderList = Nothing 'BC MAR1367 08/03/2006
    'EP2_1994
    Set xmlIntroducersList = Nothing
    Set xmlIntroducerFeeDetails = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlMortgageSubquote = Nothing
    Set xmlProcFees = Nothing
    Set xmlMortgageIntroducerFee = Nothing
    Set xmlApplnIntroducer = Nothing
    Set xmlApplnIntroducerList = Nothing
    Set xmlApplnIntroducerDetails = Nothing
    Set xmlTemp = Nothing
    
    Set objCalcsBO = Nothing
    Set objKFI = Nothing
    Set objKFIHelp = Nothing
    Set objParamDO = Nothing
    Set objPrintBO = Nothing
    Set objTemplateHandlerBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub CreateNonDisposableKFI(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlResponse As IXMLDOMElement)
    Const cstrFunctionName      As String = "CreateNonDisposableKFI"
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim xmlKFIResponseDoc       As FreeThreadedDOMDocument40
    
    Dim xmlAttribute            As IXMLDOMAttribute
    
    Dim xmlApp                  As IXMLDOMElement
    Dim xmlElement              As IXMLDOMElement
    Dim xmlMC                   As IXMLDOMElement
    Dim xmlRequest              As IXMLDOMElement
    Dim xmlResponse             As IXMLDOMElement
    Dim xmlCriticalData         As IXMLDOMElement   'MAR1021 GHun
    
    ' PSC 09/11/2006 EP2_41 - Start
    Dim xmlOneOffCosts          As IXMLDOMNode
    Dim xmlDocumentNode         As IXMLDOMNode
    Dim xmlDocContents          As IXMLDOMAttribute
    ' PSC 09/11/2006 EP2_41 - End
    
    Dim xmlNode                 As IXMLDOMNode
    
    Dim strApplicationNumber    As String
    Dim strAppFactFindNumber    As String
    Dim strMSQNumber            As String
    Dim strQuotationNumber      As String
    Dim strResponse             As String
    Dim strDecision             As String   'MAR1021 GHun

    Dim objAQBO                 As IApplicationQuoteBO
    Dim objIQuotationTxBO       As IQuotationTxBO   'MAR1130 GHun
    Dim objMSQDO                As IMortgageSubQuoteDO  'MAR1130 GHun
    Dim strCreationDate         As String           'BC MAR1106 01/02/2006
    Dim strDocumentName         As String           'BC MAR1106 01/02/2006
    Dim strPermanency           As String
    Dim strDocumentClass        As String
    
    ' PSC 09/11/2006 EP2_41 - Start
    Dim strDocumentId           As String
    Dim strContentType          As String
    'EP2_412 GHun
    'Dim arrByteDocumentContents() As Byte
    Dim strDocumentContents     As String
    'EP2_412 End
    ' PSC 09/11/2006 EP2_41 - End
    Dim strClient               As String
    
On Error GoTo CreateNonDisposableKFIVbErr

    'Save the data
    Set xmlCriticalData = SaveNonDisposableKFIData(vXMLRequest)
    
    strCreationDate = Format$(Now, "yyyymmdd_hhmmss") 'BC MAR1106 01/02/2006
    
    strQuotationNumber = xmlGetAttributeText(xmlCriticalData, "QUOTATIONNUMBER")
    strMSQNumber = xmlGetAttributeText(xmlCriticalData, "MORTGAGESUBQUOTENUMBER")
    
    ' PSC 09/11/2006 EP2_41 - Start
    Set xmlOneOffCosts = xmlGetNode(vXMLRequest, "APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/ONEOFFCOSTLIST")
    
    If Not xmlOneOffCosts Is Nothing Then
        Set xmlOneOffCosts = xmlMakeNodeElementBased(xmlOneOffCosts, True, "ONEOFFCOSTLIST")
    End If
    ' PSC 09/11/2006 EP2_41 - End
        
    'Calculate and save a new quote
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlRequest = vXMLRequest.cloneNode(False)
    xmlDoc.appendChild xmlRequest
    Set xmlMC = xmlDoc.createElement("MORTGAGECOSTS")
    xmlRequest.appendChild xmlMC
    
    Set xmlApp = xmlGetMandatoryNode(vXMLRequest, "APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONNUMBER")
    strAppFactFindNumber = "1"
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplicationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAppFactFindNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONDATE")
    xmlElement.Text = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONDATE")
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("CONTEXT")
    xmlElement.Text = "COST MODELLING"
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("QUOTATIONNUMBER")    'MAR490 GHun
    xmlElement.Text = strQuotationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElement.Text = strMSQNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("LIFESUBQUOTENUMBER")
    xmlElement.Text = 0
    xmlMC.appendChild xmlElement
    
    ' PSC 09/11/2006 EP2_41 - Start
    If Not xmlOneOffCosts Is Nothing Then
        xmlMC.appendChild xmlOneOffCosts.cloneNode(True)
    End If
    ' PSC 09/11/2006 EP2_41 - End

    If m_objContext Is Nothing Then
        Set objAQBO = New omAQ.ApplicationQuoteBO
    Else
        Set objAQBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    End If

    Set xmlResponse = objAQBO.CalculateMortgageCosts(xmlDoc.documentElement)
    errCheckXMLResponseNode xmlResponse, , True
    
    'MAR1130 GHun
    'Get the TotalNetMonthlyCost
    Set xmlRequest = vXMLRequest.cloneNode(False)
    Set xmlMC = xmlDoc.createElement("MORTGAGESUBQUOTE")
    xmlRequest.appendChild xmlMC
        
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplicationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAppFactFindNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElement.Text = strMSQNumber
    xmlMC.appendChild xmlElement
    
    If m_objContext Is Nothing Then
        Set objMSQDO = New MortgageSubQuoteDO
        Set objIQuotationTxBO = New QuotationTxBO
    Else
        Set objMSQDO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteDO")
        Set objIQuotationTxBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationTxBO")
    End If
    
    Set xmlResponse = objMSQDO.GetData(xmlRequest)
    
    'Store the quotation
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlRequest = vXMLRequest.cloneNode(False)
    Set xmlMC = xmlDoc.createElement("QUOTATION")
    xmlRequest.appendChild xmlMC
        
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplicationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAppFactFindNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strQuotationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("TOTALQUOTATIONCOST")
    xmlElement.Text = xmlGetMandatoryNodeText(xmlResponse, "TOTALNETMONTHLYCOST")
    xmlMC.appendChild xmlElement
        
    Set xmlResponse = objIQuotationTxBO.StoreQuotation(xmlRequest)
    errCheckXMLResponseNode xmlResponse, , True
    'MAR1130 End
    
    'Accept the quote
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlRequest = vXMLRequest.cloneNode(False)
    xmlDoc.appendChild xmlRequest
    Set xmlMC = xmlDoc.createElement("QUOTATION")
    xmlRequest.appendChild xmlMC
        
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplicationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAppFactFindNumber
    xmlMC.appendChild xmlElement
        
    Set xmlResponse = objAQBO.AcceptQuotation(xmlDoc.documentElement)
    errCheckXMLResponseNode xmlResponse, , True
    
    ' PSC 08/11/2006 EP2_41 Start
    'MAR1021 GHun Perform a CriticalDataCheck if necessary
'    strDecision = vbNullString  'MAR1143 GHun
'    If Not xmlGetNode(xmlCriticalData, "BEFORE") Is Nothing Then
'        strDecision = CriticalDataCheck(vxmlRequest, xmlCriticalData)
'    End If
    'MAR1021 End
    ' PSC 08/11/2006 EP2_41 End

    
    'MAR1143 GHun Don't bother creating the PDF if the decision is a decline
    ' PSC 08/11/2006 EP2_41 Start
'    If Not (strDecision = "3" Or strDecision = "4" Or strDecision = "5") Then
    
    'EP2_202
    If vXMLRequest.Attributes.getNamedItem("NODOC") Is Nothing Then
    
        'Create a KFI
        Dim objPrintBO          As omPrintBO
        Dim xmlPrintAttributes  As IXMLDOMElement
        Dim strTemplateId       As String
        Dim strDeliveryType     As String
        Dim strDPSTemplateId    As String
        'Dim strOutputType       As String
        Dim objPDMBO            As PrintDataManagerBO
        Dim objParamDO          As GlobalParameterDO
        
        If m_objContext Is Nothing Then
            Set objPrintBO = New omPrintBO
            Set objPDMBO = New PrintDataManagerBO
            Set objParamDO = New GlobalParameterDO
        Else
            Set objPrintBO = m_objContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
            Set objPDMBO = m_objContext.CreateInstance("omPDM.PrintDataManagerBO")
            Set objParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        End If
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
    
        Set xmlRequest = vXMLRequest.cloneNode(False)
        xmlRequest.setAttribute "OPERATION", "GetPrintAttributes"
        xmlDoc.appendChild xmlRequest
        
        strClient = UCase(xmlGetAttributeText(vXMLRequest, "CLIENT", ""))
        If strClient = "CREATEKFIWS" Then
            strTemplateId = CStr(objParamDO.GetCurrentParameterByType("STKFIWebQuotationTemplateID", "STRING"))
        Else
        ' PSC 09/11/2006 EP2_41
            strTemplateId = CStr(objParamDO.GetCurrentParameterByType("STKFIQuotationTemplateID", "STRING"))
        End If
        
        Set xmlElement = xmlDoc.createElement("FINDATTRIBUTES")
        xmlElement.setAttribute "HOSTTEMPLATEID", strTemplateId
        xmlRequest.appendChild xmlElement
        
        strResponse = objPrintBO.omRequest(xmlRequest.xml)
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlDoc.documentElement, , True
        
        Set xmlPrintAttributes = xmlGetMandatoryNode(xmlDoc, ".//ATTRIBUTES")
        strDeliveryType = xmlGetMandatoryAttributeText(xmlPrintAttributes, "DELIVERYTYPE")
        strDPSTemplateId = xmlGetMandatoryAttributeText(xmlPrintAttributes, "DPSTEMPLATEID")
        
        ' MAR1405
        ' Peter Edney - 18/04/2006
        ' strPermanency = xmlGetMandatoryAttributeText(xmlPrintAttributes, "PERMANENCY")
        ' If the document permanency has not been set for the template, default it to
        ' Non-Disposable (1).
        strPermanency = xmlGetAttributeText(xmlPrintAttributes, "PERMANENCY", 1)
        ' PSC 19/04/2006 MAR1649
        strDocumentClass = GetValidationTypeForValueID("DocumentPermanency", strPermanency)
        'strDocumentClass = objCombo.GetFirstComboValidation("DocumentPermanency", strPermanency)
        If strDocumentClass = "" Then
            strDocumentClass = "OutboundMortgageApplicationDocument"
        End If
        
        If Len(strDPSTemplateId) = 0 Then
            errThrowError TypeName(Me), cstrFunctionName, 502
        End If
                   
        'strOutputType = GetValidationTypeForValueID("DocumentDeliveryType", strDeliveryType)
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
        
        Set xmlRequest = vXMLRequest.cloneNode(False)
        xmlRequest.setAttribute "OPERATION", "PrintDocument"
        
        xmlDoc.appendChild xmlRequest
        Set xmlMC = xmlDoc.createElement("PRINTDATA")
        xmlMC.setAttribute "METHODNAME", "STKFIQUOTATIONTEMPLATE"
        xmlRequest.appendChild xmlMC
        
        Set xmlApp = xmlGetMandatoryNode(vXMLRequest, "APPLICATION")
        
        xmlMC.setAttribute "APPLICATIONNUMBER", strApplicationNumber
        xmlMC.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFindNumber
        
        Set xmlMC = xmlDoc.createElement("TEMPLATEDATA")
        'xmlMC.setAttribute "MORTGAGESUBQUOTENUMBER", strMSQNumber
        xmlMC.setAttribute "QUOTATIONNUMBER", strQuotationNumber
        xmlRequest.appendChild xmlMC
        
        Set xmlMC = xmlDoc.createElement("DOCUMENTDETAILS")
        xmlMC.setAttribute "DOCUMENTNAME", xmlGetAttributeText(xmlPrintAttributes, "HOSTTEMPLATENAME") _
                            + "_" + strApplicationNumber + "_" + strQuotationNumber + "_" + strCreationDate 'BC MAR1106 01/02/2006
        strDocumentName = xmlGetAttributeText(xmlMC, "DOCUMENTNAME")
        xmlMC.setAttribute "DOCUMENTGROUP", xmlGetAttributeText(xmlPrintAttributes, "TEMPLATEGROUPID")
        xmlRequest.appendChild xmlMC
        
        Set xmlMC = xmlDoc.createElement("CONTROLDATA")
                                
        xmlMC.setAttribute "COPIES", 1
        'xmlMC.setAttribute "PRINTER", m_defaultPrinter
        xmlMC.setAttribute "DOCUMENTID", xmlGetAttributeText(xmlPrintAttributes, "HOSTTEMPLATEID")
        xmlMC.setAttribute "DPSDOCUMENTID", strDPSTemplateId
        xmlMC.setAttribute "DELIVERYTYPE", strDeliveryType
    '    xmlMC.setAttribute "COMPRESSIONMETHOD", "ZLIB"
        xmlMC.setAttribute "TEMPLATEGROUPID", xmlGetAttributeText(xmlPrintAttributes, "TEMPLATEGROUPID")
        xmlMC.setAttribute "HOSTTEMPLATENAME", xmlGetAttributeText(xmlPrintAttributes, "HOSTTEMPLATENAME")
        xmlMC.setAttribute "HOSTTEMPLATEDESCRIPTION", xmlGetAttributeText(xmlPrintAttributes, "HOSTTEMPLATEDESCRIPTION")
        xmlMC.setAttribute "DOCUMENTNAME", xmlGetAttributeText(xmlPrintAttributes, "HOSTTEMPLATENAME")
        xmlMC.setAttribute "VIEWBEFOREPRINT", "1"
        xmlMC.setAttribute "URLPOSTIND", xmlGetAttributeText(xmlPrintAttributes, "URLPOSTIND")
        xmlMC.setAttribute "DESTINATIONTYPE", "L"
        xmlRequest.appendChild xmlMC
        
        strResponse = objPrintBO.omRequest(xmlRequest.xml)
    
        Set xmlKFIResponseDoc = New FreeThreadedDOMDocument40
        Set xmlKFIResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlKFIResponseDoc.documentElement, , True
    
        Set xmlNode = xmlKFIResponseDoc.selectSingleNode("//DOCUMENTCONTENTS")
        Set xmlAttribute = xmlNode.Attributes.getNamedItem("FILECONTENTS")
        'EP2_412 GHun
        'arrByteDocumentContents = xmlAttribute.nodeTypedValue
        strDocumentContents = xmlAttribute.Text
        'EP2_412 End
    
        ' PSC 07/02/2006 MAR1197
        ' PSC 09/11/2006 EP2_41 - Start
        strContentType = GetValidationTypeForValueID("DocumentDeliveryType", CSafeInt(strDeliveryType))
        xmlSetAttributeValue xmlNode, "CONTENTTYPE", strContentType
        ' PSC 09/11/2006 EP2_41 - End
        
        ' MAR1405
        ' Peter Edney - 18/04/2006
        xmlSetAttributeValue xmlNode, "DOCUMENTCLASS", strDocumentClass
       
       'MAR1062 GHun
       'MAR1937 GHun pass strQuotationNumber parameter
        'strDocumentId = GetFileNetGUIDAndUpdateDMS(vxmlRequest, strDocumentName, strApplicationNumber, strTemplateId, strQuotationNumber, xmlNode)
    
        'MAR1143 GHun
    '    End If
        ' PSC 08/11/2006 EP2_41 End
    
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    
        Set xmlResponse = GetNonDisposableKFIData(strApplicationNumber, strAppFactFindNumber, strQuotationNumber, strMSQNumber)
        'MAR1021 GHun Add the underwriters decision if there is one
        If Len(strDecision) > 0 Then
            Set xmlApp = xmlResponse.selectSingleNode("APPLICATION")
            xmlSetAttributeValue xmlApp, "UNDERWRITERSDECISION", strDecision
        End If
        'MAR1021 End
        
        vxmlResponse.appendChild xmlResponse.firstChild
        
        ' PSC 08/11/2006 EP2_41 - Start
        If Len(strDocumentId) > 0 Then
            Set xmlDocumentNode = xmlDoc.createElement("DOCUMENTLOCATION")
            xmlSetAttributeValue xmlDocumentNode, "DOCUMENTID", strDocumentId
            vxmlResponse.appendChild xmlDocumentNode
        End If
        
        'EP2_412 GHun
        'If UBound(arrByteDocumentContents) > 0 Then
        If Len(strDocumentContents) > 0 Then
        'EP2_412 End
            Set xmlDocumentNode = xmlDoc.createElement("DOCUMENTCONTENT")
            Set xmlDocContents = xmlDoc.createAttribute("CONTENT")
            xmlDocContents.Text = strDocumentContents   'EP2_412 GHun
            xmlDocContents.dataType = "bin.base64"
            'xmlDocContents.nodeTypedValue = arrByteDocumentContents    'EP2_412 GHun
            xmlDocumentNode.Attributes.setNamedItem xmlDocContents
            xmlSetAttributeValue xmlDocumentNode, "CONTENTTYPE", strContentType
            vxmlResponse.appendChild xmlDocumentNode
        End If
        ' PSC 08/11/2006 EP2_41 - End
    'EP2_202
    End If

CreateNonDisposableKFIVbErr:

    Set xmlDoc = Nothing
    Set xmlKFIResponseDoc = Nothing

    Set xmlAttribute = Nothing

    Set xmlApp = Nothing
    Set xmlElement = Nothing
    Set xmlMC = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing

    Set xmlNode = Nothing

    Set objAQBO = Nothing
    Set objIQuotationTxBO = Nothing
    Set objMSQDO = Nothing
    ' PSC 09/11/2006 EP2_41 - Start
    Set xmlOneOffCosts = Nothing
    Set xmlDocumentNode = Nothing
    Set xmlDocContents = Nothing
    ' PSC 09/11/2006 EP2_41 - End


    errCheckError cstrFunctionName, TypeName(Me)
End Sub

' PSC 07/02/2006 MAR1197
Private Function GetFileNetGUID(ByVal strApplicationNumber As String, _
                                ByVal strUserId As String, _
                                ByVal strFileContents As String, _
                                ByVal strDocumentName As String, _
                                ByVal strContentType As String, _
                                ByVal strDocumentClass As String) As String
                                
Const cstrFunctionName As String = "GetFileNetGUID"

Dim boolTraceOn As Boolean
Dim objFSO As Object
Dim objOmFDM As Object
Dim objWshShell As Object

Dim strFileNetGUID As String
Dim strPassword As String
Dim strRegValue As String
Dim strTraceFolder As String
    
    On Error Resume Next
    
' Set the trace parameters for omFDM, if set
    Set objOmFDM = CreateObject("omFDM.FileNetInterfaceBO")
    Set objWshShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\")
    If strRegValue = "1" Or strRegValue = "2" Then
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\FOLDER")
        If Len(strRegValue) > 0 Then
            If objFSO.FolderExists(strRegValue) Then
                boolTraceOn = True
                strTraceFolder = strRegValue & "\"
            End If
        End If
    End If
    Set objWshShell = Nothing
    On Error GoTo GetFileNetGUIDError
    
    ' Get the Decrypted password for the user
    strPassword = GetPassword(strUserId)
    
    If objOmFDM Is Nothing Then
        Err.Raise 40018, "omCK.CreateKFIBO." & cstrFunctionName, "Can't create omFDM.FileNetInterfaceBO"
    End If
    
    objOmFDM.traceOn = boolTraceOn
    objOmFDM.traceLocation = strTraceFolder
    
    ' MAR1191 Start
    'strFileNetGUID = objOmFDM.GetFileNetGUID(strApplicationNumber, strFileContents, strUserId, strPassword)
    ' PSC 07/02/2006 MAR1197
    strFileNetGUID = objOmFDM.GetFileNetGUID( _
        strApplicationNumber, strFileContents, strUserId, strPassword, strDocumentName, strContentType, strDocumentClass)
    ' MAR1191 End
    
    GetFileNetGUID = DeformatGUID(strFileNetGUID)
    
    Set objOmFDM = Nothing
    Exit Function
    
GetFileNetGUIDError:
    Err.Raise Err.Number, Err.Source, "Error in CreateKFIBO.GetFileNetGUID: " & Err.Description
    
End Function

'MAR1062 GHun
'MAR1937 GHun include vstrQuotationNumber parameter
Private Function GetFileNetGUIDAndUpdateDMS(ByVal vXMLRequest As IXMLDOMElement, _
                                    ByVal vstrDocumentName As String, _
                                    ByVal vstrApplicationNumber As String, _
                                    ByVal vstrHostTemplateId As String, _
                                    ByVal vstrQuotationNumber As String, _
                                    ByVal vxmlDocumentContents As IXMLDOMNode) As String
    
    Const cstrFunctionName  As String = "GetFileNetGUIDAndUpdateDMS"
    Dim omPack              As PackManagerBO
    Dim strResponse         As String
    Dim strGuid             As String
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement

On Error GoTo GetFileNetGUIDAndUpdateDMSError
    
    If m_objContext Is Nothing Then
        Set omPack = New PackManagerBO
    Else
        Set omPack = m_objContext.CreateInstance("omPack.PackManagerBO")
    End If
    
    Set xmlRequest = vXMLRequest.cloneNode(False)
    xmlRequest.setAttribute "OPERATION", "SENDDOCUMENTTOFILENET"
    xmlRequest.setAttribute "DOCUMENTNAME", vstrDocumentName
    xmlRequest.setAttribute "APPLICATIONNUMBER", vstrApplicationNumber
    xmlRequest.setAttribute "HOSTTEMPLATEID", vstrHostTemplateId
    xmlRequest.setAttribute "QUOTATIONNUMBER", vstrQuotationNumber  'MAR1937 GHun
    xmlRequest.appendChild vxmlDocumentContents.cloneNode(True)
    
    strResponse = omPack.omRequest(xmlRequest.xml)
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    strGuid = xmlGetMandatoryAttributeText(xmlDoc.documentElement, "FILENETIMAGEREF")
    
    GetFileNetGUIDAndUpdateDMS = strGuid
    
    Set omPack = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Exit Function
GetFileNetGUIDAndUpdateDMSError:
    Err.Raise Err.Number, Err.Source, "Error in CreateKFIBO." & cstrFunctionName & ": " & Err.Description
End Function
'MAR1062 End

Private Function FormatGuid(ByVal strGuid As String) As String
    Dim strFormattedGUID As String
    strFormattedGUID = "{" & Left$(strGuid, 8) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 9, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 13, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 17, 4) & "-"
    strFormattedGUID = strFormattedGUID & Mid$(strGuid, 21, 12) & "}"
    FormatGuid = strFormattedGUID
End Function

Private Function DeformatGUID(ByVal strGuid As String) As String
    DeformatGUID = Replace(Replace(Replace(strGuid, "-", ""), "{", ""), "}", "")
End Function

Private Function GetPassword(ByVal strUserId As String) As String
    Dim connOmiga As Connection
    
    Dim recPassword As ADODB.Recordset
    
    Dim strSQL As String
    
    On Error GoTo GetPasswordError
    
    strSQL = "select PASSWORDVALUE from PASSWORD where USERID = '" & strUserId & "' " _
    & " AND PASSWORDCREATIONDATE = (SELECT MAX(PASSWORDCREATIONDATE) FROM PASSWORD " _
    & " where USERID = '" & strUserId & "')"
    
    Set connOmiga = New Connection
    connOmiga.ConnectionString = adoGetDbConnectString
    connOmiga.open
    
    Set recPassword = connOmiga.Execute(strSQL)
    
    If recPassword.EOF Then
        Err.Raise 40022, "PackControl.GetPassword", "Can't find password for user name '" & strUserId & "'"
    End If
    
    GetPassword = Decrypt(recPassword(0))
    recPassword.Close
    connOmiga.Close
    
TidyUp:
    Set connOmiga = Nothing
    Set recPassword = Nothing
    Exit Function
    
GetPasswordError:
    App.LogEvent "PackControl.GetPassword " & Err.Description, vbLogEventTypeError
    Resume TidyUp
End Function

Private Function Decrypt(ByVal strPassword As String) As String
    Const cint_MIN = 10
    Const cint_MAX = 40
    
    Dim lngChar As Long
    Dim intIncrement As Integer
    Dim strOut As String
    
    Dim strChar As String
    
    ' initialise the random-number generator
    Rnd -1
    Randomize 1
    
    For lngChar = 1 To Len(strPassword)
        intIncrement = Int(Rnd() * cint_MAX) + cint_MIN
    
        strChar = Mid$(strPassword, lngChar, 1)
        If strChar = "~" Then
            lngChar = lngChar + 1
            strChar = Mid$(strPassword, lngChar, 1)
        Else
            intIncrement = 0 - intIncrement
        End If
    
        strOut = strOut & Chr(Asc(strChar) + intIncrement)
    Next

    Decrypt = strOut
        
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

'********************************************************************************
'** Function:       ConvertCombos
'** Created by:     Andy Woodhouse
'** Date:           24/06/2004
'** Description:    Converts combo ID's in the XML to combo validation ID's and combo text.
'** Parameters:     vxmlNode - the XML to convert.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Private Sub ConvertCombos(ByVal vxmlNode As IXMLDOMNode)

    Const cstrFunctionName  As String = "ConvertCombos"
      
    Dim xmlList             As IXMLDOMNodeList
    Dim xmlItem             As IXMLDOMNode
    Dim xmlCustomerVersion  As IXMLDOMNode
   
    Dim strText             As String
    Dim strValidation       As String
    Dim strValueID          As String
    
    ' PSC 19/04/2006 MAR1649 - Start
    Dim colValidationTypes As Collection
    Dim varValidationType As Variant
    ' PSC 19/04/2006 MAR1649 - End
      
    '*-select the APPLICATION node in the node and convert the
    '*-TYPEOFBUYER values to their appropriate validation type ID's
    Set xmlItem = vxmlNode.selectSingleNode("//APPLICATION")
    If xmlAttributeValueExists(xmlItem, "TYPEOFBUYER") Then
        strValueID = xmlGetAttributeText(xmlItem, "TYPEOFBUYER")
        strText = GetComboText("TypeOfBuyerNewLoan", strValueID)
        xmlSetAttributeValue xmlItem, "TYPEOFBUYER_TEXT", strText
        ' PSC 19/04/2006 MAR1649
        strValidation = GetValidationTypeForValueID("TypeOfBuyerNewLoan", strValueID)
        'strValidation = objComboDO.GetFirstComboValidation("TypeOfBuyerNewLoan", strValueID)
        xmlSetAttributeValue xmlItem, "TYPEOFBUYER", strValidation
    End If
    
    'EP2_1977 moved below to go thru list once
    '*-select all the APPLICATIONFACTFIND nodes in the node and convert the
    '*-TYPEOFAPPLICATION values to their appropriate validation type ID's
'    Set xmlList = vxmlNode.selectNodes("//APPLICATIONFACTFIND")
'    For Each xmlItem In xmlList
'        If xmlAttributeValueExists(xmlItem, "TYPEOFAPPLICATION") Then
'            strValueID = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
'            strText = GetComboText("TypeOfMortgage", strValueID)
'            xmlSetAttributeValue xmlItem, "TYPEOFAPPLICATION_TEXT", strText
'            ' PSC 19/04/2006 MAR1649 - Start
'            If IsValidationType("TypeOfMortgage", strValueID, "N") Then
'                strValidation = "N"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "F") Then
'                strValidation = "F"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "R") Then
'                strValidation = "R"
'            End If
'            ' PSC 19/04/2006 MAR1649 - End
'            'strValidation = objComboDO.GetFirstComboValidation("TypeOfMortgage", strValueID)
'            xmlSetAttributeValue xmlItem, "TYPEOFAPPLICATION", strValidation
'        End If
'    Next xmlItem
    
    'EP2_1977 moved below to go thru list once
    '*-select all the LOANCOMPONENT nodes in the node and convert the
    '*-LOANCOMPONENT values to their appropriate validation type ID's
'    Set xmlList = vxmlNode.selectNodes("//LOANCOMPONENT")
'    For Each xmlItem In xmlList
'        If xmlAttributeValueExists(xmlItem, "PURPOSEOFLOAN") Then
'            strValueID = xmlGetAttributeText(xmlItem, "PURPOSEOFLOAN")
'            ' PSC 19/04/2006 MAR1649 - Start
'            If IsValidationType("TypeOfMortgage", strValueID, "N") Then
'                strValidation = "N"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "F") Then
'                strValidation = "F"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "R") Then
'                strValidation = "R"
'            End If
'            ' PSC 19/04/2006 MAR1649 - End
'            'strValidation = objComboDO.GetFirstComboValidation("TypeOfMortgage", strValueID)
'            xmlSetAttributeValue xmlItem, "PURPOSEOFLOAN", strValidation & ","
'        End If
'    Next xmlItem

    '*-select all the CUSTOMERROLE nodes in the node and convert the
    '*-CUSTOMERROLETYPE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//CUSTOMERROLE")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "CUSTOMERROLETYPE") Then
            strValueID = xmlGetAttributeText(xmlItem, "CUSTOMERROLETYPE")
            ' PSC 19/04/2006 MAR1649
            strValidation = GetValidationTypeForValueID("CustomerRoleType", strValueID)
            'strValidation = objComboDO.GetFirstComboValidation("CustomerRoleType", strValueID)
            xmlSetAttributeValue xmlItem, "CUSTOMERROLETYPE", strValidation & ","
        End If
        
        'BC MAR1818 Begin
        Set xmlCustomerVersion = xmlItem.selectSingleNode(".//CUSTOMERVERSION")
        If xmlAttributeValueExists(xmlCustomerVersion, "TITLE") Then
            strValueID = xmlGetAttributeText(xmlCustomerVersion, "TITLE")
            strText = GetComboText("Title", strValueID)
            xmlSetAttributeValue xmlCustomerVersion, "TITLE_TEXT", strText
        End If
        'BC MAR1818 End
    Next xmlItem
    
    '*-select all the MORTGAGEONEOFFCOST nodes in the node and convert the
    '*-MORTGAGEONECOSTTYPE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//MORTGAGEONEOFFCOST")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "MORTGAGEONEOFFCOSTTYPE") Then
            strValidation = xmlGetAttributeText(xmlItem, "MORTGAGEONEOFFCOSTTYPE")
            ' PSC 19/04/2006 MAR1649
            strValueID = GetFirstComboValueId("OneOffCost", strValidation)
            strText = GetComboText("OneOffCost", strValueID)
            xmlSetAttributeValue xmlItem, "MORTGAGEONEOFFCOSTTYPE_TEXT", strText
            
            ' PSC 19/04/2006 MAR1649
            Set colValidationTypes = New Collection
            strValidation = vbNullString
 
            GetValidationTypesForValueID "OneOffCost", strValueID, colValidationTypes
            
            For Each varValidationType In colValidationTypes
                If Len(strValidation) > 0 Then
                    strValidation = strValidation & ","
                End If
                strValidation = strValidation & varValidationType
            Next
             
            Set colValidationTypes = Nothing
            ' PSC 19/04/2006 MAR1649 - End
           xmlSetAttributeValue xmlItem, "MORTGAGEONEOFFCOSTTYPE", strValidation & ","
        End If
    Next xmlItem
    
    '*-select all the LOANCOMPONENT nodes in the node and convert the
    '*-REPAYMENTMETHOD values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//LOANCOMPONENT")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "REPAYMENTMETHOD") Then
            strValueID = xmlGetAttributeText(xmlItem, "REPAYMENTMETHOD")
            ' PSC 19/04/2006 MAR1649
            strValidation = GetValidationTypeForValueID("RepaymentType", strValueID)
            'strValidation = objComboDO.GetFirstComboValidation("RepaymentType", strValueID)
            xmlSetAttributeValue xmlItem, "REPAYMENTMETHOD", strValidation & ","
            strText = GetComboText("RepaymentType", strValueID)
            xmlSetAttributeValue xmlItem, "REPAYMENTMETHOD_TEXT", strText
        End If
        'EP2_1977
        If xmlAttributeValueExists(xmlItem, "PURPOSEOFLOAN") Then
            strValueID = xmlGetAttributeText(xmlItem, "PURPOSEOFLOAN")
            ' PSC 19/04/2006 MAR1649 - Start
            If IsValidationType("TypeOfMortgage", strValueID, "N") Then
                strValidation = "N"
            ElseIf IsValidationType("TypeOfMortgage", strValueID, "F") Then
                strValidation = "F"
            ElseIf IsValidationType("TypeOfMortgage", strValueID, "R") Then
                strValidation = "R"
            End If
            ' PSC 19/04/2006 MAR1649 - End
            'strValidation = objComboDO.GetFirstComboValidation("TypeOfMortgage", strValueID)
            xmlSetAttributeValue xmlItem, "PURPOSEOFLOAN", strValidation & ","
        End If

    Next xmlItem

    '*-select all the APPLICATIONFACTFIND nodes in the node and convert the
    '*-LEVELOFADVICE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//APPLICATIONFACTFIND")
    For Each xmlItem In xmlList
        'EP2_1977 moved from above to only go thru list once
        If xmlAttributeValueExists(xmlItem, "TYPEOFAPPLICATION") Then
            strValueID = xmlGetAttributeText(xmlItem, "TYPEOFAPPLICATION")
            strText = GetComboText("TypeOfMortgage", strValueID)
            xmlSetAttributeValue xmlItem, "TYPEOFAPPLICATION_TEXT", strText
            
            'EP2_2478 need all validation types
            Set colValidationTypes = New Collection
            GetValidationTypesForValueID "TypeOfMortgage", strValueID, colValidationTypes

            For Each varValidationType In colValidationTypes
                strValidation = strValidation & varValidationType & ","
            Next
            
            Set colValidationTypes = Nothing
            
            ' PSC 19/04/2006 MAR1649 - Start
'            If IsValidationType("TypeOfMortgage", strValueID, "N") Then
'                strValidation = "N"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "F") Then
'                strValidation = "F"
'            ElseIf IsValidationType("TypeOfMortgage", strValueID, "R") Then
'                strValidation = "R"
'            End If
            ' PSC 19/04/2006 MAR1649 - End
            'strValidation = objComboDO.GetFirstComboValidation("TypeOfMortgage", strValueID)
            xmlSetAttributeValue xmlItem, "TYPEOFAPPLICATION", strValidation
        End If
        If xmlAttributeValueExists(xmlItem, "LEVELOFADVICE") Then
            strValueID = xmlGetAttributeText(xmlItem, "LEVELOFADVICE")
            ' PSC 19/04/2006 MAR1649
            strValidation = GetValidationTypeForValueID("LevelOfAdvice", strValueID)
            'strValidation = objComboDO.GetFirstComboValidation("LevelOfAdvice", strValueID)
            xmlSetAttributeValue xmlItem, "LEVELOFADVICE", strValidation & ","
            strText = GetComboText("LevelOfAdvice", strValueID)
            xmlSetAttributeValue xmlItem, "LEVELOFADVICE_TEXT", strText
        End If

        If xmlAttributeValueExists(xmlItem, "SPECIALGROUP") Then
            strValueID = xmlGetAttributeText(xmlItem, "SPECIALGROUP")
            ' PSC 19/04/2006 MAR1649 - Start
            strValidation = vbNullString
            
            Set colValidationTypes = New Collection
            GetValidationTypesForValueID "SpecialGroup", strValueID, colValidationTypes

            For Each varValidationType In colValidationTypes
                strValidation = strValidation & varValidationType & ","
            Next
            ' PSC 19/04/2006 MAR1649 - End
            Set colValidationTypes = Nothing

            xmlSetAttributeValue xmlItem, "SPECIALGROUP", strValidation
            strText = GetComboText("SpecialGroup", strValueID)
            xmlSetAttributeValue xmlItem, "SPECIALGROUP_TEXT", strText
        End If
        
        'EP2_1977
        If xmlAttributeValueExists(xmlItem, "REGULATIONINDICATOR") Then
            strValueID = xmlGetAttributeText(xmlItem, "REGULATIONINDICATOR")
            strValidation = GetValidationTypeForValueID("RegulationIndicator", strValueID)
            xmlSetAttributeValue xmlItem, "REGULATIONINDICATOR", strValidation & ","
            strText = GetComboText("RegulationIndicator", strValueID)
            xmlSetAttributeValue xmlItem, "REGULATIONINDICATOR_TEXT", strText
        End If

    Next xmlItem
    
    '*-select all the MORTGAGEPRODUCTEMPLOYMENT nodes in the node and convert the
    '*-MORTGAGEPRODUCTEMPLOYMENT values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//MORTGAGEPRODUCTEMPLOYMENT")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "EMPLOYMENTSTATUS") Then
            strValueID = xmlGetAttributeText(xmlItem, "EMPLOYMENTSTATUS")
            strText = GetComboText("EmploymentStatus", strValueID)
            xmlSetAttributeValue xmlItem, "EMPLOYMENTSTATUS_TEXT", strText
        End If
    Next xmlItem
    
    '*-select all the MORTGAGEPRODUCTPROPLOCATION nodes in the node and convert the
    '*-CMORTGAGEPRODUCTPROPLOCATION values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//MORTGAGEPRODUCTPROPLOCATION")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "MORTGAGEPRODUCTLOCATIONTYPE") Then
            strValueID = xmlGetAttributeText(xmlItem, "MORTGAGEPRODUCTLOCATIONTYPE")
            strText = GetComboText("PropertyLocation", strValueID)
            xmlSetAttributeValue xmlItem, "MORTGAGEPRODUCTLOCATIONTYPE_TEXT", strText
        End If
    Next xmlItem

    '*-select all the CONTACTDETAILS nodes in the node and convert the
    '*-CONTACTTITLE values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//CONTACTDETAILS")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "CONTACTTITLE") Then
            strValueID = xmlGetAttributeText(xmlItem, "CONTACTTITLE")
            strText = GetComboText("ContactTitle", strValueID)
            xmlSetAttributeValue xmlItem, "CONTACTTITLE_TEXT", strText
        End If
    Next xmlItem
    
    '*-select the ADDRESS nodes in the node APPLICATIONFACTFIND and convert the
    '*-COUNTRY values to their appropriate validation type ID's
    Set xmlList = vxmlNode.selectNodes("//ADDRESS")
    For Each xmlItem In xmlList
        If xmlAttributeValueExists(xmlItem, "COUNTRY") Then
            strValueID = xmlGetAttributeText(xmlItem, "COUNTRY")
            strText = GetComboText("Country", strValueID)
            xmlSetAttributeValue xmlItem, "COUNTRY_TEXT", strText
        End If
    Next xmlItem
       
    '*-select the MORTGAGEPRODUCT node and convert the
    '*-FLEXIBLEOPTION value to appropriate validation type ID's
    Set xmlItem = vxmlNode.selectSingleNode("//MORTGAGEPRODUCT")
    
    If xmlAttributeValueExists(xmlItem, "FLEXIBLEOPTION") Then
        strValueID = xmlGetAttributeText(xmlItem, "FLEXIBLEOPTION")
        ' PSC 19/04/2006 MAR1649
        strValidation = GetValidationTypeForValueID("FlexibleOptions", strValueID)
        'strValidation = objComboDO.GetFirstComboValidation("FlexibleOptions", strValueID)
        xmlSetAttributeValue xmlItem, "FLEXIBLEOPTION_VALIDID", strValidation & ","
    End If

    Set xmlList = Nothing
    Set xmlItem = Nothing
    
End Sub

' header ----------------------------------------------------------------------------------
' description:  Call the GlobalParameterBO.FindCUrrentParameterList and return the
'
' pass:
' return:       List of globalparameters requested.
'------------------------------------------------------------------------------------------
Private Function GetGlobalParameters() As IXMLDOMNode
    Const strFunctionName = "GetGlobalParameters"
'    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strRequest As String
    Dim strResponse As String
    Dim objGlobal As GlobalParameterBO
    
    On Error GoTo GetGlobalParametersErr
    
    'Append global parameters as required
    'MAR1194 GHun Added MaximumLTVallowed
    strRequest = "<REQUEST><GLOBALPARAMETER><NAME>KFIDaysValidFor</NAME></GLOBALPARAMETER>" _
                & "<GLOBALPARAMETER><NAME>KFIMDay</NAME></GLOBALPARAMETER>" _
                & "<GLOBALPARAMETER><NAME>IsKFIDebugEnabled</NAME></GLOBALPARAMETER>" _
                & "<GLOBALPARAMETER><NAME>KFIDebugPath</NAME></GLOBALPARAMETER>" _
                & "<GLOBALPARAMETER><NAME>MaximumLTVallowed</NAME></GLOBALPARAMETER>" _
                & "</REQUEST>"
        
    If m_objContext Is Nothing Then
        Set objGlobal = New GlobalParameterBO
    Else
        Set objGlobal = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    End If
       
    strResponse = objGlobal.FindCurrentParameterList(strRequest)
    'TK 17/11/2004 BBG1672 errCheckXMLResponse strResponse, True
     
    'Set xmlDoc = xmlCreateDOMObject
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    Set GetGlobalParameters = xmlDoc.documentElement.parentNode
    
    Set xmlDoc = Nothing
    Set objGlobal = Nothing

Exit Function

GetGlobalParametersErr:

    Set xmlDoc = Nothing
    Set objGlobal = Nothing
    Err.Raise Err.Number, Err.Source, "Error in PostProcessing.GetGlobalParameters: " & Err.Description

End Function

Private Function SaveNonDisposableKFIData(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
    Const cstrFunctionName      As String = "SaveNonDisposableKFIData"
    
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoMSQNumberParam       As ADODB.Parameter
    Dim adoQuotationNumberParam As ADODB.Parameter
    Dim adoStream               As ADODB.Stream 'MAR1021 GHun
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    
    Dim xmlApp                  As IXMLDOMNode
    Dim xmlAppFF                As IXMLDOMNode
    Dim xmlLTV                  As IXMLDOMNode
    Dim xmlOneOffCostList       As IXMLDOMNodeList  ' PSC 09/11/2006 EP2_41
    Dim xmlOneOffCost           As IXMLDOMNode      ' PSC 09/11/2006 EP2_41

    Dim xmlElement              As IXMLDOMElement
    
    Dim lngValue                As Long
    Dim strValue                As String
    Dim lngTotalAddedCosts      As Long             ' PSC 09/11/2006 EP2_41
    Dim dblDrawdown             As Double           ' PSC 14/12/2006 EP2_507

On Error GoTo SaveNonDisposableKFIDataExit
    
    'Calculate LTV
    Set xmlLTV = CalcLTV(vXMLRequest)
    
    Set xmlApp = xmlGetMandatoryNode(vXMLRequest, "/REQUEST/APPLICATION")
    Set xmlAppFF = xmlGetMandatoryNode(xmlApp, "APPLICATIONFACTFIND")
    
    ' PSC 09/11/2006 EP2_41 - Start
    ' Calculate added costs
    lngTotalAddedCosts = 0
    
    Set xmlOneOffCostList = xmlAppFF.selectNodes("QUOTATION/MORTGAGESUBQUOTE/ONEOFFCOSTLIST/ONEOFFCOST[@ADDTOLOAN='1']")
    
    For Each xmlOneOffCost In xmlOneOffCostList
        lngTotalAddedCosts = lngTotalAddedCosts + xmlGetAttributeAsLong(xmlOneOffCost, "AMOUNT", "0")
    Next
    ' PSC 09/11/2006 EP2_41 - End
    
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.open
    
    'MAR1021 GHun
    Set adoStream = New ADODB.Stream
    adoStream.open
    'MAR1021 End
    
    Set adoCommand = New ADODB.Command
    
    'Create the parameters and call the stored proc
    With adoCommand
        Set .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "USP_SaveNonDisposableKFIData"
        'MAR1021 GHun
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        'MAR1021 End
        
        'Input parameters
        strValue = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONNUMBER")
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(strValue), strValue)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , 1)
        strValue = Format$(xmlGetMandatoryAttributeAsDate(xmlApp, "APPLICATIONDATE"), "DD/MM/YYYY HH:NN:SS")
        .Parameters.Append .CreateParameter("@p_ApplicationDate", adDate, adParamInput, , strValue)

        lngValue = xmlGetMandatoryAttributeText(xmlAppFF, "TYPEOFAPPLICATION")
        .Parameters.Append .CreateParameter("@p_TypeOfApplication", adInteger, adParamInput, , lngValue)
        
        lngValue = xmlGetMandatoryAttributeAsLong(xmlAppFF, "PURCHASEPRICE")
        .Parameters.Append .CreateParameter("@p_PurchasePrice", adInteger, adParamInput, , lngValue)

        lngValue = xmlGetMandatoryAttributeAsLong(xmlAppFF, "LEVELOFADVICE")
        .Parameters.Append .CreateParameter("@p_LevelOfAdvice", adInteger, adParamInput, , lngValue)
    
        lngValue = xmlGetMandatoryAttributeAsLong(xmlAppFF, "REGULATIONINDICATOR")
        .Parameters.Append .CreateParameter("@p_RegulationIndicator", adInteger, adParamInput, , lngValue)
        
        lngValue = xmlGetMandatoryAttributeAsLong(xmlApp, "TYPEOFVALUATION")
        .Parameters.Append .CreateParameter("@p_TypeOfValuation", adInteger, adParamInput, , lngValue)
    
        lngValue = xmlGetMandatoryAttributeAsLong(xmlAppFF, "LOCATION")
        .Parameters.Append .CreateParameter("@p_Location", adInteger, adParamInput, , lngValue)
        
        Set xmlElement = xmlGetMandatoryNode(xmlAppFF, "QUOTATION/MORTGAGESUBQUOTE")
        lngValue = xmlGetMandatoryAttributeAsLong(xmlElement, "AMOUNTREQUESTED")
        .Parameters.Append .CreateParameter("@p_AmountRequested", adInteger, adParamInput, , lngValue)
        
        ' PSC 14/12/2006 EP2_507
        dblDrawdown = xmlGetAttributeAsDouble(xmlElement, "DRAWDOWN")
        
        Set xmlElement = xmlGetMandatoryNode(xmlElement, "LOANCOMPONENTLIST")
        strValue = xmlElement.xml
        .Parameters.Append .CreateParameter("@p_LoanComponentXML", adBSTR, adParamInput, Len(strValue), strValue)
        
        .Parameters.Append .CreateParameter("@p_LTV", adDouble, adParamInput, , CSafeDbl(xmlLTV.Text))
        
        ' PSC 09/11/2006 EP2_41
        .Parameters.Append .CreateParameter("@p_AddedCosts", adInteger, adParamInput, , lngTotalAddedCosts)
        
        ' PSC 14/12/2006 EP2_507
        .Parameters.Append .CreateParameter("@p_Drawdown", adDouble, adParamInput, , dblDrawdown)
        
        'SR 02/03/2007 : EP2_1753 - new parameter @p_OneOffCostListXML
        ' PSC 05/03/2007 EP2_1792 - Start
        Set xmlElement = xmlGetNode(xmlAppFF, "QUOTATION/MORTGAGESUBQUOTE/ONEOFFCOSTLIST")
        If Not xmlElement Is Nothing Then
            strValue = xmlElement.xml
            .Parameters.Append .CreateParameter("@p_OneOffCostListXML", adBSTR, adParamInput, Len(strValue), strValue)
        Else
            .Parameters.Append .CreateParameter("@p_OneOffCostListXML", adBSTR, adParamInput, 0, Null)
        End If
        ' PSC 05/03/2007 EP2_1792 - End

        'SR 26/03/2007 : EP2_1208
        strValue = xmlGetAttributeText(xmlApp, "PROCFEETOCUST")
        If Len(strValue) > 0 Then
            .Parameters.Append .CreateParameter("@p_ProcFeeToCust", adInteger, adParamInput, , CSafeInt(strValue))
        Else
            .Parameters.Append .CreateParameter("@p_ProcFeeToCust", adInteger, adParamInput, , Null)
        End If
        
        strValue = xmlGetAttributeText(xmlApp, "PROCFEETOCUSTAMOUNT")
        If Len(strValue) > 0 Then
            .Parameters.Append .CreateParameter("@p_ProcFeeToCustAmount", adDouble, adParamInput, , CSafeDbl(strValue))
        Else
            .Parameters.Append .CreateParameter("@p_ProcFeeToCustAmount", adDouble, adParamInput, , 0)
        End If
        'SR 26/03/2007 : EP2_1208 - End
        
        'Output parameters
        Set adoQuotationNumberParam = .CreateParameter("@p_QuotationNumber", adInteger, adParamOutput)
        .Parameters.Append adoQuotationNumberParam

        Set adoMSQNumberParam = .CreateParameter("@p_MortgageSubQuoteNumber", adInteger, adParamOutput)
        .Parameters.Append adoMSQNumberParam
        
        .Execute , , adExecuteStream 'MAR1021 GHun
    End With
            
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    'MAR1021 GHun
    If xmlDoc.loadXML(adoStream.ReadText) Then
        Set xmlElement = xmlDoc.documentElement
    Else
        Set xmlElement = xmlDoc.createElement("RESPONSE")
    End If
    'MAR1021 End
    
    xmlElement.setAttribute "QUOTATIONNUMBER", CSafeLng(adoQuotationNumberParam.Value)
    xmlElement.setAttribute "MORTGAGESUBQUOTENUMBER", CSafeLng(adoMSQNumberParam.Value)
       
    'Close the database connection
    adoStream.Close 'MAR1021 GHun
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    Set SaveNonDisposableKFIData = xmlElement

SaveNonDisposableKFIDataExit:

    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoQuotationNumberParam = Nothing
    Set adoMSQNumberParam = Nothing
    Set adoStream = Nothing
    
    Set xmlDoc = Nothing
    Set xmlApp = Nothing
    Set xmlAppFF = Nothing
    Set xmlElement = Nothing
    Set xmlOneOffCostList = Nothing  ' PSC 09/11/2006 EP2_41
    Set xmlOneOffCost = Nothing      ' PSC 09/11/2006 EP2_41

    
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'Calculate LTV
Private Function CalcLTV(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
    Const cstrFunctionName  As String = "CalcLTV"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    
    Dim xmlLTV              As IXMLDOMNode
    Dim xmlNode             As IXMLDOMNode
    Dim xmlRequest          As IXMLDOMNode
    Dim xmlTempNode         As IXMLDOMNode
    
    Dim objLTVBO            As LTVBO
    
    Dim strResponse         As String
    
On Error GoTo CalcLTVVbErr

    Set xmlRequest = vXMLRequest.cloneNode(False)

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlLTV = xmlDoc.createElement("LTV")
    xmlRequest.appendChild xmlLTV
    
    Set xmlTempNode = xmlGetMandatoryNode(vXMLRequest, "/REQUEST/APPLICATION/APPLICATIONFACTFIND")

    Set xmlNode = xmlDoc.createElement("PURCHASEPRICE")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlTempNode, "PURCHASEPRICE")
    xmlLTV.appendChild xmlNode
    
    Set xmlTempNode = xmlGetMandatoryNode(xmlTempNode, "QUOTATION/MORTGAGESUBQUOTE")
 
    Set xmlNode = xmlDoc.createElement("AMOUNTREQUESTED")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlTempNode, "AMOUNTREQUESTED")
    xmlLTV.appendChild xmlNode

    If m_objContext Is Nothing Then
        Set objLTVBO = New LTVBO
    Else
        Set objLTVBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".LTVBO")
    End If
    
    strResponse = objLTVBO.CalcMortCalcLTV(xmlRequest.xml)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set CalcLTV = xmlGetMandatoryNode(xmlDoc, "/RESPONSE/LTV")
    
CalcLTVVbErr:
    Set xmlDoc = Nothing
    
    Set xmlLTV = Nothing
    Set xmlNode = Nothing
    Set xmlRequest = Nothing
    Set xmlTempNode = Nothing
    
    Set objLTVBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function

Private Function GetNonDisposableKFIData(ByVal vstrApplicationNumber As String, _
                                        ByVal vstrAppFactFindNumber As String, _
                                        ByVal vstrQuotationNumber As String, _
                                        ByVal vstrMortgageSubQuoteNumber As String) As IXMLDOMNode
    
    Const cstrFunctionName      As String = "GetNonDisposableKFIData"
    
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoStream               As ADODB.Stream
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    
    Dim xmlElement              As IXMLDOMElement
    Dim xmlParent               As IXMLDOMElement
    Dim xmlTemp                 As IXMLDOMElement

On Error GoTo GetNonDisposableKFIDataExit
        
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.open
    
    Set adoStream = New ADODB.Stream
    adoStream.open
    
    Set adoCommand = New ADODB.Command
    
    'Create the parameters and call the stored proc
    With adoCommand
        Set .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetNondisposableKFIData"
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        
        'Input parameters
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(vstrApplicationNumber), vstrApplicationNumber)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , CSafeLng(vstrAppFactFindNumber))
        .Parameters.Append .CreateParameter("@p_QuotationNumber", adInteger, adParamInput, , CSafeLng(vstrQuotationNumber))
        .Parameters.Append .CreateParameter("@p_MortgageSubQuoteNumber", adInteger, adParamInput, , CSafeLng(vstrMortgageSubQuoteNumber))
        
        .Execute , , adExecuteStream
    End With
            
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True

    If xmlDoc.loadXML(adoStream.ReadText) Then
        'Fix XML structure
        Set xmlParent = xmlGetMandatoryNode(xmlDoc, "/RESPONSE/APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE")
        Set xmlTemp = xmlDoc.createElement("LOANCOMPONENTLIST")
        xmlParent.appendChild xmlTemp

        For Each xmlParent In xmlDoc.selectNodes("/RESPONSE/LOANCOMPONENT")
            xmlTemp.appendChild xmlParent
        Next
        
        Set xmlParent = xmlGetMandatoryNode(xmlDoc, "/RESPONSE/APPLICATION")
        For Each xmlElement In xmlDoc.selectNodes("/RESPONSE/APPLICATIONFEETYPE")
            xmlParent.appendChild xmlElement
        Next
        
        'MAR544 GHun
        For Each xmlElement In xmlDoc.selectNodes("/RESPONSE/ONEOFFCOST")
            xmlParent.appendChild xmlElement
        Next
        'MAR544 End
    End If
    
    'Close the database connection
    adoStream.Close
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    Set GetNonDisposableKFIData = xmlDoc.documentElement

GetNonDisposableKFIDataExit:

    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlParent = Nothing
    Set xmlTemp = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'MAR544 GHun
Private Function GetDisposableKFIOneOffCosts(ByVal vstrOneOffCostsXml As String) As IXMLDOMNode
    
    Const cstrFunctionName      As String = "GetDisposableKFIOneOffCosts"
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoStream               As ADODB.Stream
    Dim xmlDoc                  As FreeThreadedDOMDocument40

On Error GoTo GetDisposableKFIOneOffCostsExit
        
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.open
    
    Set adoStream = New ADODB.Stream
    adoStream.open
    
    Set adoCommand = New ADODB.Command
    
    'Create the parameters and call the stored proc
    With adoCommand
        Set .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetDisposableKFIOneOffCosts"
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "ONEOFFCOSTLIST"
        
        .Parameters.Append .CreateParameter("@p_XML", adBSTR, adParamInput, Len(vstrOneOffCostsXml), vstrOneOffCostsXml)
        
        .Execute , , adExecuteStream
    End With
            
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True

    xmlDoc.loadXML (adoStream.ReadText)
    
    'Close the database connection
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    Set GetDisposableKFIOneOffCosts = xmlDoc.documentElement

GetDisposableKFIOneOffCostsExit:

    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set xmlDoc = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'MAR544 End

'MAR1021 GHun
Private Function CriticalDataCheck(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlCriticalData As IXMLDOMElement) As String
    
    Const cstrFunctionName As String = "CriticalDataCheck"
        
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim xmlRequest As IXMLDOMElement
    Dim xmlParam As IXMLDOMElement
    
    Dim objCDRules As Object
    
    Dim strResponse As String
    Dim strTaskId As String
    Dim strOperation As String
    Dim strDecision As String

On Error GoTo CriticalDataCheckExit

    strDecision = vbNullString

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
   
    Set xmlRequest = vXMLRequest.cloneNode(False)
    For Each xmlElem In vxmlCriticalData.childNodes()
        xmlRequest.appendChild xmlElem
    Next
    
    'Create GlobalParameter nodes with dummy taskIds. The real globalParameters
    'cannot be used because the web versions of the tasks need to be used instead.
    Set xmlParam = xmlDoc.createElement("GLOBALPARAMETER")
    xmlParam.setAttribute "NAME", "TMReprocessCreditCheck"
    xmlParam.setAttribute "STRING", "REPROCESS"
    xmlRequest.appendChild xmlParam
    
    Set xmlParam = xmlDoc.createElement("GLOBALPARAMETER")
    xmlParam.setAttribute "NAME", "TMRescoreCreditCheck"
    xmlParam.setAttribute "STRING", "RESCORE"
    xmlRequest.appendChild xmlParam
                
    xmlRequest.setAttribute "CREDITCHECK_TEST", "1"
        
    If Not m_objContext Is Nothing Then
        Set objCDRules = m_objContext.CreateInstance("omCDRules.omCDRulesBO")
    Else
        Set objCDRules = CreateObject("omCDRules.omCDRulesBO")
    End If
    
    If objCDRules Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, "error creating Critical Data rules component: omCDRules.omCDRulesBO"
    End If
    
    strResponse = objCDRules.RunRules(xmlRequest.xml)
    xmlDoc.loadXML strResponse
    
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set xmlElem = xmlDoc.selectSingleNode("RESPONSE/CASETASK")
    If Not xmlElem Is Nothing Then
        'Map the dummy taskIds back to the correct operation
        strTaskId = xmlGetMandatoryAttributeText(xmlElem, "TASKID")
        If strTaskId = "REPROCESS" Then
            strOperation = "RunXMLReprocessCreditCheck"
        ElseIf strTaskId = "RESCORE" Then
            strOperation = "RunXMLRescoreCreditCheck"
        End If
        strDecision = GetDecision(vXMLRequest, strOperation)
    End If
    
    CriticalDataCheck = strDecision
    
CriticalDataCheckExit:
    
    Set objCDRules = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequest = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 End

'MAR1021 GHun
Private Function GetDecision(ByVal vXMLRequest As IXMLDOMElement, ByVal vstrOperation As String) As String
    
    Const cstrFunctionName As String = "GetDecision"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim strResponse As String
    Dim strDecision As String
    Dim strTaskId As String
    Dim strStageId As String    'MAR1143 GHun
    
    Dim lngErr As Long
    
On Error GoTo GetDecisionExit
    
    If Len(vstrOperation) > 0 Then
    
        LoadConfigFile
    
        strTaskId = GetTaskId(vstrOperation)
        
        strStageId = CreateTask(vXMLRequest, strTaskId) 'MAR1143 GHun
    
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.setProperty "NewParser", True
        xmlDoc.async = False
        Set xmlNode = xmlDoc.appendChild(vXMLRequest.cloneNode(False))
        
        xmlNode.setAttribute "OPERATION", vstrOperation
        xmlNode.setAttribute "NOADDRESSTARGETING", "1"
        
        Set xmlElem = xmlDoc.createElement("APPLICATION")
        xmlElem.setAttribute "APPLICATIONNUMBER", vXMLRequest.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
        Set xmlNode = xmlNode.appendChild(xmlElem)
        
        Dim objTm As Object
        If Not m_objContext Is Nothing Then
            Set objTm = m_objContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
        Else
            Set objTm = CreateObject(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
        End If
        
        strResponse = objTm.OmTmRequest(xmlDoc.xml)
        Set objTm = Nothing
        
        xmlDoc.loadXML strResponse
        If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/DECISION") Is Nothing Then
            lngErr = 0
            Set xmlElem = xmlDoc.selectSingleNode("RESPONSE[@TYPE='APPERR']/ERROR/NUMBER")
            If Not xmlElem Is Nothing Then
                If Len(xmlElem.Text) > 0 Then
                    lngErr = errGetOmigaErrorNumber(CLng(xmlElem.Text))
                End If
            End If
            
            If lngErr <> 8540 Then
                errCheckXMLResponseNode xmlDoc.documentElement, , True
            End If
        Else
        
            strDecision = xmlDoc.selectSingleNode("RESPONSE/DECISION").Text
            
            CompleteCreditCheckTask vXMLRequest, strTaskId
            
            'MAR1143 GHun
            If vstrOperation = "RunXMLRescoreCreditCheck" And (Not (strDecision = "3" Or strDecision = "4" Or strDecision = "5")) Then
                CreateSubsequentTasks vXMLRequest, vstrOperation, strStageId
            End If
            'MAR1143 End
                
        End If
        
        GetDecision = strDecision
    End If
    
GetDecisionExit:
    
    Set objTm = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 GHun

'MAR1021 GHun
'MAR1143 GHun Return the StageId
Private Function CreateTask(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vstrTaskId As String) As String
    
    Const cstrFunctionName As String = "CreateTask"
    
    Dim xmlCaseActivityDoc As FreeThreadedDOMDocument40
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlApplicationPriority As IXMLDOMElement
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strStageId As String    'MAR1143 GHun
    Dim objTm As Object
    
On Error GoTo CreateTaskExit
       
    Set xmlCaseActivityDoc = New FreeThreadedDOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlApplicationPriority = xmlCaseActivityDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY")
    Set xmlCaseActivity = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivity.selectSingleNode("CASESTAGE")
    strStageId = xmlCaseStage.getAttribute("STAGEID")   'MAR1143 GHun
    
    If Not OutstandingTaskExists(xmlCaseStage, vstrTaskId) Then
    
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
        xmlRequestDoc.async = False
        xmlRequestDoc.setProperty "NewParser", True
        
        Set xmlElem = xmlRequestDoc.createElement("REQUEST")
        xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
        xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
        xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
        xmlElem.setAttribute "OPERATION", "CreateAdhocCaseTask"
        Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("CASETASK")
        xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseActivity.getAttribute("SOURCEAPPLICATION")
        xmlElem.setAttribute "CASEID", xmlCaseActivity.getAttribute("CASEID")
        xmlElem.setAttribute "ACTIVITYID", xmlCaseActivity.getAttribute("ACTIVITYID")
        xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseActivity.getAttribute("ACTIVITYINSTANCE")
        xmlElem.setAttribute "STAGEID", strStageId
        xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
        xmlElem.setAttribute "TASKID", vstrTaskId
        xmlNode.appendChild xmlElem
        
        Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
        xmlElem.setAttribute "APPLICATIONPRIORITY", xmlApplicationPriority.getAttribute("APPLICATIONPRIORITYVALUE")
        xmlNode.appendChild xmlElem
    
        Set xmlResponseDoc = New FreeThreadedDOMDocument40
        xmlResponseDoc.async = False
        xmlResponseDoc.setProperty "NewParser", True
    
        If Not m_objContext Is Nothing Then
            Set objTm = m_objContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
        Else
            Set objTm = CreateObject(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
        End If
        xmlResponseDoc.loadXML objTm.OmTmRequest(xmlRequestDoc.xml)
        
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
            
    End If
    
    CreateTask = strStageId 'MAR1143 GHun
    
CreateTaskExit:
        
    Set objTm = Nothing
    
    Set xmlCaseActivityDoc = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlApplicationPriority = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 End

'MAR1021 GHun
Private Sub GetCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseDoc As FreeThreadedDOMDocument40)
    
    Const cstrFunctionName As String = "GetCaseActivity"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim crudObj As Object
    
On Error GoTo GetCaseActivityExit
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONPRIORITY")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "CASEID", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASESTAGE")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASETASK")
    xmlNode.appendChild xmlElem
    
    If Not m_objContext Is Nothing Then
        Set crudObj = m_objContext.CreateInstance("omCRUD.omCRUDBO")
    Else
        Set crudObj = CreateObject("omCRUD.omCRUDBO")
    End If
    
    If Not vxmlResponseDoc.loadXML(crudObj.omRequest(xmlDoc.xml)) Then
        errThrowError cstrFunctionName, oeXMLParserError, vxmlResponseDoc.parseError.reason
    End If
        
    If vxmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        errThrowError cstrFunctionName, oeUnspecifiedError, vxmlResponseDoc.xml
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving APPLICATIONPRIORITY details"
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGE") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving CASEACTIVITY details"
    End If
        
GetCaseActivityExit:
    Set crudObj = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'MAR1021 End

'MAR1021 GHun
Private Function CompleteCreditCheckTask(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vstrTaskId As String) As Boolean
    
    Const cstrFunctionName As String = "CompleteCreditCheckTask"
    
    Dim xmlCaseActivityDoc As FreeThreadedDOMDocument40
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCaseStage As IXMLDOMElement
    Dim xmlCaseTask As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim objTm As Object
       
On Error Resume Next
       
    Set xmlCaseActivityDoc = New FreeThreadedDOMDocument40
    xmlCaseActivityDoc.async = False
    xmlCaseActivityDoc.setProperty "NewParser", True
    
    GetCaseActivity vxmlRequestNode, xmlCaseActivityDoc
    
    Set xmlCaseActivity = xmlCaseActivityDoc.selectSingleNode("RESPONSE/CASEACTIVITY")
    Set xmlCaseStage = xmlCaseActivity.selectSingleNode("CASESTAGE")
    
    If OutstandingTaskExists(xmlCaseStage, vstrTaskId) Then
    
        Set xmlCaseTask = GetOutstandingTask(xmlCaseStage, vstrTaskId)
    
        Set xmlRequestDoc = New FreeThreadedDOMDocument40
        xmlRequestDoc.async = False
        xmlRequestDoc.setProperty "NewParser", True
        
        Set xmlElem = xmlRequestDoc.createElement("REQUEST")
        xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
        xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
        xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
        xmlElem.setAttribute "OPERATION", "UpdateCaseTask"
        Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlRequestDoc.createElement("CASETASK")
        xmlElem.setAttribute "SOURCEAPPLICATION", xmlCaseActivity.getAttribute("SOURCEAPPLICATION")
        xmlElem.setAttribute "CASEID", xmlCaseActivity.getAttribute("CASEID")
        xmlElem.setAttribute "ACTIVITYINSTANCE", xmlCaseActivity.getAttribute("ACTIVITYINSTANCE")
        xmlElem.setAttribute "ACTIVITYID", xmlCaseActivity.getAttribute("ACTIVITYID")
        xmlElem.setAttribute "STAGEID", xmlCaseStage.getAttribute("STAGEID")
        xmlElem.setAttribute "CASESTAGESEQUENCENO", xmlCaseStage.getAttribute("CASESTAGESEQUENCENO")
        xmlElem.setAttribute "TASKID", xmlCaseTask.getAttribute("TASKID")
        xmlElem.setAttribute "TASKINSTANCE", xmlCaseTask.getAttribute("TASKINSTANCE")
        xmlElem.setAttribute "CASEACTIVITYGUID", xmlCaseTask.getAttribute("CASEACTIVITYGUID")
        xmlElem.setAttribute "TASKSTATUS", "40"
        xmlNode.appendChild xmlElem
    
        Set xmlResponseDoc = New FreeThreadedDOMDocument40
        xmlResponseDoc.async = False
        xmlResponseDoc.setProperty "NewParser", True
    
        If Not m_objContext Is Nothing Then
            Set objTm = m_objContext.CreateInstance("msgTm.msgTmBO")
        Else
            Set objTm = CreateObject("msgTm.msgTmBO")
        End If
        xmlResponseDoc.loadXML objTm.TmRequest(xmlRequestDoc.xml)
        
        If Not xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            CompleteCreditCheckTask = True
        End If
            
    End If
    
    Set objTm = Nothing
    
    Set xmlCaseActivityDoc = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlCaseStage = Nothing
    Set xmlCaseTask = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing

End Function
'MAR1021 End

'MAR1021 GHun
Private Function OutstandingTaskExists(ByVal vxmlCaseStage As IXMLDOMNode, ByVal vstrTaskId As String) As Boolean
    
    Const cstrFunctionName As String = "OutstandingTaskExists"
    Dim strXpath As String
        
On Error GoTo OutstandingTaskExistsExit
        
    strXpath = "CASETASK[@TASKID='" & vstrTaskId & "'][@TASKSTATUS='10' or @TASKSTATUS='20']"
    
    If vxmlCaseStage.selectNodes(strXpath).length > 0 Then
        OutstandingTaskExists = True
    End If
    
OutstandingTaskExistsExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 End

'MAR1021 GHun
Private Function GetOutstandingTask(ByVal vxmlCaseStage As IXMLDOMNode, ByVal vstrTaskId As String) As IXMLDOMNode
    
    Const cstrFunctionName As String = "GetOutstandingTask"
    Dim strXpath As String
    
On Error GoTo GetOutstandingTaskExit
        
    strXpath = "CASETASK[@TASKID='" & vstrTaskId & "'][@TASKSTATUS='10' or @TASKSTATUS='20']"
    
    Set GetOutstandingTask = vxmlCaseStage.selectSingleNode(strXpath)
    
GetOutstandingTaskExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 End

'MAR1021 GHun
Private Function GetTaskId(ByVal vstrCreditCheckOp As String) As String
    
    Const cstrFunctionName As String = "GetTaskId"
    
On Error GoTo GetTaskIdExit
    
    If m_xmldocConfig.selectSingleNode("omIngestion/operation[@name='" & vstrCreditCheckOp & "'][@task]") Is Nothing Then
        Err.Raise oeXMLMissingAttribute, cstrFunctionName, "no taskid in ingestionConfig.xml for credit check operation " & vstrCreditCheckOp
    Else
        GetTaskId = m_xmldocConfig.selectSingleNode("omIngestion/operation[@name='" & vstrCreditCheckOp & "'][@task]").Attributes.getNamedItem("task").Text
    End If
    
GetTaskIdExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MAR1021 End

'MAR1021 GHun
Private Sub LoadConfigFile()

    Const cstrFunctionName As String = "LoadConfigFile"
    On Error GoTo LoadConfigFileExit
    
    Dim strFileSpec As String
    
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\ingestionConfig.xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set m_xmldocConfig = New FreeThreadedDOMDocument40
    m_xmldocConfig.async = False
    m_xmldocConfig.setProperty "NewParser", True
    If Not m_xmldocConfig.Load(strFileSpec) Then
        errThrowError cstrFunctionName, oeXMLParserError, m_xmldocConfig.parseError.reason
    End If
    
LoadConfigFileExit:
    
    If Err.Number <> 0 Then
        If Err.Source = cstrFunctionName Then
            Err.Raise Err.Number, Err.Source, Err.Description
        Else
            Err.Raise Err.Number, cstrFunctionName & "." & Err.Source, Err.Description
        End If
    End If

End Sub
'MAR1021 End

'MAR1143 GHun
Private Sub CreateSubsequentTasks(ByVal vXMLRequest As IXMLDOMElement, ByVal vstrOperation As String, ByVal vstrCurrentStageId As String)

    Const cstrFunctionName          As String = "CreateSubsequentTasks"
    Dim msgTmBO                     As Object
    Dim xmlDoc                      As FreeThreadedDOMDocument40
    Dim xmlRequest                  As IXMLDOMElement
    Dim xmlElem                     As IXMLDOMElement
    Dim strResponse                 As String
    Dim strAiPDataCaptureStageId    As String
    Dim intCurrentSeqNo             As Integer
    Dim intAiPDataCaptureSeqNo      As Integer
    
On Error GoTo CreateSubsequentTasksExit

    Set xmlElem = xmlGetMandatoryNode(m_xmldocConfig.documentElement, "operation[@name='AiPDataCapture']")
    strAiPDataCaptureStageId = xmlGetMandatoryAttributeText(xmlElem, "stage")
    
    If Not m_objContext Is Nothing Then
        Set msgTmBO = m_objContext.CreateInstance("MsgTM.MsgTmBO")
    Else
        Set msgTmBO = CreateObject("MsgTM.MsgTmBO")
    End If
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    'Get the stage list
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlRequest.setAttribute "OPERATION", "GetStageList"
    
    Set xmlElem = xmlDoc.createElement("STAGE")
    'xmlElem.setAttribute "ACTIVITYID", strActivityId
    xmlElem.setAttribute "DELETEFLAG", "0"
    xmlRequest.appendChild xmlElem
        
    strResponse = msgTmBO.TmRequest(xmlRequest.xml)
    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set xmlElem = xmlGetNode(xmlDoc.documentElement, "STAGE[@STAGEID='" & vstrCurrentStageId & "']")
    intCurrentSeqNo = xmlGetAttributeAsInteger(xmlElem, "STAGESEQUENCENO")
    
    Set xmlElem = xmlGetNode(xmlDoc.documentElement, "STAGE[@STAGEID='" & strAiPDataCaptureStageId & "']")
    intAiPDataCaptureSeqNo = xmlGetAttributeAsInteger(xmlElem, "STAGESEQUENCENO")

    'Check that stage is at least AiPDataCapture
    If intCurrentSeqNo >= intAiPDataCaptureSeqNo Then
    
        For Each xmlElem In m_xmldocConfig.documentElement.selectNodes("operation[@name='" & vstrOperation & "']/task")
            CreateTask vXMLRequest, xmlElem.getAttribute("id")
        Next
    End If

CreateSubsequentTasksExit:
    Set msgTmBO = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlElem = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'MAR1143 End

' PSC 07/11/2006 EP2_41 - Start
Private Function CalculateDisposableKFIFees( _
    ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlResponse As IXMLDOMElement) As IXMLDOMNode
    
    Const cstrFunctionName As String = "CalculateDisposableKFIFees"
    
    Dim xmlDoc  As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlLTV  As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlMortgageLender As IXMLDOMNode
    
    Dim xmlOneOffCost As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objCalcsBO As CalcsBO
    
    Dim strIdentifier As String
    Dim strNodeName As String
    Dim blnCanAddToLoan As Boolean
    Dim strResponse As String
        

On Error GoTo CalculateDisposableKFIFeesVbErr
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Dim xmlCalcFeesOnly As IXMLDOMElement
    Dim xmlCalcsNode As IXMLDOMNode
    
    'Calculate LTV
    Set xmlLTV = CalcLTV(vXMLRequest)
    
    'Calculate Mortgage Costs
    Set xmlNode = vXMLRequest.cloneNode(False)
    xmlNode.appendChild xmlMakeNodeElementBased(vXMLRequest.firstChild, True, "CALCS")
    Set xmlElem = xmlGetMandatoryNode(xmlNode, "//MORTGAGESUBQUOTE")
    xmlElem.appendChild xmlLTV
    
    Set xmlCalcFeesOnly = xmlDoc.createElement("CALCULATEFEESONLY")
    xmlCalcFeesOnly.Text = "1"
    
    Set xmlCalcsNode = xmlGetMandatoryNode(xmlNode, "CALCS")
    xmlCalcsNode.appendChild xmlCalcFeesOnly
    
    If m_objContext Is Nothing Then
        Set objCalcsBO = New omCM.CalcsBO
    Else
        Set objCalcsBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".CalcsBO")
    End If
    
    strResponse = objCalcsBO.CalcMortCalcCosts(xmlNode.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    
    Set xmlNode = xmlResponseDoc.selectSingleNode("//RESPONSE/ONEOFFCOSTLIST")
    Set xmlMortgageLender = xmlResponseDoc.selectSingleNode("//RESPONSE/CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGELENDER")
    
    Set CalculateDisposableKFIFees = FormatCalculatedKFIFees(xmlNode, xmlMortgageLender)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateDisposableKFIFeesVbErr:
    Set xmlDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlLTV = Nothing
    Set xmlNode = Nothing
    Set xmlMortgageLender = Nothing
    
    Set xmlElem = Nothing
    Set objCalcsBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Function


Private Function CalculateNonDisposableKFIFees(ByVal vXMLRequest As IXMLDOMElement, ByVal vxmlResponse As IXMLDOMElement) As IXMLDOMNode
    Const cstrFunctionName      As String = "CalculateNonDisposableKFIFees"
    
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    
    
    Dim xmlApp                  As IXMLDOMElement
    Dim xmlElement              As IXMLDOMElement
    Dim xmlMC                   As IXMLDOMElement
    Dim xmlRequest              As IXMLDOMElement
    Dim xmlResponse             As IXMLDOMElement
    Dim xmlCriticalData         As IXMLDOMElement
    Dim xmlOneOffCosts          As IXMLDOMNode
    Dim xmlMortgageLender       As IXMLDOMNode
   
    Dim strApplicationNumber    As String
    Dim strAppFactFindNumber    As String
    Dim strMSQNumber            As String
    Dim strQuotationNumber      As String

    Dim objAQBO                 As IApplicationQuoteBO
    Dim strCreationDate         As String

On Error GoTo CalculateNonDisposableKFIFeesVbErr

    'Save the data
    Set xmlCriticalData = SaveNonDisposableKFIData(vXMLRequest)
    
    strCreationDate = Format$(Now, "yyyymmdd_hhmmss")
    
    strQuotationNumber = xmlGetAttributeText(xmlCriticalData, "QUOTATIONNUMBER")
    strMSQNumber = xmlGetAttributeText(xmlCriticalData, "MORTGAGESUBQUOTENUMBER")
        
    'Calculate and save a new quote
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlRequest = vXMLRequest.cloneNode(False)
    xmlDoc.appendChild xmlRequest
    Set xmlMC = xmlDoc.createElement("MORTGAGECOSTS")
    xmlRequest.appendChild xmlMC
    
    Set xmlApp = xmlGetMandatoryNode(vXMLRequest, "APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONNUMBER")
    strAppFactFindNumber = "1"
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplicationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAppFactFindNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONDATE")
    xmlElement.Text = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONDATE")
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("CONTEXT")
    xmlElement.Text = "COST MODELLING"
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("QUOTATIONNUMBER")
    xmlElement.Text = strQuotationNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElement.Text = strMSQNumber
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("LIFESUBQUOTENUMBER")
    xmlElement.Text = 0
    xmlMC.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("CALCULATEFEESONLY")
    xmlElement.Text = "1"
    xmlMC.appendChild xmlElement
        
    If m_objContext Is Nothing Then
        Set objAQBO = New omAQ.ApplicationQuoteBO
    Else
        Set objAQBO = m_objContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    End If

    Set xmlResponse = objAQBO.CalculateMortgageCosts(xmlDoc.documentElement)
    errCheckXMLResponseNode xmlResponse, , True
    
    Set xmlOneOffCosts = xmlGetMandatoryNode(xmlResponse, "ONEOFFCOSTLIST")
    Set xmlMortgageLender = xmlGetMandatoryNode(xmlResponse, "MORTGAGELENDER")
    
    Set CalculateNonDisposableKFIFees = FormatCalculatedKFIFees(xmlOneOffCosts, xmlMortgageLender)
    
CalculateNonDisposableKFIFeesVbErr:

    Set xmlDoc = Nothing

    Set xmlApp = Nothing
    Set xmlElement = Nothing
    Set xmlMC = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set objAQBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Function

Private Function FormatCalculatedKFIFees(ByVal xmlCostList As IXMLDOMNode, ByVal xmlMortgageLender As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo FormatCalculatedKFIFeesVbErr

    Const cstrFunctionName As String = "FormatCalculatedKFIFees"
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeTemp As IXMLDOMNode
    Dim xmlOneOffCost As IXMLDOMNode
    Dim xmlOneOffCostList As IXMLDOMNodeList
    
    Dim strIdentifier As String
    Dim strNodeName As String
    Dim blnCanAddToLoan As Boolean
    
    Set xmlNodeTemp = GetDisposableKFIOneOffCosts(xmlCostList.xml)
    Set xmlOneOffCostList = xmlNodeTemp.selectNodes("ONEOFFCOST")
    
    For Each xmlOneOffCost In xmlOneOffCostList
        
        strIdentifier = xmlGetAttributeText(xmlOneOffCost, "IDENTIFIER")
        blnCanAddToLoan = False
        
        Select Case strIdentifier
            Case "ARR"
                strNodeName = "ALLOWARRGEMTFEEADDED"
            Case "POR"
                strNodeName = "ALLOWPORTINGFEEADDED"
            Case "VAL", "TPV"
                strNodeName = "ALLOWVALNFEEADDED"
            Case "REI"
                strNodeName = "ALLOWREINSPTFEEADDED"
            Case "ADM"
                strNodeName = "ALLOWADMINFEEADDED"
            Case "SEA"
                strNodeName = "ALLOWSEALINGFEEADDED"
            Case "DEE"
                strNodeName = "ALLOWDEEDSRELFEEADD"
            Case "TTF", "SRF"
                strNodeName = "ALLOWTTFEEADDED"
            Case "LEG"
                strNodeName = "ALLOWLEGALFEEADDED"
            Case "OTH"
                strNodeName = "ALLOWOTHERFEEADDED"
            Case "MIG"
                strNodeName = "ALLOWMIGFEEADDED"
            Case "REV"
                strNodeName = "ALLOWREVALUATIONFEEADDED"
            Case "PSF"
                strNodeName = "ALLOWPRODUCTSWITCHFEEADDED"
            Case "AB"
                strNodeName = "ALLOWFURTHERLENDINGFEEADDED"
            Case "CLI"
                strNodeName = "ALLOWCREDITLIMITINCFEEADDED"
            Case Else
                strNodeName = ""
        End Select
        
        If Len(strNodeName) > 0 Then
            blnCanAddToLoan = xmlGetNodeAsBoolean(xmlMortgageLender, strNodeName)
        End If
        
        If blnCanAddToLoan Then
            xmlSetAttributeValue xmlOneOffCost, "CANBEADDEDTOLOAN", "1"
        Else
            xmlSetAttributeValue xmlOneOffCost, "CANBEADDEDTOLOAN", "0"
        End If
        
        ' PSC 05/03/2007 EP2_1347 - Start
        If Len(xmlGetAttributeText(xmlOneOffCost, "REFUNDAMOUNT")) > 0 Then
            xmlSetAttributeValue xmlOneOffCost, "CANBEREFUNDED", "1"
        Else
            xmlSetAttributeValue xmlOneOffCost, "CANBEREFUNDED", "0"
        End If
        ' PSC 05/03/2007 EP2_1347 - End
    Next
    
    Set FormatCalculatedKFIFees = xmlNodeTemp
    
FormatCalculatedKFIFeesVbErr:

    Set xmlNode = Nothing
    Set xmlNodeTemp = Nothing
    Set xmlOneOffCost = Nothing
    Set xmlOneOffCostList = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    

End Function
' PSC 07/11/2006 EP2_41 - End

'EP2_1994 Borrowed main logic from omAQ.CostModelAssistBO.GetIntroducerProcFeeDetails
Private Function GetDisposableKFIProcFees(ByVal introducerApplicationList As IXMLDOMNodeList, _
                                            ByVal xmlCrudRequest As IXMLDOMNode) As IXMLDOMNode
On Error GoTo GetDisposableKFIProcFeesErr
    
    Const cstrFunctionName As String = "ICostModelAssistBO_GetDisposableKFIProcFees"
    
    Dim objCRUD As omCRUDBO
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCrudResponse As FreeThreadedDOMDocument40
    
    Dim xmlResponse As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlAssociation As IXMLDOMNode
    Dim xmlPackager As IXMLDOMNode
    Dim xmlNetwork As IXMLDOMNode
    Dim xmlAppointedRep As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlOperation As IXMLDOMNode
    Dim xmlGlobalParameter As IXMLDOMNode
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strResponse As String
    Dim strSubmissionRoute As String
    Dim strLoanAndLTVIndicator As String
    Dim strOnlineIndicator As String
    Dim strAssociationFeeAmount As String
    Dim strAssociationFeeRate As String
    Dim strSubmissionRouteId As String
    Dim strPackagerAssociationId As String
    Dim strMortgageClubID As String
    Dim strPackagerID As String
    Dim strPrincipalNetworkFirmId As String
    Dim strARFirmID As String
    Dim strLandGFSARef As String
    Dim strMortgageClubIdentifier As String
    Dim strPrincipalNetworkFSARef As String
    Dim strReturnNullStr As String
    Dim strBasicProcFeeAdj As String

    Dim strRecipientId As String
    Dim strRecipientType As String
    Dim principalFirmId1 As String
    Dim principalFirmId2 As String
    Dim ClubNetworkAssocId  As String
    Dim arFirmId  As String
    Dim IntroducerId As String
    Dim tempPrincipalFirmId As String
    Dim tempClubNetworkAssocId  As String
    Dim tempARFirmId  As String
    Dim tempIntroducerId As String
    Dim introducerApplication As IXMLDOMNode
    Dim xmlApplicationIntroducer As IXMLDOMNode

    'Will be one or two ApplicationIntroducer nodes
    For Each introducerApplication In introducerApplicationList
        tempPrincipalFirmId = xmlGetAttributeText(introducerApplication, "PRINCIPALFIRMID")
        tempClubNetworkAssocId = xmlGetAttributeText(introducerApplication, "CLUBNETWORKASSOCID")
        tempARFirmId = xmlGetAttributeText(introducerApplication, "ARFIRMID")
        tempIntroducerId = xmlGetAttributeText(introducerApplication, "INTRODUCERID")
        If Len(principalFirmId1) > 0 Then
            principalFirmId2 = tempPrincipalFirmId
        Else
            principalFirmId1 = tempPrincipalFirmId
        End If
        If Len(tempClubNetworkAssocId) > 0 Then
            ClubNetworkAssocId = tempClubNetworkAssocId
        End If
        If Len(tempARFirmId) > 0 Then
            arFirmId = tempARFirmId
        End If
        If Len(tempIntroducerId) > 0 Then
            IntroducerId = tempIntroducerId
        End If
        
    Next introducerApplication
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlSetAttributeValue xmlResponse, "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objCRUD = New omCRUDBO
    Else
        Set objCRUD = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCRUDBO")
    End If
    
    Set xmlOperation = xmlOut.createElement("OPERATION")
    xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
    xmlSetAttributeValue xmlOperation, "ENTITY_REF", "GLOBALPARAMETER"
    xmlCrudRequest.appendChild xmlOperation

    Set xmlRoot = xmlOut.createElement("GLOBALPARAMETER")
    xmlSetAttributeValue xmlRoot, "NAME", "FSARefNoLandG"   'EP2_704 GHun
    xmlOperation.appendChild xmlRoot
    
    Set xmlOperation = xmlOut.createElement("OPERATION")
    xmlSetAttributeValue xmlOperation, "CRUD_OP", "READ"
    xmlSetAttributeValue xmlOperation, "SCHEMA_NAME", "omAQ"
    xmlSetAttributeValue xmlOperation, "ENTITY_REF", "DISPOSEKFIINTRODUCERPROCFEEDETAILS"
    xmlCrudRequest.appendChild xmlOperation

    Set xmlRoot = xmlOut.createElement("DISPOSEKFIINTRODUCERPROCFEEDETAILS")
    xmlSetAttributeValue xmlRoot, "PRINCIPALFIRMID1", principalFirmId1
    xmlSetAttributeValue xmlRoot, "PRINCIPALFIRMID2", principalFirmId2
    xmlSetAttributeValue xmlRoot, "CLUBNETWORKASSOCID", ClubNetworkAssocId
    xmlSetAttributeValue xmlRoot, "ARFIRMID", arFirmId
    xmlSetAttributeValue xmlRoot, "INTRODUCERID", IntroducerId
    xmlOperation.appendChild xmlRoot
    
    strResponse = objCRUD.omRequest(xmlCrudRequest.xml)
    Set xmlCrudResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlCrudResponse.documentElement, xmlResponse, True
    
    Set xmlGlobalParameter = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER")
    strLandGFSARef = xmlGetAttributeText(xmlGlobalParameter, "STRING")

    Set xmlAssociation = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/MORTGAGECLUBNETWORKASSOCIATION")
    ' EP2_893 - Move this line to here so we can check if is Nothing BEFORE processing.
    Set xmlNetwork = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/NETWORK")
    ' EP2_893 - Move this line to here so we can check if is Nothing BEFORE processing.
    Set xmlPackager = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/PACKAGER")

    Set xmlApplicationIntroducer = xmlCrudResponse.selectSingleNode("RESPONSE/OPERATION[@ENTITY_REF='DISPOSEKFIINTRODUCERPROCFEEDETAILS']/RESPONSE")
     
    ' EP2_893 - If no values present
    If (xmlAssociation Is Nothing) And (xmlNetwork Is Nothing) And (xmlPackager Is Nothing) Then
        'Set all values to null and finish
        ' EP2_893 - Set strReturnNullStr
        strReturnNullStr = ""
        
        Set xmlRoot = xmlOut.createElement("INTRODUCERPROCFEEDETAILS")
        xmlResponse.appendChild xmlRoot
        
        Set xmlTempNode = xmlOut.createElement("SUBMISSIONROUTE")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("LOANANDLTVINDICATOR")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ONLINEINDICATOR")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ASSOCIATIONFEEAMOUNT")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ASSOCIATIONFEERATE")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("SUBMISSIONROUTEID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PACKAGERASSOCIATIONID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("MORTGAGECLUBID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PACKAGERID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PRINCIPALFIRMNETWORKID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ARFIRMID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        ' PSC 05/02/2007 EP2_1219 - Start
        Set xmlTempNode = xmlOut.createElement("BASICPROCFEEADJ")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode

        Set xmlTempNode = xmlOut.createElement("RECIPIENTID")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("RECIPIENTYPE")
        xmlTempNode.Text = strReturnNullStr
        xmlRoot.appendChild xmlTempNode
        ' PSC 05/02/2007 EP2_1219 - End
        
        xmlResponse.appendChild xmlRoot
        
        Set GetDisposableKFIProcFees = xmlResponse
        
    Else  ' EP2_893 - Else process as before
        
        If Not xmlAssociation Is Nothing Then
            If xmlGetAttributeAsBoolean(xmlAssociation, "PACKAGERINDICATOR") Then
                strPackagerAssociationId = xmlGetMandatoryAttributeText(xmlAssociation, "CLUBNETWORKASSOCIATIONID")
                
                ' PSC 13/02/2007 EP2_1288 - Start
                If xmlPackager Is Nothing Then
                    errThrowError cstrFunctionName, oeRecordNotFound, "A packager cannot be found for this application"
                End If
                
                strPackagerID = xmlGetMandatoryAttributeText(xmlPackager, "PACKAGERID")
                ' PSC 13/02/2007 EP2_1288 - End
            Else
                strMortgageClubID = xmlGetMandatoryAttributeText(xmlAssociation, "CLUBNETWORKASSOCIATIONID")
                strMortgageClubIdentifier = xmlGetAttributeText(xmlAssociation, "IDENTIFIER", "-1")
            End If
            
            strAssociationFeeAmount = xmlGetAttributeText(xmlAssociation, "ASSOCIATIONFEEAMOUNT")
            strAssociationFeeRate = xmlGetAttributeText(xmlAssociation, "ASSOCIATIONFEERATE")
            ' PSC 29/01/2007 EP2_1034
            strBasicProcFeeAdj = xmlGetAttributeText(xmlAssociation, "AGREEDPROCFEERATE")
        Else
            ' EP2_893 - Move this line to new position above to test if all values Nothing BEFORE processing.
            'Set xmlPackager = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/PACKAGER")
            
            If Not xmlPackager Is Nothing Then
                strPackagerID = xmlGetMandatoryAttributeText(xmlPackager, "PACKAGERID")
                ' PSC 29/01/2007 EP2_1034
                strBasicProcFeeAdj = xmlGetAttributeText(xmlPackager, "AGREEDPROCFEERATE")
            End If
        End If
        
        ' EP2_893 - Move this line to new position above to test if both values Nothing BEFORE processing.
        'Set xmlNetwork = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/NETWORK")
        
        If Not xmlNetwork Is Nothing Then
            strPrincipalNetworkFirmId = xmlGetMandatoryAttributeText(xmlNetwork, "PRINCIPALFIRMNETWORKID")
            strPrincipalNetworkFSARef = xmlGetAttributeText(xmlNetwork, "FSAREF")
            
            ' PSC 29/01/2007 EP2_1034 - Start
            If Len(strBasicProcFeeAdj) = 0 Then
                strBasicProcFeeAdj = xmlGetAttributeText(xmlNetwork, "AGREEDPROCFEERATE")
            End If
            ' PSC 29/01/2007 EP2_1034 - End
            
        End If
            
        Set xmlAppointedRep = xmlGetNode(xmlCrudResponse, "RESPONSE/OPERATION/RESPONSE/ARFIRM")
        
        If Not xmlAppointedRep Is Nothing Then
            strARFirmID = xmlGetMandatoryAttributeText(xmlAppointedRep, "ARFIRMID")
        End If
        
        ' Online Indicator
        If Not xmlPackager Is Nothing Then
            strOnlineIndicator = xmlGetAttributeText(xmlPackager, "ONLINEINDICATOR")
        End If
        
        If strOnlineIndicator <> "1" Then
            If Not xmlNetwork Is Nothing Then
                strOnlineIndicator = xmlGetAttributeText(xmlNetwork, "ONLINEINDICATOR")
            End If
        End If
        
        If strOnlineIndicator <> "1" Then
            If Not xmlAppointedRep Is Nothing Then
                strOnlineIndicator = xmlGetAttributeText(xmlAppointedRep, "ONLINEINDICATOR")
            End If
        End If
        
        ' Loan and LTV Indicator
        If Not xmlAssociation Is Nothing Then
            strLoanAndLTVIndicator = xmlGetAttributeText(xmlAssociation, "LOANANDLTVINDICATOR")
        ElseIf Not xmlPackager Is Nothing Then
            strLoanAndLTVIndicator = xmlGetAttributeText(xmlPackager, "LOANANDLTVINDICATOR")
        ElseIf Not xmlNetwork Is Nothing Then
            strLoanAndLTVIndicator = xmlGetAttributeText(xmlNetwork, "LOANANDLTVINDICATOR")
        End If
        
        ' Submission route and Submission Id
        ' PSC 05/02/2007 EP2_1219 - Start
        If Len(strPackagerAssociationId) > 0 Then
            
            strSubmissionRoute = GetFirstComboValueId("ProcFeeSubmissionRoute", "PA")
            strSubmissionRouteId = strPackagerAssociationId
    
            If strPrincipalNetworkFSARef = strLandGFSARef Then
                strRecipientId = strPrincipalNetworkFirmId
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "PF")
            Else
                strRecipientId = strPackagerAssociationId
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "PA")
            End If
            
        ElseIf Len(strMortgageClubID) > 0 Then
            
            strSubmissionRoute = GetFirstComboValueId("ProcFeeSubmissionRoute", "MC")
            strSubmissionRouteId = strMortgageClubID
            
            If IsValidationType("MortgageClubIdType", CSafeInt(strMortgageClubIdentifier), "PMS") Then
                strRecipientId = strPrincipalNetworkFirmId
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "PF")
            Else
                strRecipientId = strMortgageClubID
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "MC")
            End If
        
        ElseIf Len(strPackagerID) > 0 Then
        
            strSubmissionRoute = GetFirstComboValueId("ProcFeeSubmissionRoute", "P")
            strSubmissionRouteId = strPackagerID
            
            If strPrincipalNetworkFSARef = strLandGFSARef Then
                strRecipientId = strPrincipalNetworkFirmId
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "PF")
            Else
                strRecipientId = strPackagerID
                strRecipientType = GetFirstComboValueId("ProcFeeSubmissionRoute", "P")
            End If
            
        Else
            strSubmissionRoute = GetFirstComboValueId("ProcFeeSubmissionRoute", "PF")
            strSubmissionRouteId = strPrincipalNetworkFirmId
            strRecipientId = strSubmissionRouteId
            strRecipientType = strSubmissionRoute
        End If
        ' PSC 05/02/2007 EP2_1219 - End
        
        Set xmlRoot = xmlOut.createElement("INTRODUCERPROCFEEDETAILS")
        xmlResponse.appendChild xmlRoot
        
        Set xmlTempNode = xmlOut.createElement("SUBMISSIONROUTE")
        xmlTempNode.Text = strSubmissionRoute
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("LOANANDLTVINDICATOR")
        xmlTempNode.Text = strLoanAndLTVIndicator
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ONLINEINDICATOR")
        xmlTempNode.Text = strOnlineIndicator
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ASSOCIATIONFEEAMOUNT")
        xmlTempNode.Text = strAssociationFeeAmount
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ASSOCIATIONFEERATE")
        xmlTempNode.Text = strAssociationFeeRate
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("SUBMISSIONROUTEID")
        xmlTempNode.Text = strSubmissionRouteId
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PACKAGERASSOCIATIONID")
        xmlTempNode.Text = strPackagerAssociationId
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("MORTGAGECLUBID")
        xmlTempNode.Text = strMortgageClubID
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PACKAGERID")
        xmlTempNode.Text = strPackagerID
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("PRINCIPALFIRMNETWORKID")
        xmlTempNode.Text = strPrincipalNetworkFirmId
        xmlRoot.appendChild xmlTempNode
        
        Set xmlTempNode = xmlOut.createElement("ARFIRMID")
        xmlTempNode.Text = strARFirmID
        xmlRoot.appendChild xmlTempNode
        
        ' PSC 29/01/2007 EP2_1034 - Start
        Set xmlTempNode = xmlOut.createElement("BASICPROCFEEADJ")
        xmlTempNode.Text = strBasicProcFeeAdj
        xmlRoot.appendChild xmlTempNode
        ' PSC 29/01/2007 EP2_1034 - End
        
        ' PSC 05/02/2007 EP2_1219 - Start
        Set xmlTempNode = xmlOut.createElement("RECIPIENTID")
        xmlTempNode.Text = strRecipientId
        xmlRoot.appendChild xmlTempNode
        ' PSC 05/02/2007 EP2_1219 - End

        Set xmlTempNode = xmlOut.createElement("RECIPIENTTYPE")
        xmlTempNode.Text = strRecipientType
        xmlRoot.appendChild xmlTempNode
        
        xmlResponse.appendChild xmlRoot
        If Not xmlApplicationIntroducer Is Nothing Then
            xmlResponse.appendChild xmlApplicationIntroducer
        End If
           
        Set GetDisposableKFIProcFees = xmlResponse
        
    End If ' EP2_893 - End IF

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetDisposableKFIProcFeesErr:

    'Clean up section
    Set objCRUD = Nothing
    Set xmlOut = Nothing
    Set xmlCrudResponse = Nothing
    Set xmlResponse = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlAssociation = Nothing
    Set xmlPackager = Nothing
    Set xmlNetwork = Nothing
    Set xmlAppointedRep = Nothing
    Set xmlTempNode = Nothing
    Set xmlOperation = Nothing
    Set xmlGlobalParameter = Nothing
    Set introducerApplication = Nothing
    Set xmlApplicationIntroducer = Nothing
 
    errCheckError cstrFunctionName, TypeName(Me)

End Function

'EP2_1994
Private Function CalcProcFees(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
On Error GoTo CalcProcFeesVbErr

    Const cstrFunctionName As String = "CalcProcFees"
        
    Dim xmlIntroducerProcFees As IXMLDOMNode
    Dim xmlProcFees As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlNatureOfLoan As IXMLDOMNode
    Dim xmlComponentList As IXMLDOMNode
    Dim xmlComponent As IXMLDOMNode
    Dim xmlProcFee As IXMLDOMNode
    Dim xmlAssociationFee As IXMLDOMNode
    Dim xmlPackagingFee As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlLoanComponent As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode

    Dim xmlLoanComponents As IXMLDOMNodeList
    
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    
    Dim objCalcsBO As CalcsBO
    Dim objIMortgageSubQuote As IMortgageSubQuoteBO
    Dim objCRUD As omCRUDBO
    
    Dim strSubmissionRouteId As String
    Dim strResponse As String
    Dim strExclusivityLoading As String
    Dim colValueIDs As Collection
    Dim valueID As String
    Dim feeValue As String
    Set colValueIDs = New Collection
    Dim i As Integer
  
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    
    Set xmlNatureOfLoan = xmlGetMandatoryNode(vXMLRequest, "//CALCS/APPLICATIONFACTFIND/NATUREOFLOAN").cloneNode(True)
    Set xmlIntroducerProcFees = xmlGetNode(vXMLRequest, "//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/INTRODUCERPROCFEEDETAILS")
    Set xmlProcFees = xmlIntroducerProcFees.cloneNode(True)
    xmlChangeNodeName xmlProcFees, "INTRODUCERPROCFEEDETAILS", "PROCFEES"
    
    strSubmissionRouteId = xmlGetMandatoryNodeText(xmlProcFees, "SUBMISSIONROUTEID")
    
    xmlChangeNodeName xmlNatureOfLoan, "NATUREOFLOAN", "PRODUCTCATEGORY"
   
    xmlProcFees.appendChild xmlNatureOfLoan
    
    Set xmlComponentList = xmlResponseDoc.createElement("LOANCOMPONENTLIST")
    xmlProcFees.appendChild xmlComponentList
    
    Set xmlLoanComponents = vXMLRequest.selectNodes("//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT")
    
    ' Setup crud request to get loadings
    Set objCRUD = m_objContext.CreateInstance(gstrCRUD_COMPONENT & ".omCrudBO")

    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    xmlSetAttributeValue xmlTempRequest, "CRUD_OP", "READ"
    xmlSetAttributeValue xmlTempRequest, "SCHEMA_NAME", "omCM"
    xmlSetAttributeValue xmlTempRequest, "ENTITY_REF", "MORTGAGEPRODUCTEXCLUSIVITY"
    Set xmlRoot = xmlResponseDoc.createElement("MORTGAGEPRODUCTEXCLUSIVITY")
    xmlSetAttributeValue xmlRoot, "FIRMID", strSubmissionRouteId
    xmlTempRequest.appendChild xmlRoot

    For Each xmlLoanComponent In xmlLoanComponents
    
        strExclusivityLoading = "0"
        
        Set xmlComponent = xmlResponseDoc.createElement("LOANCOMPONENT")
        Set xmlNode = xmlResponseDoc.createElement("LOANAMOUNT")
        'EP2_2448
        xmlNode.Text = xmlGetNodeText(xmlLoanComponent, "TOTALLOANCOMPONENTAMOUNT")
        xmlComponent.appendChild xmlNode
        
        xmlSetAttributeValue xmlRoot, "MORTGAGEPRODUCTCODE", xmlGetMandatoryNodeText(xmlLoanComponent, "MORTGAGEPRODUCTCODE")
        strResponse = objCRUD.omRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True

        Set xmlNode = xmlGetNode(xmlResponseDoc, "RESPONSE/MORTGAGEPRODUCTEXCLUSIVITY")
                
        If Not xmlNode Is Nothing Then
            strExclusivityLoading = xmlGetAttributeText(xmlNode, "PROCFEELOADING", "0")
        End If

        Set xmlNode = xmlResponseDoc.createElement("EXCLUSIVELOADING")
        xmlNode.Text = strExclusivityLoading
        xmlComponent.appendChild xmlNode
        
        xmlComponentList.appendChild xmlComponent
    Next
    'EP2_2448
    Set xmlNode = xmlResponseDoc.createElement("PRODUCTSCHEME")
    xmlNode.Text = xmlGetNodeText(xmlLoanComponents(0), "//CALCS/APPLICATIONFACTFIND/PRODUCTSCHEME")
    xmlProcFees.appendChild xmlNode
    'EP2_2448
    Set xmlNode = xmlResponseDoc.createElement("LOANAMOUNT")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, "//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/AMOUNTREQUESTED")
    xmlProcFees.appendChild xmlNode
    'EP2_2448
    Set xmlNode = xmlResponseDoc.createElement("LTV")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, "//CALCS/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LTV")
    xmlProcFees.appendChild xmlNode
    'EP2_2448
    Set xmlNode = xmlResponseDoc.createElement("TYPEOFAPPLICATION")
    xmlNode.Text = xmlGetNodeText(vXMLRequest, "//CALCS/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
    xmlProcFees.appendChild xmlNode
   
    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    xmlTempRequest.appendChild xmlProcFees
    
    Set objCalcsBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".CalcsBO")
    
    strResponse = objCalcsBO.CalculateProcFees(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    
    Set xmlProcFee = xmlGetNode(xmlResponseDoc, "RESPONSE/PROCFEE")
    Set xmlAssociationFee = xmlGetNode(xmlResponseDoc, "RESPONSE/ASSOCIATIONFEE")
    Set xmlPackagingFee = xmlGetNode(xmlResponseDoc, "RESPONSE/PACKAGINGFEE")
    
    Set xmlTempRequest = xmlGetRequestNode(vXMLRequest)
    Set xmlRoot = xmlResponseDoc.createElement("MORTGAGEINTRODUCERFEELIST")
    xmlTempRequest.appendChild xmlRoot

    If Not xmlProcFee Is Nothing Then
        GetValueIdsForValidationType "IntroducerFeeType", "PR", colValueIDs
        If colValueIDs.Count <> 0 Then
            valueID = colValueIDs.Item(1)
            For i = 1 To colValueIDs.Count
                colValueIDs.Remove (i)
            Next i
        End If
        feeValue = xmlProcFee.Text
        Set xmlTemp = xmlResponseDoc.createElement("MORTGAGEINTRODUCERFEE")
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE", valueID
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_TEXT", "Procuration Fee"
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_VALIDID", "PR"
        xmlSetAttributeValue xmlTemp, "FEEAMOUNT", feeValue
        xmlRoot.appendChild xmlTemp

    End If

    If Not xmlProcFee Is Nothing Then
        GetValueIdsForValidationType "IntroducerFeeType", "A", colValueIDs
        If colValueIDs.Count <> 0 Then
            valueID = colValueIDs.Item(1)
            For i = 1 To colValueIDs.Count
                colValueIDs.Remove (i)
            Next i

        End If
        feeValue = xmlAssociationFee.Text
        Set xmlTemp = xmlResponseDoc.createElement("MORTGAGEINTRODUCERFEE")
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE", valueID
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_TEXT", "Association Fee"
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_VALIDID", "A"
        xmlSetAttributeValue xmlTemp, "FEEAMOUNT", feeValue
        xmlRoot.appendChild xmlTemp

    End If

    If Not xmlProcFee Is Nothing Then
        GetValueIdsForValidationType "IntroducerFeeType", "PA", colValueIDs
        If colValueIDs.Count <> 0 Then
            valueID = colValueIDs.Item(1)
        End If
        feeValue = xmlPackagingFee.Text
        Set xmlTemp = xmlResponseDoc.createElement("MORTGAGEINTRODUCERFEE")
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE", valueID
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_TEXT", "Packaging Fee"
        xmlSetAttributeValue xmlTemp, "INTRODUCERFEETYPE_VALIDID", "PA"
        xmlSetAttributeValue xmlTemp, "FEEAMOUNT", feeValue
        xmlRoot.appendChild xmlTemp

    End If
    
    Set CalcProcFees = xmlRoot
     

CalcProcFeesVbErr:
    
    Set xmlIntroducerProcFees = Nothing
    Set xmlProcFees = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNatureOfLoan = Nothing
    Set xmlComponentList = Nothing
    Set xmlComponent = Nothing
    Set xmlProcFee = Nothing
    Set xmlAssociationFee = Nothing
    Set xmlPackagingFee = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlLoanComponents = Nothing
    Set xmlResponseDoc = Nothing
    Set objCalcsBO = Nothing
    Set objIMortgageSubQuote = Nothing
    Set objCRUD = Nothing
    
    '   re-raise error to calling function
    errCheckError cstrFunctionName, TypeName(Me)

End Function

