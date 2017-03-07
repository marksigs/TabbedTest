VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omHIBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'History:
'
' Prog   Date     Description
'JR      28/08/01  Omiplus24 Modified BuildApplicant, Guarantor/Solicitor/Solicitor/
'                  Employer/Accountant/Valuer/Intemediary/ValuationArrangement Data Blocks
'-------------------------------------------------------------------------------------------------
'History:
'
' Prog   Date           Description
' MDC    23/05/2002     BMIDS00025 IWP1 BM059
' MDC    11/06/2002     BMIDS00142 - BM059 code review change
' MDC    29/07/2002     BMIDS00260 - Code updates related to spec changes and AQR
' MDC    01/08/2002     BMIDS00279 - Fix Security Address and Employment
' MDC    27/09/2002     BMIDS00285 - Self Employed Income
' MDC    30/09/2002     BMIDS00286 - Missing delimiter in Employment blocks
' MO     28/10/2002     BMIDS00752 - Made changes to the Introducer block after the completion of
'                                       the INWP1 work package
' MO     08/11/2002     BMIDS00752 - Hunter fails when there is no quotation
' MDC    03/12/2002     BM0104/BM0127
' GHun   10/01/2003     BM0214     - Guarantor details need to be placed in fields the same way as for applicants
' GHun   20/01/2003     BM0273     - Remove extra CR from last record written to file
' GHun   23/01/2003     BM0283     - Only set introducer address for indirect business
' GD     12/02/2003     BM0346 - Added optional parameter to BuildAddress
' BS    16/04/2003      BM0526  Pick up Security Address for Buy-To-Let
'-------------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date     Description
'MV     08/06/2004  BBG48 - Amended WriteToTextFile() Ref: BMIDS646 Call Compensating Resource Manager to write to text files within a
'                   transaction, allowing roll back if the transaction fails after the text write.
'MV     09/06/2004  BBG48 - Amended WriteToTextFile()
'MV     05/07/2004  BBG05 and BBG454 - Addition of 4 new datablocks and handling of ltd companies
'                    Removal of bank details from header record when not processing A1
'MV     15/07/2004  BBG454 - Amended ApplicantDataBlock() to Inlcude the CustomerRoleType = '1' and '3'
'MV     16/07/2004  BBG1041 - Amended PreviousAddressDataBlock()
'MV     16/07/2004  BBG1045 - Amended PreviousAddressDataBlock
'MV     02/08/2004  BBG1147 - Amended GenerateRecords() to Include IntermediaryData
'MV     12/08/2004  BBG1147 - Amended IntermediaryDataBlock()
'MV     13/08/2004  E2EM00000647 - Amended RebuildHeader()
'MV     22/08/2004  BBG1147 - Amended IntermediaryDataBlock()
'MV     02/09/2004  BBG1147 - Amended IntermediaryDataBlock()
'TK     03/09/2004  BBG1147 - Amended ValuerDataBlock()
'MV     08/09/2004  E2EM00001422 - Amended quite a few methods inline with the spec
'MV     08/09/2004  E2EM00001422 - Amended GuarantorDataBlock()
'MV     11/09/2004  E2EM00001422 - Amended inline with spec
'MV     13/09/2004  E2EM00001422 - Amended inline with spec
'MV     13/09/2004  E2EM00001422 - Amended KeyHolderDataBlock() and EmployersDataBlock()
'MV     21/09/2004  E2EM00001422 - Amended inline with spec
'MSla   23/09/2004  E2EM00001422 Additional formating required
'MV     15/10/2004  E2EM00002435/BBG1411 - Amended AccountantDataBlock(); PreviousAddressDataBlock()
'                   ExistingMortgageDataBlock() and commented ValuationDataBlock in GenerateRecords()
'MV     27/10/2004  BBG1686/BBG1653 - Amended ExistingMortgageDataBlock(); FinancialInstitutionDataBlock;
'                   Created CreateFinancialInstitutionDataBlock()
'MV     29/10/2004  BBG1730 - Amended AccountantDataBlock() and IntermediaryDataBlock()
'IK     04/11/2004  BBG1758 - EmployerDataBlock, amend selection criteria for AccountantDataBlock
'                             to include Employed Contract &  Self-Employed Contract
'IK     04/11/2004  BBG1764 - work telephone not mandatory for IntermediaryDataBlock
'MSla   09/11/2004  BBG1774 Limited Company name should be written to company field, not surname field
'IK     11/11/2004  E2EM00002919 -  various fornatting errors
'                   concatenated fields (some) AddressLine1,AddressLine2,TelephoneNumber losing terminator
'IK     18/11/2004  BBG1808, missing terminator for FI, NK phone number
'IK     18/11/2004  E2EM00003049 - Incorrect Mortgage types of ECROS accounts
'IK     25/11/2004  BBG1810 - SolicitorDataBlock, use applicant solicitor
'TK     26/11/2004  E2EM00003062 - Amend BuildApplicant
'IK     03/12/2004  E2EM00002919 - remove duplicate address values in ME block
'-------------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Private gobjContext As ObjectContext
Private strRecordHeader As String
Private strApplicationNumber As String, strApplicationDate As String
Private strProductType As String
Private strPurchasePrice As String, strAllDataBlocks As String
Private m_strFieldSeparator As String
Const cstrRecordSeparator = vbCrLf
Const cstrFiller = " "
Const cstrLeadingCharForNumbers = vbNullChar '"0"
Const cstrTrailingCharForStrings = vbNullChar '" "
Const cstrContactPhoneGrpName = "ContactTelephoneUsage"
Const cstrContactPhoneValType = "W"
Private m_blnFixedLengthRecords As Boolean
Private m_strHeaderSuffix As String
Private m_objFSO As Scripting.FileSystemObject
Private m_blnNewAdvance As Boolean
Private m_blnBuyToLet As Boolean
Private Function GetData(ByVal vstrXMLRequest As String) As IXMLDOMElement
On Error GoTo GetDataExit
    Const cstrFunctionName = "GetData"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim vxmlRequest As IXMLDOMElement, xmlApplication As IXMLDOMElement
    Dim objRBDO As Object
    Dim strAFFNumber As String, strResponse As String
    'Get Application Number from Request
    Set xmlDoc = xmlLoad(vstrXMLRequest, cstrFunctionName)
    Set xmlApplication = xmlGetMandatoryNode(xmlDoc.documentElement, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
        
    'Load Request XML file and assign the Application Number;
    Set xmlDoc = LoadRequest()
    Set vxmlRequest = xmlDoc.documentElement
    Set xmlApplication = Nothing
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequest, ".//APPLICATION")
    xmlSetAttributeValue xmlApplication, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlApplication, "APPLICATIONFACTFINDNUMBER", strAFFNumber
    ' Call Request Broker component omRB
    
    Set objRBDO = gobjContext.CreateInstance(gstrREQUEST_BROKER_COMPONENT & ".OmRequestDO")
    ' Note the Reuqest is run with UserId and and other attrubutes as mentioned RequestXML file
    
    strResponse = objRBDO.OmDataRequest(xmlDoc.xml)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    Set GetData = xmlDoc.documentElement
GetDataExit:
    Set xmlDoc = Nothing
    Set vxmlRequest = Nothing
    Set xmlApplication = Nothing
    Set objRBDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Function
Private Function GetQuotationNode(ByVal vxmlResponse As IXMLDOMNode) As IXMLDOMNode
'------------------------------------------------------------------------------------
'Find the required quotation node, to fetch the values (TotalLoanAMount, TermInYears)
'from, and return it
'------------------------------------------------------------------------------------
    On Error GoTo GetQuotationNodeExit
    
    Const cstrFunctionName As String = "GetQuotationNode"
    
    Dim xmlApplication As IXMLDOMNode
    Dim xmlQuotation As IXMLDOMNode
    
    Dim strQuoteNumber As String
    
    'if an accepted quote exists, return it else the node corresponding to the ActiveQuote number
    Set xmlApplication = vxmlResponse.selectSingleNode(".//APPLICATION")
    strQuoteNumber = xmlGetAttributeText(xmlApplication, "ACCEPTEDQUOTENUMBER")
    If strQuoteNumber = "" Then
        strQuoteNumber = xmlGetAttributeText(xmlApplication, "ACTIVEQUOTENUMBER")
    End If
        
    'If there is no quotation, raise an error
    If strQuoteNumber = "" Then
        errThrowError cstrFunctionName, oeNoAcceptedOrActiveQuotation
    End If
        
    Set xmlQuotation = vxmlResponse.selectSingleNode(".//QUOTATION[@QUOTATIONNUMBER=" & strQuoteNumber & "]")
    Set GetQuotationNode = xmlQuotation
    
GetQuotationNodeExit:
    
    Set xmlApplication = Nothing
    Set xmlQuotation = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function LoadRequest() As FreeThreadedDOMDocument40
    
On Error GoTo LoadRequestExit
    
    Const cstrFunctionName = "LoadRequest"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim hFile As Integer
    Dim strFileName As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileName = App.Path & "\HunterInterfaceRequest.xml"
    strFileName = Replace(strFileName, "DLL", "XML", 1, 1, vbTextCompare)
    hFile = FreeFile
    Open strFileName For Input Access Read As hFile
    Close hFile
    'If anything now breaks it will be XML related.
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    xmlDoc.Load (strFileName)
    Set LoadRequest = xmlDoc
    If xmlDoc.parseError.errorCode <> 0 Then
        errThrowError cstrFunctionName, oeXMLParserError
    End If
LoadRequestExit:
    
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function
Private Function LoadFormatTemplate() As FreeThreadedDOMDocument40
    
On Error GoTo LoadFormatTemplateExit
    
    Const cstrFunctionName = "LoadFormatTemplate"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim hFile As Integer
    Dim strFileName As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileName = App.Path & "\HIFormatTemplate.xml"
    strFileName = Replace(strFileName, "DLL", "XML", 1, 1, vbTextCompare)
    hFile = FreeFile
    Open strFileName For Input Access Read As hFile
    Close hFile
    'If anything now breaks it will be XML related.
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    xmlDoc.Load (strFileName)
    Set LoadFormatTemplate = xmlDoc
    If xmlDoc.parseError.errorCode <> 0 Then
        errThrowError cstrFunctionName, oeXMLParserError
    End If
LoadFormatTemplateExit:
    
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function

'MSla 23/09/2004 E2EM00001422
Private Function GetLatestAddressNode(ByVal vxmlPrevAddressList As IXMLDOMNodeList) As IXMLDOMNode

    On Error GoTo GetLatestAddressNodeExit
                                           
    Const cstrFunctionName As String = "PreviousAddressDataBlock"
    
    
    Dim xmlAddress As IXMLDOMNode
    Dim xmlAddressNew As IXMLDOMNode
    
    
    For Each xmlAddressNew In vxmlPrevAddressList
    
        If Not xmlAddress Is Nothing Then
            If (CSafeDate(xmlGetAttributeText(xmlAddressNew, "DATEMOVEDOUT"))) > _
                (CSafeDate(xmlGetAttributeText(xmlAddress, "DATEMOVEDOUT"))) Then
                
                    Set xmlAddress = xmlAddressNew.cloneNode(True)
            End If
    
        Else
            Set xmlAddress = xmlAddressNew.cloneNode(True)
        End If
    
    Next
    
    
    Set GetLatestAddressNode = xmlAddress.cloneNode(True)
    
GetLatestAddressNodeExit:
                                            
    Set xmlAddress = Nothing
    Set xmlAddressNew = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'MSla 23/09/2004 E2EM00001422 - End

Private Function RecordHeader(ByVal vxmlResponse As IXMLDOMElement, _
                              ByVal vxmlFormatTemplate As IXMLDOMElement) As Boolean
On Error GoTo RecordHeaderExit
    
    Const cstrFunctionName = "RecordHeader"
        
    Dim xmlApplication As IXMLDOMElement, xmlApplicationFormat As IXMLDOMElement
    Dim xmlTable As IXMLDOMElement, xmlFormat As IXMLDOMElement
    Dim xmlQuotation As IXMLDOMNode
    'BMIDS00025 MDC 05/06/2002
    Dim lngAppType As Long
    Dim strApplicationType As String
    'BMIDS00025 MDC 05/06/2002 - End
    'BM0104 MDC 03/12/2002 - Buy To Let
    Dim intTemp As Integer
    'BM0104 MDC 03/12/2002 - End
    'BMIDS00260 MDC 29/07/2002
    Dim strScheme As String
    'BMIDS00260 MDC 29/07/2002 - End
    Set xmlApplication = xmlGetMandatoryNode(vxmlResponse, ".//APPLICATION")
    Set xmlApplicationFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICATION")
    strRecordHeader = ""
    strRecordHeader = strRecordHeader & GetFieldValueString(xmlApplication, xmlApplicationFormat, "APPLICATIONNUMBER")
    
'IK 18/11/2004 E2EM00003049 - Incorrect Mortgage types of ECROS accounts
    intTemp = xmlGetAttributeAsInteger(xmlApplication, "SPECIALGROUP", "0")
    If IsValidationType("SpecialGroup", intTemp, "BTL") Then
        m_blnBuyToLet = True
        strApplicationType = "B2L"
    ElseIf IsValidationType("SpecialGroup", intTemp, "L") Then
        ' m_blnBuyToLet used to indicate that NEWPROPERTY to be used for Security Address
        m_blnBuyToLet = True
        strApplicationType = "L"
    End If
'IK 18/11/2004 E2EM00003049 - ends
    
    If Len(Trim$(strApplicationType)) = 0 Then
        'BMIDS00025 MDC 05/06/2002 - Product code
    '    strRecordHeader = strRecordHeader & GetFieldValueString(xmlApplication, xmlApplicationFormat, "PRODUCTCODE")
        lngAppType = xmlGetMandatoryAttributeAsLong(xmlApplication, "TYPEOFAPPLICATION")
        If IsValidationType("TypeOfMortgage", lngAppType, "N") Then
            'New Advance
            strApplicationType = "NA"
            'BMIDS00260 MDC 29/07/2002
            m_blnNewAdvance = True
            'BMIDS00260 MDC 29/07/2002 - End
        ElseIf IsValidationType("TypeOfMortgage", lngAppType, "R") Then
            'Remortgage
            strApplicationType = "RM"
        ElseIf IsValidationType("TypeOfMortgage", lngAppType, "F") Then
            'Further Advance
            strApplicationType = "FA"
        ElseIf IsValidationType("TypeOfMortgage", lngAppType, "T") Then
            'Transfer of Equity
            strApplicationType = "TOFE"
        ElseIf IsValidationType("TypeOfMortgage", lngAppType, "P") Then
            'Purchase of Equity
            strApplicationType = "POFE"
        Else
            strApplicationType = ""
        End If
    End If
    'BM0104 MDC 03/12/2002 - End
    strRecordHeader = strRecordHeader & strApplicationType & m_strFieldSeparator
    'BMIDS00025 MDC 05/06/2002 - End
    strApplicationDate = GetFieldValueString(xmlApplication, xmlApplicationFormat, "APPLICATIONDATE")
    'BMIDS00025 MDC 24/05/2002 - Support variable length records
    strRecordHeader = strRecordHeader & strApplicationDate
    If m_blnFixedLengthRecords Then
        strRecordHeader = strRecordHeader & String(5, cstrFiller)
    End If
    'BMIDS00025 MDC 24/05/2002 - End
    ' Loan Amount
    Set xmlQuotation = GetQuotationNode(vxmlResponse)
    Set xmlTable = xmlGetNode(xmlQuotation, ".//MORTGAGESUBQUOTE")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//MORTGAGESUBQUOTE")
    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "TOTALLOANAMOUNT")
    'Term in Years
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    If Not xmlQuotation Is Nothing Then
        Set xmlTable = xmlGetNode(xmlQuotation, ".//LOANCOMPONENT[@LOANCOMPONENTSEQUENCENUMBER=1]")
    End If
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//LOANCOMPONENT")
    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "TERMINYEARS")
    'BMIDS00260 MDC 29/07/2002
    Set xmlTable = xmlGetNode(vxmlResponse, ".//BANKORBSOC")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICATIONBANKBUILDINGSOC")
    If xmlTable Is Nothing Then
        'Empty fields
        If m_blnFixedLengthRecords Then
            strRecordHeader = strRecordHeader & MakeString(10, cstrTrailingCharForStrings)   'Bank name
            strRecordHeader = strRecordHeader & MakeString(8, cstrTrailingCharForStrings)   'Sort code
            strRecordHeader = strRecordHeader & MakeString(16, cstrTrailingCharForStrings)   'Account number
        Else
            strRecordHeader = strRecordHeader & String(3, m_strFieldSeparator)
        End If
    Else
        'Bank Name
        strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "COMPANYNAME")
        'Sort code - Third Party or Directory?
        If xmlAttributeValueExists(xmlTable, "THIRDPARTYBANKSORTCODE") Then
            strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "THIRDPARTYBANKSORTCODE")
        ElseIf xmlAttributeValueExists(xmlTable, "NAMEANDADDRESSBANKSORTCODE") Then
            strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "NAMEANDADDRESSBANKSORTCODE")
        Else
            strRecordHeader = strRecordHeader & m_strFieldSeparator
        End If
        'Account Number
        strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "ACCOUNTNUMBER")
    End If
'    'BMIDS00025 MDC 24/05/2002
'    'Bank Name
'    Set xmlTable = xmlGetNode(vxmlResponse, ".//PAYEEHISTORY")
'    Set xmlFormat = xmlGetNode(vxmlFormatTemplate, ".//PAYEEHISTORY")
'    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "BANKNAME")
'    'strRecordHeader = strRecordHeader & Space(10) & m_strFieldSeparator
'
'    'Bank Sort Code
''    Set xmlTable = Nothing: Set xmlFormat = Nothing
''    Set xmlTable = xmlGetNode(vxmlResponse, ".//BANKORBSOC")
''    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICATIONBANKBUILDINGSOC")
''    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "THIRDPARTYBANKSORTCODE")
'    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "BANKSORTCODE")
'
'    'Account Number*****
''    strRecordHeader = strRecordHeader & Space(8)
'    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "ACCOUNTNUMBER")
'    'BMIDS00025 MDC 24/05/2002 - End
'
'    'BMIDS00025 MDC 24/05/2002
'    'Filler
'    If m_blnFixedLengthRecords Then
'        strRecordHeader = strRecordHeader & String(4, cstrFiller) & m_strFieldSeparator
'    End If
'    'BMIDS00025 MDC 24/05/2002 - End
    
    'Latest Valuation
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    'BMIDS00025 MDC 05/06/2002
'    Set xmlTable = xmlGetNode(vxmlResponse, ".//NEWPROPERTY")
'    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//NEWPROPERTY")
'    strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "VALUATIONPRICE")
    Set xmlTable = xmlGetNode(vxmlResponse, ".//VALUATIONREPORT")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUATIONREPORT")
' ik_E2EM00002919
    strPurchasePrice = GetFieldValueString(xmlTable, xmlFormat, "PRESENTVALUATION", False)
    If CSafeLng(strPurchasePrice) = 0 Then
        strPurchasePrice = GetFieldValueString(xmlApplication, xmlApplicationFormat, "PURCHASEPRICEORESTIMATEDVALUE", False)
    End If
' ik_E2EM00002919_ends
    If m_blnFixedLengthRecords Then
        strRecordHeader = strRecordHeader & MakeString(8 - Len(strPurchasePrice), cstrLeadingCharForNumbers) & strPurchasePrice
    Else
        strRecordHeader = strRecordHeader & strPurchasePrice & m_strFieldSeparator
    End If
    'strRecordHeader = strRecordHeader & GetFieldValueString(xmlTable, xmlFormat, "PRESENTVALUATION")
    'BMIDS00025 MDC 05/06/2002 - End
    'BMIDS00260 MDC 29/07/2002
'    'BMIDS00025 MDC 24/05/2002
''    'Product Type
''    strRecordHeader = strRecordHeader & Space(8)
'    'Special Scheme
'    If m_blnFixedLengthRecords Then
'        strRecordHeader = strRecordHeader & Space(15)
'    End If
'    strRecordHeader = strRecordHeader & m_strFieldSeparator
    Set xmlTable = xmlGetMandatoryNode(vxmlResponse, "APPLICATION")
    strScheme = xmlGetAttributeText(xmlTable, "SPECIALSCHEME_TEXT")
    If m_blnFixedLengthRecords Then
        strRecordHeader = strRecordHeader & strScheme & MakeString(15 - Len(strScheme), cstrTrailingCharForStrings)
    Else
' ik_E2EM00002919
        strRecordHeader = strRecordHeader & Trim(Left(strScheme, 15)) & m_strFieldSeparator
' ik_E2EM00002919_ends
    End If
    'BMIDS00260 MDC 29/07/2002 - End
    'Years at address
    'Only required for Applicant records
    'Account Spare 3
    If m_blnFixedLengthRecords Then
        m_strHeaderSuffix = Space$(10)
    End If
    m_strHeaderSuffix = m_strHeaderSuffix & m_strFieldSeparator
    'BMIDS00025 MDC 24/05/2002 - End
    'Purchase Price, ProductType - This is required in some data blocks built later
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    Set xmlTable = xmlGetNode(vxmlResponse, ".//REPORTONTITLE")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//REPORTONTITLE")
    strPurchasePrice = GetFieldValueString(xmlTable, xmlFormat, "PURCHASEPRICE")
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    Set xmlTable = xmlGetNode(vxmlResponse, ".//APPLICATION")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICATION")
    strProductType = GetFieldValueString(xmlTable, xmlFormat, "TYPEOFAPPLICATION")
    RecordHeader = True
RecordHeaderExit:
        
    Set xmlApplication = Nothing
    Set xmlApplicationFormat = Nothing
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
End Function
Private Function ApplicantDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode, _
                                    ByVal vxmlApplicantList As IXMLDOMNodeList) As Boolean
                                    
    On Error GoTo BuildApplicantExit
    
    Const cstrFunctionName As String = "ApplicantDataBlock"
    
    Dim strApplicantBlock As String
    
    Dim xmlApplicantList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    
    Dim intApplicants As Integer
    Dim intCurrentApplicant As Integer
    
    Set xmlApplicantList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1'or @CUSTOMERROLETYPE='3']")
    intApplicants = IIf(xmlApplicantList.length > 4, 4, xmlApplicantList.length)
    
    If intApplicants > 0 Then
        
        For intCurrentApplicant = 1 To intApplicants
            Set xmlApplicant = xmlApplicantList.Item(intCurrentApplicant - 1)
            ' Build Applicant Data for this applicant
            strApplicantBlock = "A" + Trim(Str(intCurrentApplicant)) & m_strFieldSeparator
            BuildApplicant xmlApplicant, vxmlFormatTemplate, strApplicantBlock
        Next intCurrentApplicant
        
    Else
        
        errThrowError cstrFunctionName, oeMissingApplicantsForApplication '127     'Missing Applicants For Application
        
    End If
    
    ApplicantDataBlock = True

BuildApplicantExit:
    
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

    
End Function
Private Function BuildApplicant(ByVal vxmlApplicant As IXMLDOMNode, _
                                ByVal vxmlFormatTemplate As IXMLDOMNode, _
                                ByRef strApplicantBlock As String) As String
    
    On Error GoTo BuildApplicantExit
    Const cstrFunctionName As String = "BuildApplicant"
    
    Dim xmlTable As IXMLDOMElement
    Dim xmlFormat As IXMLDOMElement
    Dim xmlIncome As IXMLDOMNode
    Dim xmlIncomeList As IXMLDOMNodeList
    
    Dim strTemp As String
    Dim strIncome As String
    Dim strValueId  As String
    
    Dim dblTotalAnnualIncome As Double
    Dim dblIncome As Double
    
    Dim dtDateMovedIn As Date
    
    Dim intFrequency As Integer
    Dim intNumberOfYears As Integer
    
    Dim strCustomerRoleType As String
    
    ' Customer Title - Get the desc for TITLE, if it is empty use value of 'TITLEOTHER'
    Set xmlTable = xmlGetNode(vxmlApplicant, ".//CUSTOMER")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT")
    
    strCustomerRoleType = xmlGetAttributeText(vxmlApplicant, "CUSTOMERROLETYPE")
    
' ik_E2EM00002919 TITLE will not be empty, will contain terminator
'    strTemp = GetFieldValueString(xmlTable, xmlFormat, "TITLE")
'    If strTemp <> "" Then
'        strApplicantBlock = strApplicantBlock & strTemp
'    Else
'        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "TITLEOTHER")
'    End If
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "TITLE", False)
    If strTemp <> "" Then
        strApplicantBlock = strApplicantBlock & strTemp & m_strFieldSeparator
    Else
        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "TITLEOTHER")
    End If
' ik_E2EM00002919_ends

    'Initials - FirstForeName[1] + SecondForeName[1]
' ik_E2EM00002919
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", False) & _
              GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME")
' ik_E2EM00002919_ends

    strTemp = Trim$(strTemp)
    
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & strTemp & MakeString(3 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strApplicantBlock = strApplicantBlock & strTemp
    End If
    
    'MSla   09/11/2004  BBG1774
    If strCustomerRoleType = "3" Then 'COMPANY
        ' Sur Name
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
        
        'Date Of Birth
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
        
        'Sex
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
        
        'Company Name
        strTemp = GetFieldValueString(xmlTable, xmlFormat, "SURNAME")
        
        If m_blnFixedLengthRecords Then
            strApplicantBlock = strApplicantBlock & strTemp & MakeString(30 - Len(strTemp), cstrTrailingCharForStrings)
        Else
            strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "SURNAME")
        End If
    Else
        ' Sur Name
        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "SURNAME")

        'Date Of Birth
        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "DATEOFBIRTH")

        'Sex
        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "GENDER")

        'Company Name
        If m_blnFixedLengthRecords Then
            strApplicantBlock = strApplicantBlock & Space(30)
        Else
            strApplicantBlock = strApplicantBlock & m_strFieldSeparator
        End If
    End If
    'MSla   09/11/2004  BBG1774 - End
    
    ' Choose Current Address, and Address Format mentioned below Customer node
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
    
    Set xmlTable = vxmlApplicant.selectSingleNode(".//CUSTOMERADDRESS[@ADDRESSTYPE='1']")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT/ADDRESS")
    
        
' ik_E2EM00002919 - use standard address routine
    strApplicantBlock = strApplicantBlock & BuildAddress(xmlTable, xmlFormat, True)
' ik_E2EM00002919_ends

    ' Get DateMovedIn and find years at address
    If Not xmlTable Is Nothing Then
        strTemp = xmlGetAttributeText(xmlTable, "DATEMOVEDIN")
    Else
        strTemp = ""
    End If
    
    If strTemp <> "" Then
        dtDateMovedIn = Format(strTemp, "dd/mm/yyyy")
        intNumberOfYears = DateDiff("YYYY", dtDateMovedIn, Now)
    Else
        intNumberOfYears = 0
    End If
    
    'Telephone
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
    
    'Use Work or mobile if home no not present
    strValueId = GetFirstComboValueId("TelephoneUsage", "H")
    Set xmlTable = vxmlApplicant.selectSingleNode(".//CUSTOMERTELEPHONENUMBER[@USAGE='" & strValueId & "']")
    If xmlTable Is Nothing Then
        strValueId = GetFirstComboValueId("TelephoneUsage", "W")
        Set xmlTable = vxmlApplicant.selectSingleNode(".//CUSTOMERTELEPHONENUMBER[@USAGE='" & strValueId & "']")
    End If
    If xmlTable Is Nothing Then
        strValueId = GetFirstComboValueId("TelephoneUsage", "M")
        Set xmlTable = vxmlApplicant.selectSingleNode(".//CUSTOMERTELEPHONENUMBER[@USAGE='" & strValueId & "']")
    End If
    
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT/CUSTOMERTELEPHONENUMBER")
' ik_E2EM00002919
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "AREACODE", False)
' ik_E2EM00002919_ends
    If Len(Trim$(strTemp)) > 0 Then
        strTemp = strTemp & " "
    End If
    
'   ik_E2EM00002919 - telephone number needs truncating
    strTemp = strTemp & GetFieldValueString(xmlTable, xmlFormat, "TELEPHONENUMBER", False)
    strTemp = Left(Trim(strTemp), 15) & m_strFieldSeparator
'   ik_E2EM00002919_ends
    
    
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & strTemp & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strApplicantBlock = strApplicantBlock & strTemp
    End If
    
    'Valuer Code
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & Space(8)
    Else
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
    End If
    
    'Person Spare 1 - Forename
    Set xmlTable = xmlGetNode(vxmlApplicant, ".//CUSTOMER")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//FIRSTNAME")
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME")
    'Person Spare 2 - Income
    
    'Calculate earned income related to main employment for this applicant
    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]/EMPLOYEDDETAIL/EARNEDINCOME[@EARNEDINCOMEAMOUNT > 0]")
    For Each xmlIncome In xmlIncomeList
        dblIncome = xmlGetAttributeAsDouble(xmlIncome, "EARNEDINCOMEAMOUNT")
        intFrequency = CSafeInt(xmlGetMandatoryAttributeText(xmlIncome, "PAYMENTFREQUENCYTYPE"))
        dblTotalAnnualIncome = dblTotalAnnualIncome + (dblIncome * intFrequency)
    Next
    'TK 26/11/2004 E2EM00003062
    'Include Self Employed Income
'    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]/SELFEMPLOYEDDETAIL[@YEAR1AMOUNT > 0  or  @YEAR2AMOUNT > 0  or  @YEAR2AMOUNT > 0]")
'    For Each xmlIncome In xmlIncomeList
'        dblIncome = CalculateAverageNetProfit(xmlIncome)
'        dblTotalAnnualIncome = dblTotalAnnualIncome + dblIncome
'    Next
    
    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]/SELFEMPLOYEDDETAIL[@ANNUALGROSSINCOME > 0]")
    For Each xmlIncome In xmlIncomeList
        dblIncome = xmlGetAttributeAsDouble(xmlIncome, "ANNUALGROSSINCOME")
        dblTotalAnnualIncome = dblTotalAnnualIncome + dblIncome
    Next
    'TK 26/11/2004 E2EM00003062 End
    strIncome = Format$(dblTotalAnnualIncome, "########0.00")
    
    If m_blnFixedLengthRecords Then

        strApplicantBlock = strApplicantBlock & strIncome & MakeString(15 - Len(strIncome), cstrLeadingCharForNumbers)
    Else
        strApplicantBlock = strApplicantBlock & strIncome & m_strFieldSeparator
    End If

    'Person Spare 3
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & String(10, cstrLeadingCharForNumbers)
    Else
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
    End If
    
    'Add/Update
    strApplicantBlock = strApplicantBlock & "A"

    'Years at Address
    strTemp = Trim(Str(intNumberOfYears))
    
    '----------------------------------------------------------------------------------------------
    ' Modify the record header, append this string to Record header and then to the full data block.
    '-----------------------------------------------------------------------------------------------
    'Set ProductCode to MORTGAGE in RecordHeader
    Dim strNewRecordHeader As String
    If m_blnFixedLengthRecords Then
        strTemp = strTemp & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strTemp = strTemp & m_strFieldSeparator
    End If
    strNewRecordHeader = strRecordHeader & strTemp & m_strHeaderSuffix
    
    '' BBG454 - need to remove bank details if we arent on the A1 record
    If Left(strApplicantBlock, 2) <> "A1" Then
        strNewRecordHeader = RebuildHeader(strNewRecordHeader)
    End If
    strAllDataBlocks = strAllDataBlocks & strNewRecordHeader & strApplicantBlock & cstrRecordSeparator
BuildApplicantExit:
    
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
    Set xmlIncome = Nothing
    Set xmlIncomeList = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Function
'''''''''''''''''''''''
'' BBG454 - rebuild record header to remove bank details
'' if we are not on the A1 record
Private Function RebuildHeader(strHeader As String) As String


On Error GoTo RebuildHeaderExit

    Const cstrFunctionName = "RebuildHeader"
    Dim iPos As Integer
    Dim iStart, iEnd As Integer
    Dim strHeaderStart, strHeaderEnd As String
    Dim strHeaderMiddle As String
    
    iStart = 1
    iEnd = 1
    For iPos = 1 To 9
        If iPos < 6 Then
            iStart = InStr(iStart + 1, strHeader, "|")
        End If
        iEnd = InStr(iEnd + 1, strHeader, "|")
    Next
    strHeaderStart = Mid(strHeader, 1, iStart - 1)
    strHeaderEnd = Mid(strHeader, iEnd + 1)
    
    If m_blnFixedLengthRecords Then
        strHeaderMiddle = strHeaderMiddle & String(10, cstrTrailingCharForStrings) 'Bank Name
        strHeaderMiddle = strHeaderMiddle & String(8, cstrTrailingCharForStrings) 'Sort Code
        strHeaderMiddle = strHeaderMiddle & String(16, cstrTrailingCharForStrings) 'Account Num
        strHeaderMiddle = strHeaderMiddle & String(8, cstrTrailingCharForStrings) 'valuation

    Else
        strHeaderMiddle = MakeString(5, m_strFieldSeparator)
    End If
    
    RebuildHeader = strHeaderStart & strHeaderMiddle & strHeaderEnd
    
    Exit Function
RebuildHeaderExit:
    errCheckError cstrFunctionName, TypeName(Me)
End Function

Private Function CalculateAverageNetProfit(ByVal xmlIncomeNode As IXMLDOMNode) As Double
    
    On Error GoTo CalculateAverageNetProfitExit
    Dim dblAnnualNetProfit As Double
    Dim dblTotalProfit As Double
    Dim intNumberOfYears As Integer
    
    'Year 1
    dblAnnualNetProfit = xmlGetAttributeAsDouble(xmlIncomeNode, "YEAR1AMOUNT")
    If dblAnnualNetProfit > 0 Then
        dblTotalProfit = dblTotalProfit + dblAnnualNetProfit
        intNumberOfYears = intNumberOfYears + 1
    End If
    
    'Year 2
    dblAnnualNetProfit = xmlGetAttributeAsDouble(xmlIncomeNode, "YEAR2AMOUNT")
    If dblAnnualNetProfit > 0 Then
        dblTotalProfit = dblTotalProfit + dblAnnualNetProfit
        intNumberOfYears = intNumberOfYears + 1
    End If
    
    'Year 3
    dblAnnualNetProfit = xmlGetAttributeAsDouble(xmlIncomeNode, "YEAR3AMOUNT")
    If dblAnnualNetProfit > 0 Then
        dblTotalProfit = dblTotalProfit + dblAnnualNetProfit
        intNumberOfYears = intNumberOfYears + 1
    End If
    
    'Average
    If dblTotalProfit > 0 And intNumberOfYears > 0 Then
        CalculateAverageNetProfit = dblTotalProfit / intNumberOfYears
    Else
        CalculateAverageNetProfit = 0
    End If
    
CalculateAverageNetProfitExit:
    
    If Err.Number <> 0 Then
        Err.Clear
        CalculateAverageNetProfit = 0
    End If
    
End Function
'BMIDS00285 MDC 27/09/2002 - End

' ik_E2EM00002919 functions simplified, unused parameters removed etc.
Private Function GetFieldValueString(ByVal vxmlDataElement As IXMLDOMElement, _
                                     ByVal vxmlFormatElement As IXMLDOMElement, _
                                     ByVal strFieldName As String, _
                                     Optional blnAddFieldSeparator As Boolean = True)
'--------------------------------------------------------------------------------------
'Description
'   Format the value based on the info provided in the Formatting Template
'Pass:
'   vxmlDataElement   : Data Node
'   vxmlFormatElement : Node from formatting template
'   strFieldName      : Name of the field in Data Element and Formatting Template
'   blnLastField      : Is this last field in the record ?
'--------------------------------------------------------------------------------------
    
    Dim xmlFieldFormat As IXMLDOMElement
    Dim strFieldValue As String, strIsEmpty As String, strSeparator As String
    Dim intFieldSize As Integer, strFixedValue As String, strDataType As String
    Dim strRet As String    'BMIDS00025 MDC 24/05/2002
    
    'Get the Format from Format Element
    If vxmlFormatElement.nodeName = strFieldName Then
        Set xmlFieldFormat = vxmlFormatElement
    Else
        Set xmlFieldFormat = xmlGetMandatoryNode(vxmlFormatElement, ".//" & strFieldName)
    End If
    strFixedValue = xmlGetAttributeText(xmlFieldFormat, "FIXEDVALUE")
    strIsEmpty = xmlGetAttributeText(xmlFieldFormat, "ISEMPTY")
    intFieldSize = xmlGetMandatoryAttributeText(xmlFieldFormat, "FIELDSIZE")
    strDataType = xmlGetMandatoryAttributeText(xmlFieldFormat, "DATATYPE")
    Set xmlFieldFormat = Nothing
    
    'Get The value from Data Element
    If Not vxmlDataElement Is Nothing Then
        If strDataType <> "COMBO" Then
            strFieldValue = xmlGetAttributeText(vxmlDataElement, strFieldName)
        Else    ' Get the value of the attrib the gives description for the code
            strFieldValue = xmlGetAttributeText(vxmlDataElement, strFieldName & "_TEXT")
            strDataType = "STRING"
        End If
    Else
        strFieldValue = ""
    End If
        
    'If fixed value is mentioned, just return it
    If Len(strFixedValue) > 0 Then
        strFieldValue = strFixedValue
    End If
    
    'BMIDS00025 MDC 24/05/2002
    strRet = Trim(Left(GetFormattedValue(strFieldValue, intFieldSize, strDataType), intFieldSize))
    
    If blnAddFieldSeparator Then
        strRet = strRet & m_strFieldSeparator
    End If
    
    GetFieldValueString = strRet
    'GetFieldValueString = GetFormattedValue(strFieldValue, intFieldSize, strDataType) & strSeparator
    'BMIDS00025 MDC 24/05/2002 - End
End Function
Private Function GetFormattedValue(ByVal strFieldValue As String, _
                                   ByVal intFieldSize As Integer, _
                                   ByVal strDataType As String) As String
'--------------------------------------------------------------------------
'Description : Format the value passed in
'Pass        : strFieldValue - value to be formatted
'              intFieldSize  - size of the formatted value
'              strDataType   - Data Type (String, Number, string)
'Return      : Formatted string
'--------------------------------------------------------------------------
    Dim intValueLength As Integer
    intValueLength = Len(strFieldValue)
    If intValueLength = 0 Then
        'BMIDS00025 MDC 24/05/2002
        'Only pad items if using fixed length format
        If m_blnFixedLengthRecords Then
            Select Case strDataType
                Case "STRING"  'return spaces
                    GetFormattedValue = Space(intFieldSize)
                Case "NUMBER"  'return zeros
                    GetFormattedValue = String(intFieldSize, "0")
                Case "DATE"    'return spaces
                    GetFormattedValue = Space(intFieldSize)
                Case Else
                    errThrowError "omHIBO.GetFormattedValue", oeInValidDataTypeValue
            End Select
        End If
        'BMIDS00025 MDC 24/05/2002 - End
    Else
        Select Case strDataType
            Case "STRING"  'add trailing spaces
                'Only pad field if using fixed length format
                If m_blnFixedLengthRecords Then
                    strFieldValue = strFieldValue & MakeString(intFieldSize - intValueLength, cstrTrailingCharForStrings)
                End If
                GetFormattedValue = strFieldValue
                'BMIDS00025 MDC 23/05/2002 - End
            Case "NUMBER"
                'Only pad field if using fixed length format
                If m_blnFixedLengthRecords Then
                    'Pad with leading zeros
                    'GetFormattedValue = String(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                    strFieldValue = MakeString(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                End If
                GetFormattedValue = strFieldValue
            Case "DATE"
                GetFormattedValue = Left(strFieldValue, 2) & Mid(strFieldValue, 4, 2) & Mid(strFieldValue, 7, 4)
            Case Else
                errThrowError "omHIBO.GetFormattedValue", oeInValidDataTypeValue
        End Select
    End If
End Function
' ik_E2EM00002919_ends

Private Function GuarantorDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo GuarantorDataBlockExit
    Const cstrFunctionName As String = "GuarantorDataBlock"
    
    Dim xmlGuarantorList As IXMLDOMNodeList
    Dim xmlGuarantor As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlTable As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    
    Dim strGuarantorBlock As String
    Dim strCustomerOrder As String
    Dim strTemp As String
    Dim strNewHeader As String
    Dim strCompanyName  As String
      
    Dim intCustomerOrder As Integer
    
    Set xmlGuarantorList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='2']")
    
    ' Find the first guarantor
    intCustomerOrder = 999  'Initialise Customer Order with high value
    For Each xmlNode In xmlGuarantorList
        strCustomerOrder = xmlGetAttributeText(xmlNode, "CUSTOMERORDER")
        If strCustomerOrder <> "" Then
            If CSafeInt(strCustomerOrder) < intCustomerOrder Then
                Set xmlGuarantor = xmlNode
                intCustomerOrder = CSafeInt(strCustomerOrder)
            End If
        End If
    Next xmlNode
    
    'If guarantor record is found, build the data block
    If Not xmlGuarantor Is Nothing Then
        
        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMER")
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//GUARANTOR")
        
        strGuarantorBlock = "GA" & m_strFieldSeparator
        
        'MSla 23/09/2004 E2EM00001422
        
'        If m_blnFixedLengthRecords Then
'            strGuarantorBlock = strGuarantorBlock & MakeString(5, cstrTrailingCharForStrings)   'Title
'            strGuarantorBlock = strGuarantorBlock & MakeString(3, cstrTrailingCharForStrings)   'Initials
'            strGuarantorBlock = strGuarantorBlock & MakeString(20, cstrTrailingCharForStrings)  'Surname
'            strGuarantorBlock = strGuarantorBlock & MakeString(10, cstrTrailingCharForStrings)  'Date of Birth
'            strGuarantorBlock = strGuarantorBlock & MakeString(1, cstrTrailingCharForStrings)   'Gender
'        Else
'            strGuarantorBlock = strGuarantorBlock & MakeString(5, m_strFieldSeparator)
'        End If
            
'        'Company Name
'        strGuarantorBlock = strGuarantorBlock & _
'                Trim$(Left$(GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", False, , False) & " " & _
'                GetFieldValueString(xmlTable, xmlFormat, "SURNAME", , , True), 30))
    
    
        strTemp = GetFieldValueString(xmlTable, xmlFormat, "TITLE")
        If strTemp <> "" Then
            strGuarantorBlock = strGuarantorBlock & strTemp
        Else
            strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "TITLEOTHER")
        End If
        
' ik_E2EM00002919
        'Initials - FirstForeName[1] + SecondForeName[1]
        strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", False) & _
                  GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME") & _
                  GetFieldValueString(xmlTable, xmlFormat, "SURNAME")
' ik_E2EM00002919_ends

        strTemp = Trim$(strTemp)

' ik_E2EM00002919
        'Date Of Birth
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "DATEOFBIRTH")

        'Sex
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "GENDER")
        
'        If m_blnFixedLengthRecords Then
'            strGuarantorBlock = strGuarantorBlock & strTemp & MakeString(3 - Len(strTemp), cstrTrailingCharForStrings)
'            strGuarantorBlock = strGuarantorBlock & MakeString(2, m_strFieldSeparator)
'        Else
'            strGuarantorBlock = strGuarantorBlock & strTemp
'            strGuarantorBlock = strGuarantorBlock & MakeString(2, m_strFieldSeparator)
'        End If
' ik_E2EM00002919_ends
        
        'Company Name
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & Space(30)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If

        'MSla 23/09/2004 E2EM00001422 - End
 
        Set xmlTable = Nothing
        Set xmlFormat = Nothing
        
        'ADDRESS
        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMERADDRESS[@ADDRESSTYPE='1']")
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//GUARANTOR/ADDRESS")
        
' ik_E2EM00002919 - use standard address routine
        strGuarantorBlock = strGuarantorBlock & BuildAddress(xmlTable, xmlFormat, True)
' ik_E2EM00002919_ends
        
' ik_E2EM00002919 Person spare 1, first forename for Guarantor
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & MakeString(15, cstrTrailingCharForStrings) 'Telephone
            strGuarantorBlock = strGuarantorBlock & String(8, cstrTrailingCharForStrings)  'Valuer Code
        Else
            strGuarantorBlock = strGuarantorBlock & MakeString(2, m_strFieldSeparator)
        End If
    
        'Person Spare 1 - Forename
        
        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMER")
        strTemp = xmlGetAttributeText(xmlTable, "FIRSTFORENAME")
        If m_blnFixedLengthRecords Then
            strTemp = strTemp & String(15, " ")
        Else
            strTemp = Trim(strTemp)
        End If
        strGuarantorBlock = strGuarantorBlock & Left(strTemp, 15) & m_strFieldSeparator
            
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & String(15, cstrLeadingCharForNumbers)   'Person Spare 3
            strGuarantorBlock = strGuarantorBlock & String(10, cstrLeadingCharForNumbers)  'Person Spare 2
        Else
            strGuarantorBlock = strGuarantorBlock & MakeString(2, m_strFieldSeparator)
        End If
' ik_E2EM00002919_ends
        
        strGuarantorBlock = strGuarantorBlock & "A"
        
        strNewHeader = RebuildHeader(strRecordHeader)
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strGuarantorBlock & cstrRecordSeparator
    End If
    
    GuarantorDataBlock = True
    
GuarantorDataBlockExit:
    
    Set xmlGuarantor = Nothing
    Set xmlGuarantorList = Nothing
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function SecurityAddressDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo SecurityAddressDataBlockExit
    
    Const cstrFunctionName As String = "SecurityAddressDataBlock"
    
    Dim xmlNewProperty As IXMLDOMElement
    Dim xmlFormat As IXMLDOMElement
    
    Dim strSecurityAddrBlock As String
    Dim strTemp As String
    Dim strValueId As String
    Dim strNewHeader As String
    
    If m_blnNewAdvance Or m_blnBuyToLet Then
        Set xmlNewProperty = vxmlResponse.selectSingleNode(".//NEWPROPERTY")
    Else
        'Get valueid for Current/Home Address
        strValueId = GetFirstComboValueId("CustomerAddressType", "H")
        Set xmlNewProperty = vxmlResponse.selectSingleNode(".//CUSTOMER/CUSTOMERADDRESS[@ADDRESSTYPE='" & strValueId & "']")
    End If
    
    
    If Not xmlNewProperty Is Nothing Then
        
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//NEWPROPERTY")
        
        strSecurityAddrBlock = "SA" & m_strFieldSeparator
    
        ' Title, Initials, Surname, DateofBirth, Sex - not applicable to NewProperty..
        If m_blnFixedLengthRecords Then
            strSecurityAddrBlock = strSecurityAddrBlock & String(4, cstrTrailingCharForStrings)
            strSecurityAddrBlock = strSecurityAddrBlock & String(2, cstrTrailingCharForStrings)
            strSecurityAddrBlock = strSecurityAddrBlock & String(20, cstrTrailingCharForStrings)
            strSecurityAddrBlock = strSecurityAddrBlock & String(8, cstrLeadingCharForNumbers)
            strSecurityAddrBlock = strSecurityAddrBlock & String(1, cstrTrailingCharForStrings)
            strSecurityAddrBlock = strSecurityAddrBlock & String(30, cstrTrailingCharForStrings)
        Else
            strSecurityAddrBlock = strSecurityAddrBlock & MakeString(6, m_strFieldSeparator)
        End If
        
        'Address
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//NEWPROPERTY/ADDRESS")
                
'   ik_E2EM00002919 - Left(x,35) with separator appended could lose field separator
        'Address Line 1
        strSecurityAddrBlock = strSecurityAddrBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlNewProperty, xmlFormat, "FLATNUMBER", False)) & " " & _
                Trim$(GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENAME", False))), 35) & _
                m_strFieldSeparator
                
        'Address Line 2
        strSecurityAddrBlock = strSecurityAddrBlock & _
                Left$(Trim$(Trim(GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENUMBER", False)) & _
                " " & Trim$(GetFieldValueString(xmlNewProperty, xmlFormat, "STREET", False))), 35) & _
                m_strFieldSeparator
'   ik_E2EM00002919_ends
        
        'Address Line 3, 4, 5, 6
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "DISTRICT")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "TOWN")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "COUNTY")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "POSTCODE")
        
        If m_blnFixedLengthRecords Then
            
            ' Telephone Number
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrTrailingCharForStrings)
                    
            'Valuer Code
            
            strSecurityAddrBlock = strSecurityAddrBlock & String(8, cstrTrailingCharForStrings)
            
            'Years at address, Income are zeros
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrLeadingCharForNumbers)
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrLeadingCharForNumbers)
            
            strSecurityAddrBlock = strSecurityAddrBlock & String$(10, cstrLeadingCharForNumbers)
            
        Else
            
            strSecurityAddrBlock = strSecurityAddrBlock & MakeString(5, m_strFieldSeparator)
            
        End If
        
        ' Add/Update
        strSecurityAddrBlock = strSecurityAddrBlock & "A"
        
        ''BBG454 - remove bank details from header
        strNewHeader = RebuildHeader(strRecordHeader)
                    
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strSecurityAddrBlock & cstrRecordSeparator
        
    End If
        
    SecurityAddressDataBlock = True
    
SecurityAddressDataBlockExit:
    
    Set xmlNewProperty = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function SolicitorDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo SolicitorDataBlockExit
    
    Const cstrFunctionName As String = "SolicitorDataBlock"
    
    Dim xmlSolicitor As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    Dim xmlAddressOrTelephone As IXMLDOMNode
    
    Dim strSolicitorBlock As String
    Dim strTemp As String
    Dim strNewHeader As String
    
' IK BBG1810 25/11/2004 - use applicant solicitor
    Set xmlSolicitor = vxmlResponse.selectSingleNode(".//LEGALREP[not(@SEPARATELEGALREPRESENTATIVE) or @SEPARATELEGALREPRESENTATIVE = 0]/NAMEANDADDRESSDIRECTORY")
    If xmlSolicitor Is Nothing Then
        Set xmlSolicitor = vxmlResponse.selectSingleNode(".//LEGALREP[not(@SEPARATELEGALREPRESENTATIVE) or @SEPARATELEGALREPRESENTATIVE = 0]/THIRDPARTY")
    End If
' IK BBG1810 25/11/2004 - ends
    
    If Not xmlSolicitor Is Nothing Then
        
        strSolicitorBlock = "ST" & m_strFieldSeparator
        
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP")
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
        If m_blnFixedLengthRecords Then
            strSolicitorBlock = strSolicitorBlock & String(4, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(2, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(20, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(8, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(1, cstrTrailingCharForStrings)
        Else
            strSolicitorBlock = strSolicitorBlock & MakeString(5, m_strFieldSeparator)
        End If
        
        'Company Name
        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "COMPANYNAME")
                
        'Address
        Set xmlFormat = Nothing
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP/ADDRESS")
    
        Set xmlAddressOrTelephone = xmlGetNode(xmlSolicitor, "ADDRESS")
    
        strSolicitorBlock = strSolicitorBlock & BuildAddress(xmlAddressOrTelephone, xmlFormat, True)
        
        'Telephone Number
        Set xmlFormat = Nothing
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP/CONTACTDETAILS")
        
        strTemp = GetFirstComboValueId("ContactTelephoneUsage", "W")
        Set xmlAddressOrTelephone = xmlGetNode(xmlSolicitor, "CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE='" & strTemp & "']")
        
'   ik_E2EM00002919 - Left(x,15) with separator appended could lose field separator
        strSolicitorBlock = strSolicitorBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlAddressOrTelephone, xmlFormat, "AREACODE", False)) & " " & _
                Trim$(GetFieldValueString(xmlAddressOrTelephone, xmlFormat, "TELENUMBER", False))), 15) & _
                m_strFieldSeparator
'   ik_E2EM00002919_ends
                
        'Valuer code, Years at address, income, purchase price are not applicable.
        If m_blnFixedLengthRecords Then
            strSolicitorBlock = strSolicitorBlock & String(8, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(15, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(15, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(10, cstrLeadingCharForNumbers)
        Else
            strSolicitorBlock = strSolicitorBlock & MakeString(4, m_strFieldSeparator)
        End If
        
        'Add/Update
        strSolicitorBlock = strSolicitorBlock & "A"
            
        strNewHeader = RebuildHeader(strRecordHeader)
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strSolicitorBlock & cstrRecordSeparator
        
        
    End If
                                    
    SolicitorDataBlock = True
                    
SolicitorDataBlockExit:
    
    Set xmlSolicitor = Nothing
    Set xmlFormat = Nothing
    Set xmlAddressOrTelephone = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
                    
End Function
Private Function EmployerDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode, _
                                    ByVal vxmlApplicantList As IXMLDOMNodeList) As Boolean
'-------------------------------------------------------------------------------
' For each applicant, build Employment records using the latest employment data
'--------------------------------------------------------------------------------
        
    On Error GoTo EmployerDataBlockExit
    
    Const cstrFunctionName As String = "EmployerDataBlock"
    
    Dim xmlApplicantList As IXMLDOMNodeList
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode
    Dim xmlEmployer As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    Dim xmlEmploymentList As IXMLDOMNodeList
    
    Dim intNoOfApplicants As Integer
    Dim intCounter As Integer
    Dim intValue As Integer
    Dim intCustomerOrder As Integer
    
    Dim strEmployerBlock As String
    Dim strTemp As String
    Dim strNewHeader As String
    
    'Find all current employments
    Set xmlEmploymentList = vxmlResponse.selectNodes(".//CUSTOMER/INCOME/EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]")
    For Each xmlEmployment In xmlEmploymentList
        
        Set xmlEmployer = xmlGetNode(xmlEmployment, ".//EMPLOYER")
        
        If Not xmlEmployer Is Nothing Then
            
            intCounter = intCounter + 1
            
            If intCounter <= 4 Then
                'Record Type
                strEmployerBlock = "E" & Left(Trim(Str(intCounter)), 1) & m_strFieldSeparator
                
                ' Title, Initials, Surname, DateOfBirth, Sex are not applicable
                If m_blnFixedLengthRecords Then
                    strEmployerBlock = strEmployerBlock & String(5, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(3, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(20, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(10, cstrLeadingCharForNumbers)
                    strEmployerBlock = strEmployerBlock & String(1, cstrTrailingCharForStrings)
                Else
                    strEmployerBlock = strEmployerBlock & MakeString(5, m_strFieldSeparator)
                End If
                
                'Company Name
                Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER")
                Set xmlEmployer = xmlGetMandatoryNode(xmlEmployment, ".//EMPLOYER") 'BMIDS00279 MDC 01/08/2002
                strEmployerBlock = strEmployerBlock & GetFieldValueString(xmlEmployer, xmlFormat, "COMPANYNAME")
                
                'Address
                strEmployerBlock = strEmployerBlock & BuildAddress(xmlEmployer, xmlFormat)
                
                'Telephone Number
                Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/CONTACTDETAILS")
    
                intValue = xmlGetAttributeAsInteger(xmlEmployer, "USAGE")
                If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
' ik_E2EM00002919
                    strTemp = Left$(Trim$(Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "AREACODE", False)) _
                                & " " & Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "TELENUMBER", False))), 15)
' ik_E2EM00002919_ends
                    strEmployerBlock = strEmployerBlock & Trim$(strTemp)
                    
                    If m_blnFixedLengthRecords Then
                        strEmployerBlock = strEmployerBlock & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings)
                    Else
                        strEmployerBlock = strEmployerBlock & m_strFieldSeparator
                    End If
                Else
                    If m_blnFixedLengthRecords Then
                        strEmployerBlock = strEmployerBlock & Space$(15) & cstrTrailingCharForStrings
                    Else
                        strEmployerBlock = strEmployerBlock & m_strFieldSeparator
                    End If
                End If
                
                'Valuers code, Person Spare 1, Person Spare 2, Person Spare 3 not required
                If m_blnFixedLengthRecords Then
                    strEmployerBlock = strEmployerBlock & String(8, cstrTrailingCharForStrings) 'Valuer Code
                    strEmployerBlock = strEmployerBlock & String(15, cstrLeadingCharForNumbers) 'Person Spare 1
                    strEmployerBlock = strEmployerBlock & String(15, cstrLeadingCharForNumbers) 'Person Spare 2
                    strEmployerBlock = strEmployerBlock & String(10, cstrLeadingCharForNumbers) 'Person Spare 3
                Else
                    strEmployerBlock = strEmployerBlock & MakeString(4, m_strFieldSeparator)
                End If
                
                'Add/Update
                strEmployerBlock = strEmployerBlock & "A"  'Add/Update
                 
                'BBG454 - remove bank details from header
                strNewHeader = RebuildHeader(strRecordHeader)
                
                strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strEmployerBlock & cstrRecordSeparator
                   
                'Determine if Self Employed
                intValue = CSafeInt(xmlGetMandatoryAttributeAsLong(xmlEmployment, "EMPLOYMENTSTATUS"))
' BBG1758 IK 04/11/2004
                If IsValidationType("EmploymentStatus", intValue, "S") _
                Or IsValidationType("EmploymentStatus", intValue, "C") _
                Or IsValidationType("EmploymentStatus", intValue, "EC") _
                Then
' BBG1758 IK 04/11/2004 ends
                    'Self Employed
                    If Not AccountantDataBlock(xmlEmployment, vxmlFormatTemplate) Then
                        EmployerDataBlock = False
                        Exit For
                    End If
                End If
            End If
        End If
    Next
                                        
    EmployerDataBlock = True
EmployerDataBlockExit:
    
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    Set xmlEmployment = Nothing
    Set xmlEmployer = Nothing
    Set xmlFormat = Nothing
    Set xmlEmploymentList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
                                        
End Function
Private Function ExistingMortgageDataBlock(ByVal xmlResponse As IXMLDOMNode, _
                                            ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
                                            
    On Error GoTo ExistingMortgageDataBlockExit

    Const cstrFunctionName As String = "ExistingMortgageDataBlock"
    
    Dim xmlCustomers As IXMLDOMNodeList
    Dim xmlFormat As IXMLDOMNode
    Dim xmlMortgageAccountList As IXMLDOMNodeList
    Dim xmlAccountRelationship As IXMLDOMNode
    Dim xmlMortgage As IXMLDOMNode
    Dim xmlAddress As IXMLDOMNode
    Dim xmlMortgageNode As IXMLDOMNode
    Dim xmlThirdPartyNode As IXMLDOMNode
    Dim xmlAddressDetails As IXMLDOMNode
    Dim xmlContactDetails As IXMLDOMNode
    
    Dim strAccounts() As String
    Dim strCompanyName As String
    Dim strTelephone As String
    Dim strMortgageGUID As String
    Dim strAccountGUID As String
    Dim strThirdPartyGUID As String
    Dim strExistingAccount As String
    Dim strAddress1 As String
    Dim strAddress2 As String
    Dim strNewHeader As String
    
    Dim intAccountLoop As Integer
    Dim intNumAccounts As Integer
    
    Dim bAccountProcessed As Boolean
    
    Set xmlCustomers = xmlResponse.selectNodes(".//APPLICATION/APPLICATIONCUSTOMERROLE/CUSTOMER")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EXISTINGMORTGAGE")

    If Not xmlCustomers Is Nothing Then
        intNumAccounts = 0
        For Each xmlAccountRelationship In xmlCustomers
            Set xmlMortgageAccountList = xmlAccountRelationship.selectNodes(".//MORTGAGEACCOUNT")
            For Each xmlMortgage In xmlMortgageAccountList
            '' read off all mortgage entries in each account relationship
                Set xmlAddress = xmlMortgage.selectSingleNode(".//ADDRESS")
                ReDim Preserve strAccounts(intNumAccounts)
                strAccountGUID = xmlGetAttributeText(xmlMortgage, "ACCOUNTGUID")
            
                '' see if we have processed this account already, if so skip it
                bAccountProcessed = False
                For intAccountLoop = 0 To UBound(strAccounts) - 1
                    If strAccountGUID = strAccounts(intAccountLoop) Then
                        bAccountProcessed = True
                    End If
                Next
            
                If bAccountProcessed = False Then
                    strAccounts(intNumAccounts) = strAccountGUID
                    intNumAccounts = intNumAccounts + 1
                    '' start getting the address details, this could be either in the NameAndAddressDirectory
                    '' or held in a ThirdParty node, so check them both
                    Set xmlMortgageNode = xmlMortgage.selectSingleNode(".//ACCOUNT/NAMEANDADDRESSDIRECTORY")
                    If Not xmlMortgageNode Is Nothing Then
                        strMortgageGUID = xmlGetAttributeText(xmlMortgageNode, "DIRECTORYGUID")
                        strCompanyName = xmlGetAttributeText(xmlMortgageNode, "COMPANYNAME")
                        Set xmlAddressDetails = xmlMortgageNode.selectSingleNode(".//ADDRESS")
                    End If
                    
                    Set xmlThirdPartyNode = xmlMortgage.selectSingleNode(".//ACCOUNT/THIRDPARTY")
                    If Not xmlThirdPartyNode Is Nothing Then
                        strThirdPartyGUID = xmlGetAttributeText(xmlThirdPartyNode, "THIRDPARTYGUID")
                        strCompanyName = xmlGetAttributeText(xmlThirdPartyNode, "COMPANYNAME")
                        Set xmlAddressDetails = xmlThirdPartyNode.selectSingleNode(".//ADDRESS")
                    End If
                End If
            
                If Not xmlAddressDetails Is Nothing Then
                    '' ok finally found some address details, now populate the Hunter entry
                    strExistingAccount = "ME" & m_strFieldSeparator
                
                    If m_blnFixedLengthRecords Then
                        strExistingAccount = strExistingAccount & String(5, cstrTrailingCharForStrings) '' Title
                        strExistingAccount = strExistingAccount & String(3, cstrTrailingCharForStrings) '' Initials
                        strExistingAccount = strExistingAccount & String(20, cstrTrailingCharForStrings) '' Surname
                        strExistingAccount = strExistingAccount & String(10, cstrLeadingCharForNumbers) '' DOB
                        strExistingAccount = strExistingAccount & String(1, cstrTrailingCharForStrings) '' Sex
                    Else
                        strExistingAccount = strExistingAccount & MakeString(5, m_strFieldSeparator)
                    End If
            
                    strExistingAccount = strExistingAccount & Trim$(strCompanyName) & m_strFieldSeparator
                    
        
' ik_E2EM00002919 - use standard address routine
                    strExistingAccount = strExistingAccount & BuildAddress(xmlAddressDetails, xmlFormat, True)
' ik_E2EM00002919_ends
                                                        
' ik_E2EM00002919 03/12/2004 - remove duplicate values
'                    strExistingAccount = strExistingAccount & GetFieldValueString(xmlAddressDetails, xmlFormat, "DISTRICT")
'                    strExistingAccount = strExistingAccount & GetFieldValueString(xmlAddressDetails, xmlFormat, "TOWN")
'                    strExistingAccount = strExistingAccount & GetFieldValueString(xmlAddressDetails, xmlFormat, "COUNTY")
'                    strExistingAccount = strExistingAccount & GetFieldValueString(xmlAddressDetails, xmlFormat, "POSTCODE")
' ik_E2EM00002919 03/12/2004 - ends
                
                    Set xmlContactDetails = xmlMortgage.selectSingleNode(".//CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=10]")
                    If Not xmlContactDetails Is Nothing Then
'   ik_E2EM00002919 - Left(x,15) with separator appended could lose field separator
                        strExistingAccount = strExistingAccount & _
                            Left$(Trim$(Trim$(GetFieldValueString(xmlContactDetails, xmlFormat, "AREACODE", False)) _
                            & " " & Trim$(GetFieldValueString(xmlContactDetails, xmlFormat, "TELENUMBER", False))), 15) _
                            & m_strFieldSeparator
'   ik_E2EM00002919_ends
                    Else
                        strExistingAccount = strExistingAccount & MakeString(1, m_strFieldSeparator)
                    End If
                
                    If m_blnFixedLengthRecords Then
                        strExistingAccount = strExistingAccount & String(8, cstrTrailingCharForStrings) '' valuer code
                        strExistingAccount = strExistingAccount & String(15, cstrTrailingCharForStrings) '' person spare 1
                        strExistingAccount = strExistingAccount & String(15, cstrTrailingCharForStrings) '' person spare 2
                        strExistingAccount = strExistingAccount & String(10, cstrTrailingCharForStrings) '' person spare 3
                    Else
                        strExistingAccount = strExistingAccount & MakeString(4, m_strFieldSeparator)
                    End If
                                                                       
                    '' Add/Update
                    strExistingAccount = strExistingAccount & "A"
                    
                    ''BBG454 - remove bank details from header
                    strNewHeader = RebuildHeader(strRecordHeader)
                    strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                               m_strHeaderSuffix & strExistingAccount & cstrRecordSeparator
                
                End If
            Next
        Next
    End If
   
    ExistingMortgageDataBlock = True

ExistingMortgageDataBlockExit:
    Set xmlMortgageAccountList = Nothing
    Set xmlFormat = Nothing
    Set xmlAccountRelationship = Nothing
    Set xmlMortgage = Nothing
    Set xmlAddress = Nothing
    Set xmlMortgageNode = Nothing
    Set xmlThirdPartyNode = Nothing
    Set xmlAddressDetails = Nothing
    Set xmlContactDetails = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function
Private Function KeyHolderDataBlock(ByVal xmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
                                    
    On Error GoTo KeyHolderDataBlockExit
                                    
    Const cstrFunctionName As String = "KeyHolderDataBlock"
    
    Dim xmlKeyHolder As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    
    Dim strKeyHolderBlock As String
    Dim strContactName As String
    Dim strTeleNumber As String
    Dim strTelephone As String
    Dim strNewHeader As String
    
    Dim intArrangement As Integer
    
    Set xmlKeyHolder = xmlResponse.selectSingleNode(".//APPLICATION/NEWPROPERTY")
        
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//KEYHOLDER")
    
    If Not xmlKeyHolder Is Nothing Then
        
        intArrangement = xmlGetAttributeAsInteger(xmlKeyHolder, "ARRANGEMENTSFORACCESS")
        strContactName = xmlGetAttributeText(xmlKeyHolder, "ACCESSCONTACTNAME")
        strTeleNumber = xmlGetAttributeText(xmlKeyHolder, "ACCESSTELEPHONENUMBER")
        
        strKeyHolderBlock = "KH" & m_strFieldSeparator

        ' If the arrangement access is "VENDOR/APPLICANT"
        If intArrangement = 1 Then
        
            Set xmlKeyHolder = xmlResponse.selectSingleNode(".//APPLICATION/NEWPROPERTY/NEWPROPERTYVENDOR")
            strKeyHolderBlock = strKeyHolderBlock & GetFieldValueString(xmlKeyHolder, xmlFormat, "CONTACTTITLE_TEXT")
            strKeyHolderBlock = strKeyHolderBlock & GetFieldValueString(xmlKeyHolder, xmlFormat, "CONTACTFORENAME")
            strKeyHolderBlock = strKeyHolderBlock & GetFieldValueString(xmlKeyHolder, xmlFormat, "CONTACTSURNAME")
            
            If m_blnFixedLengthRecords Then
                strKeyHolderBlock = strKeyHolderBlock & String(10, cstrTrailingCharForStrings) '' DOB
                strKeyHolderBlock = strKeyHolderBlock & String(1, cstrTrailingCharForStrings) '' sex
                strKeyHolderBlock = strKeyHolderBlock & String(30, cstrTrailingCharForStrings) '' company name
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr1
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr2
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr3
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr4
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr5
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' addr6
            Else
                strKeyHolderBlock = strKeyHolderBlock & MakeString(9, m_strFieldSeparator)
            End If
            
' ik_E2EM00002919
            strTelephone = GetFieldValueString(xmlKeyHolder, xmlFormat, "AREACODE", False)
            strTelephone = strTelephone & " " & Trim$(GetFieldValueString(xmlKeyHolder, xmlFormat, "TELENUMBER", False))
            strKeyHolderBlock = strKeyHolderBlock & Left(Trim(strTelephone), 15) & m_strFieldSeparator
' ik_E2EM00002919_ends
            
        Else
            'If Others
            If m_blnFixedLengthRecords Then
                strKeyHolderBlock = strKeyHolderBlock & String(5, cstrTrailingCharForStrings) '' Title
                strKeyHolderBlock = strKeyHolderBlock & String(3, cstrTrailingCharForStrings) '' Initials
                strKeyHolderBlock = strKeyHolderBlock & String(20, cstrTrailingCharForStrings) '' Surname
                strKeyHolderBlock = strKeyHolderBlock & String(10, cstrLeadingCharForNumbers) '' DOB
                strKeyHolderBlock = strKeyHolderBlock & String(1, cstrTrailingCharForStrings) '' Sex
            Else
                strKeyHolderBlock = strKeyHolderBlock & MakeString(5, m_strFieldSeparator)
            End If
            
            strKeyHolderBlock = strKeyHolderBlock & GetFieldValueString(xmlKeyHolder, xmlFormat, "ACCESSCONTACTNAME")

            If m_blnFixedLengthRecords Then
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' Addr1
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' Addr2
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' Addr3
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrLeadingCharForNumbers) '' Addr4
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' Addr5
                strKeyHolderBlock = strKeyHolderBlock & String(35, cstrTrailingCharForStrings) '' Addr6
            Else
                strKeyHolderBlock = strKeyHolderBlock & MakeString(6, m_strFieldSeparator)
            End If
            
' ik_E2EM00002919
            strTelephone = Trim(GetFieldValueString(xmlKeyHolder, xmlFormat, "AREACODE", False))
            strTelephone = strTelephone & " " & Trim(GetFieldValueString(xmlKeyHolder, xmlFormat, "ACCESSTELEPHONENUMBER", False))
            strKeyHolderBlock = strKeyHolderBlock & Left(Trim(strTelephone), 15) & m_strFieldSeparator
' ik_E2EM00002919_ends
            
        End If
        
        If m_blnFixedLengthRecords Then
            strKeyHolderBlock = strKeyHolderBlock & String(8, cstrTrailingCharForStrings) '' Valuer Code
            strKeyHolderBlock = strKeyHolderBlock & String(15, cstrTrailingCharForStrings) '' person spare 1
            strKeyHolderBlock = strKeyHolderBlock & String(15, cstrTrailingCharForStrings) '' person spare 2
            strKeyHolderBlock = strKeyHolderBlock & String(10, cstrTrailingCharForStrings) '' person spare 3
        Else
            strKeyHolderBlock = strKeyHolderBlock & MakeString(4, m_strFieldSeparator)
        End If
            
        ' Add/Update
        strKeyHolderBlock = strKeyHolderBlock & "A"
        
        'remove bank details from header
        strNewHeader = RebuildHeader(strRecordHeader)
                
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                       m_strHeaderSuffix & strKeyHolderBlock & cstrRecordSeparator
            
    End If
        
    KeyHolderDataBlock = True
    
KeyHolderDataBlockExit:

    Set xmlKeyHolder = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
End Function
Private Function CreateFinancialInstitutionDataBlock(ByVal xmlInstitution As IXMLDOMNode, _
                                                     ByVal vxmlFormatTemplate As IXMLDOMNode) As String
            
    On Error GoTo CreateFinancialInstitutionDataBlockExit

    Const cstrFunctionName As String = "CreateFinancialInstitutionDataBlock"
            
            
    Dim strBankBlock As String
    Dim strAddr1  As String
    Dim strAddr2  As String
    Dim strTelephone  As String
    
    Dim xmlFormat  As IXMLDOMNode
    
    Dim intUsage  As Integer
    
    intUsage = xmlGetAttributeAsInteger(xmlInstitution, "USAGE")
    
    If m_blnFixedLengthRecords Then
        strBankBlock = strBankBlock & String(5, cstrTrailingCharForStrings) 'Title
        strBankBlock = strBankBlock & String(3, cstrTrailingCharForStrings) 'Initials
        strBankBlock = strBankBlock & String(20, cstrTrailingCharForStrings) 'Surname
        strBankBlock = strBankBlock & String(10, cstrTrailingCharForStrings) 'DOB
        strBankBlock = strBankBlock & String(1, cstrTrailingCharForStrings) 'Sex
    Else
        strBankBlock = strBankBlock & MakeString(5, m_strFieldSeparator)
    End If

    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//FINANCIALINSTITUTION")

    strBankBlock = strBankBlock & GetFieldValueString(xmlInstitution, xmlFormat, "COMPANYNAME")
        
' ik_E2EM00002919 - use standard address routine
    strBankBlock = strBankBlock & BuildAddress(xmlInstitution, xmlFormat, True)
' ik_E2EM00002919_ends

    'Area Code & Tele Number
    If intUsage > 0 Then
        If IsValidationType("ContactTelephoneUsage", intUsage, "W") Then
' ik_E2EM00002919
            strTelephone = GetFieldValueString(xmlInstitution, xmlFormat, "AREACODE", False)
            strTelephone = Trim(strTelephone) & " " & Trim(GetFieldValueString(xmlInstitution, xmlFormat, "TELENUMBER", False))
' ik_BBG1808
            strBankBlock = strBankBlock & Left(Trim(strTelephone), 15) & m_strFieldSeparator
' ik_BBG1808_ends
' ik_E2EM00002919_ends
        End If
    Else
        strBankBlock = strBankBlock & MakeString(1, m_strFieldSeparator)
    End If
    
    If m_blnFixedLengthRecords Then
        strBankBlock = strBankBlock & String(8, cstrTrailingCharForStrings) 'VALUER CODE
        strBankBlock = strBankBlock & String(15, cstrTrailingCharForStrings) 'Person Spare 1
        strBankBlock = strBankBlock & String(15, cstrTrailingCharForStrings) 'Person Spare 2
        strBankBlock = strBankBlock & String(10, cstrTrailingCharForStrings) 'Person Spare 3
    Else
        strBankBlock = strBankBlock & MakeString(4, m_strFieldSeparator)
    End If

    '' Add/Update
    strBankBlock = strBankBlock & "A"
    
    CreateFinancialInstitutionDataBlock = strBankBlock
    
CreateFinancialInstitutionDataBlockExit:
    
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

'-----------------------------------------------------------
' BBG05 - new datablock for Previous Addresses contacts
'-----------------------------------------------------------
Private Function PreviousAddressDataBlock(ByVal xmlResponse As IXMLDOMNode, _
                                            ByVal vxmlFormatTemplate As IXMLDOMNode, _
                                            ByVal vxmlApplicantList As IXMLDOMNodeList) As Boolean
                                               
    On Error GoTo PreviousAddressDataBlockExit
                                            
    Const cstrFunctionName As String = "PreviousAddressDataBlock"
    
    Dim xmlApplicant As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    Dim xmlCustomerDetails As IXMLDOMNode
    Dim xmlPrevAddressList As IXMLDOMNodeList
    Dim xmlAddress As IXMLDOMNode
    
    Dim intApplicants As Integer
    Dim intCurrentApplicant As Integer
    Dim intCustRoleType As Integer
    
    Dim strPrevAddr As String
    Dim strForename As String
    Dim strAddr1 As String
    Dim strAddr2 As String
    Dim strNewHeader As String
    
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//PREVIOUSADDRESS")
                                            
    '' find out how many applicants there are, limit to 4 if more.
    intApplicants = IIf(vxmlApplicantList.length > 4, 4, vxmlApplicantList.length)
    If intApplicants > 0 Then
        For intCurrentApplicant = 1 To intApplicants
            Set xmlApplicant = vxmlApplicantList.Item(intCurrentApplicant - 1)
            
            '' only want previous addresses so retrieve nodes with a filter
            Set xmlPrevAddressList = xmlApplicant.selectNodes(".//CUSTOMERADDRESS[@ADDRESSTYPE='3']")
                        
            'MSla 23/09/2004 E2EM00001422
            'If Not xmlPrevAddressList Is Nothing Then
            If xmlPrevAddressList.length > 0 Then
                '' ok got the type of customer and all their previous addresses, now build the data block up
                intCustRoleType = xmlGetAttributeAsInteger(xmlApplicant, "CUSTOMERROLETYPE")
                Set xmlCustomerDetails = xmlApplicant.selectSingleNode(".//CUSTOMER")
                
                If xmlPrevAddressList.length > 1 Then
                
                    Set xmlAddress = GetLatestAddressNode(xmlPrevAddressList)
                    
                Else
                
                    
                    Set xmlAddress = xmlPrevAddressList.Item(0)
                    
                End If
            'MSla 23/09/2004 E2EM00001422 - End
            strPrevAddr = "P" & CStr(intCurrentApplicant) & m_strFieldSeparator
            
            If intCustRoleType = 1 Then '' Individual
            
                strPrevAddr = strPrevAddr & GetFieldValueString(xmlCustomerDetails, xmlFormat, "TITLE_TEXT")
' ik_E2EM00002919
                strForename = GetFieldValueString(xmlCustomerDetails, xmlFormat, "FIRSTFORENAME", False)
' ik_E2EM00002919_ends
                strForename = Trim$(strForename) & " " & GetFieldValueString(xmlCustomerDetails, xmlFormat, "SECONDFORENAME")
                strPrevAddr = strPrevAddr & Trim$(strForename)
                
                strPrevAddr = strPrevAddr & GetFieldValueString(xmlCustomerDetails, xmlFormat, "SURNAME")
                
                If m_blnFixedLengthRecords Then
                    strPrevAddr = strPrevAddr & String(1, cstrTrailingCharForStrings) 'DOB
                    strPrevAddr = strPrevAddr & String(30, cstrTrailingCharForStrings) 'sex
                    strPrevAddr = strPrevAddr & String(30, cstrTrailingCharForStrings) 'Company Name
                Else
                    strPrevAddr = strPrevAddr & MakeString(3, m_strFieldSeparator)
                End If
                
            Else '' limited company
                
                If m_blnFixedLengthRecords Then
                    strPrevAddr = strPrevAddr & String(5, cstrTrailingCharForStrings) 'Title
                    strPrevAddr = strPrevAddr & String(3, cstrTrailingCharForStrings) 'Initials
                    strPrevAddr = strPrevAddr & String(20, cstrTrailingCharForStrings) 'Surname
                    strPrevAddr = strPrevAddr & String(10, cstrTrailingCharForStrings) 'DOB
                    strPrevAddr = strPrevAddr & String(1, cstrTrailingCharForStrings) 'Sex
                Else
                    strPrevAddr = strPrevAddr & MakeString(5, m_strFieldSeparator)
                End If
                
                strPrevAddr = strPrevAddr & GetFieldValueString(xmlCustomerDetails, xmlFormat, "COMPANYNAME")
                
            End If
            
' ik_E2EM00002919
            strAddr1 = GetFieldValueString(xmlAddress, xmlFormat, "FLATNUMBER", False)
' ik_E2EM00002919_ends
            strAddr1 = strAddr1 & " " & Trim$(GetFieldValueString(xmlAddress, xmlFormat, "BUILDINGORHOUSENAME"))
            strPrevAddr = strPrevAddr & Trim$(strAddr1)
            
            'MSla 23/09/2004 E2EM00001422
' ik_E2EM00002919
            strAddr2 = GetFieldValueString(xmlAddress, xmlFormat, "BUILDINGORHOUSENUMBER", False)
' ik_E2EM00002919_ends
            strAddr2 = strAddr2 & " " & Trim$(GetFieldValueString(xmlAddress, xmlFormat, "STREET"))
            strPrevAddr = strPrevAddr & Trim$(strAddr2)
            
            strPrevAddr = strPrevAddr & GetFieldValueString(xmlAddress, xmlFormat, "DISTRICT") '& m_strFieldSeparator
            strPrevAddr = strPrevAddr & GetFieldValueString(xmlAddress, xmlFormat, "TOWN") '& m_strFieldSeparator
            strPrevAddr = strPrevAddr & GetFieldValueString(xmlAddress, xmlFormat, "COUNTY") '& m_strFieldSeparator
            strPrevAddr = strPrevAddr & GetFieldValueString(xmlAddress, xmlFormat, "POSTCODE") '& m_strFieldSeparator
            
            If m_blnFixedLengthRecords Then
                strPrevAddr = strPrevAddr & String(15, cstrTrailingCharForStrings) 'telephone
                strPrevAddr = strPrevAddr & String(8, cstrTrailingCharForStrings) 'VALUER CODE
                strPrevAddr = strPrevAddr & String(15, cstrTrailingCharForStrings) 'Person Spare 1
                strPrevAddr = strPrevAddr & String(15, cstrTrailingCharForStrings) 'Person Spare 2
                strPrevAddr = strPrevAddr & String(10, cstrTrailingCharForStrings) 'Person Spare 3
            Else
                strPrevAddr = strPrevAddr & MakeString(5, m_strFieldSeparator)
            End If
            'MSla 23/09/2004 E2EM00001422 - End
            '' Add/Update
            strPrevAddr = strPrevAddr & "A"
             
            ''BBG454 - remove bank details from header
            strNewHeader = RebuildHeader(strRecordHeader)
        
            strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strPrevAddr & cstrRecordSeparator
                        
            End If
        Next
    End If
            
    PreviousAddressDataBlock = True
    
PreviousAddressDataBlockExit:

    Set xmlApplicant = Nothing
    Set xmlFormat = Nothing
    Set xmlCustomerDetails = Nothing
    Set xmlPrevAddressList = Nothing
    Set xmlAddress = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function


Private Function FinancialInstitutionDataBlock(ByVal xmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean

    On Error GoTo FinancialInstitutionDataBlockExit

    Const cstrFunctionName As String = "FinancialInstitutionDataBlock"
    
    Dim xmlFinancialInstitutionList As IXMLDOMNodeList
    Dim xmlFormat As IXMLDOMNode
    Dim xmlInstitution As IXMLDOMNode
    
    Dim strFinancialInstitueBlock As String
    Dim strBankBlock As String
    Dim strAddr1 As String
    Dim strAddr2 As String
    Dim strTelephone As String
    Dim strFIBankDetails As String
    Dim strNKBankDetails As String
    Dim strNewHeader As String
    
    Dim intUseage As Integer
    Dim intRepayment As Integer
    Dim intUsage  As Integer
    
    Dim bFirstRecord As Boolean
    Dim blnNKBlock  As Boolean
    
    strBankBlock = ""
    strFIBankDetails = ""
    strNKBankDetails = ""

    bFirstRecord = True
    blnNKBlock = False
    
    Set xmlFinancialInstitutionList = xmlResponse.selectNodes(".//APPLICATION/BANKORBSOC")
    
    For Each xmlInstitution In xmlFinancialInstitutionList
    
        intRepayment = xmlGetAttributeAsInteger(xmlInstitution, "REPAYMENTBANKACCOUNTINDICATOR")
        
        'If Repayment indicator is true then
        If intRepayment = 1 Then
        
            intUsage = xmlGetAttributeAsInteger(xmlInstitution, "USAGE")
            
            'if the Usage is Telephone (10) or nothing then
            If intUsage = 10 Or intUsage = 0 Then
                
                'Create FI Block
                strFIBankDetails = "FI" & m_strFieldSeparator
                
                strFIBankDetails = strFIBankDetails & _
                                    CreateFinancialInstitutionDataBlock(xmlInstitution, vxmlFormatTemplate)
                
                strNewHeader = RebuildHeader(strRecordHeader)
                
                strFIBankDetails = strNewHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strFIBankDetails & cstrRecordSeparator
                'Create NK Block
                strNKBankDetails = "NK" & m_strFieldSeparator
                
                strNKBankDetails = strNKBankDetails & _
                                    CreateFinancialInstitutionDataBlock(xmlInstitution, vxmlFormatTemplate)
                 
                strNewHeader = RebuildHeader(strRecordHeader)
                
                strNKBankDetails = strNewHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strNKBankDetails & cstrRecordSeparator
            End If

        End If
        
    Next xmlInstitution
    
    '' we need the FI Block to be first but not sure where the FI entry is in the XML
    '' so build the FI and NK blocks up seperately and string them together now inthe
    '' right order
    strAllDataBlocks = strAllDataBlocks & strFIBankDetails & strNKBankDetails
    
    FinancialInstitutionDataBlock = True
    
FinancialInstitutionDataBlockExit:
    
    Set xmlFinancialInstitutionList = Nothing
    Set xmlInstitution = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Function AccountantDataBlock(ByVal xmlEmployment As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo AccountantDataBlockExit
    
    Const cstrFunctionName As String = "AccountantDataBlock"
    
    Dim xmlAccountant As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    
    Dim strAccountantBlock As String
    Dim strTemp As String
    Dim strNewHeader As String
    
    Dim intValue As Integer
    Static intCounter As Integer
    
    Set xmlAccountant = xmlEmployment.selectSingleNode(".//ACCOUNTANT")
    
    If Not xmlAccountant Is Nothing Then
        
        intCounter = intCounter + 1
        strAccountantBlock = "C" & CStr(intCounter) & m_strFieldSeparator
    
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable
        If m_blnFixedLengthRecords Then
            strAccountantBlock = strAccountantBlock & String(5, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(3, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(20, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(10, cstrLeadingCharForNumbers)
            strAccountantBlock = strAccountantBlock & String(1, cstrTrailingCharForStrings)
        Else
            strAccountantBlock = strAccountantBlock & MakeString(5, m_strFieldSeparator)
        End If
        
        'Accountant Data
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT")
        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "COMPANYNAME")
        
        ' Address
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT/ADDRESS")
        strAccountantBlock = strAccountantBlock & BuildAddress(xmlAccountant, xmlFormat)
        
        'Telephone Number
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT/CONTACTDETAILS")
        
        intValue = xmlGetAttributeAsInteger(xmlAccountant, "USAGE")
        If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
'   ik_E2EM00002919 - Left(x,15) with separator appended could lose field separator
            strAccountantBlock = strAccountantBlock & _
                        Left$(Trim$(Trim$(GetFieldValueString(xmlAccountant, xmlFormat, "AREACODE", False)) _
                        & " " & Trim$(GetFieldValueString(xmlAccountant, xmlFormat, "TELENUMBER", False))), 15) _
                        & m_strFieldSeparator
'   ik_E2EM00002919_ends
        Else
            If m_blnFixedLengthRecords Then
                strAccountantBlock = strAccountantBlock & String$(15, cstrTrailingCharForStrings)
            Else
                strAccountantBlock = strAccountantBlock & MakeString(1, m_strFieldSeparator)
            End If
        End If
        
        If m_blnFixedLengthRecords Then
            strAccountantBlock = strAccountantBlock & String(8, cstrTrailingCharForStrings) 'Valuer Code
            strAccountantBlock = strAccountantBlock & String(15, cstrLeadingCharForNumbers) 'Years at address
            strAccountantBlock = strAccountantBlock & String(15, cstrLeadingCharForNumbers) 'Income
            strAccountantBlock = strAccountantBlock & String(10, cstrLeadingCharForNumbers) 'Purchase Price
        Else
            strAccountantBlock = strAccountantBlock & MakeString(4, m_strFieldSeparator)
        End If

        'Add/Update
        strAccountantBlock = strAccountantBlock & "A"
        
        strNewHeader = RebuildHeader(strRecordHeader)
        
        ' Add this block to the string that is written to txt file
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strAccountantBlock & cstrRecordSeparator
           
    End If
    
    AccountantDataBlock = True
    
AccountantDataBlockExit:
    
    Set xmlAccountant = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Function ValuerDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo ValuerDataBlockExit
    
    Const cstrFunctionName As String = "ValuerDataBlock"
    
    Dim xmlValuer As IXMLDOMNode
    Dim xmlFormat As IXMLDOMElement
    
    Dim strValuerBlock As String
    Dim strTemp As String
    Dim strPanelNo As String
    Dim strNewHeader As String
    
    Dim intValue As Integer
    
    ' Response will have valuerInstruction records sorted in descending order. Take the first one for our purpose
    Set xmlValuer = vxmlResponse.selectSingleNode(".//VALUERINSTRUCTION")
    If Not xmlValuer Is Nothing Then
        
        strValuerBlock = "VA" & m_strFieldSeparator
    
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
        If m_blnFixedLengthRecords Then
            strValuerBlock = strValuerBlock & String(4, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(2, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(20, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(8, cstrLeadingCharForNumbers)
            strValuerBlock = strValuerBlock & String(1, cstrTrailingCharForStrings)
        Else
            strValuerBlock = strValuerBlock & MakeString(5, m_strFieldSeparator)
        End If
        
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER")
        
'   ik_E2EM00002919 Panel No is (MX) REFERENCENUMBER
        'Panel No
        strPanelNo = GetFieldValueString(xmlValuer, xmlFormat, "REFERENCENUMBER", False)
        If Len(strPanelNo) = 0 Then
            strPanelNo = GetFieldValueString(xmlValuer, xmlFormat, "VALUERPANELNO", False)
        End If
        strPanelNo = strPanelNo & m_strFieldSeparator
'   ik_E2EM00002919_ends

        'Company Name
        Set xmlValuer = vxmlResponse.selectSingleNode(".//VALUERINSTRUCTION/VALUER")
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "COMPANYNAME")
    
        'Address
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER/ADDRESS")
        
' ik_E2EM00002919 - use standard address routinr
        strValuerBlock = strValuerBlock & BuildAddress(xmlValuer, xmlFormat, True)
' ik_E2EM00002919_ends
        
        'Telephone Number
        Set xmlFormat = Nothing
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER/CONTACTDETAILS")
        
        If Not xmlValuer Is Nothing Then
            intValue = xmlGetAttributeAsInteger(xmlValuer, "USAGE")

            If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
'   ik_E2EM00002919 - Left(x,15) with separator appended could lose field separator
                strValuerBlock = strValuerBlock & _
                        Left$(Trim$(Trim$(GetFieldValueString(xmlValuer, xmlFormat, "AREACODE", False)) & " " & _
                        Trim$(GetFieldValueString(xmlValuer, xmlFormat, "TELENUMBER", False))), 15) & _
                        m_strFieldSeparator
'   ik_E2EM00002919_ends
            Else
                If m_blnFixedLengthRecords Then
                    strValuerBlock = strValuerBlock & String$(15, cstrTrailingCharForStrings)
                Else
                    strValuerBlock = strValuerBlock & m_strFieldSeparator
                End If
            End If
        Else
            If m_blnFixedLengthRecords Then
                strValuerBlock = strValuerBlock & String$(15, cstrTrailingCharForStrings)
            Else
                strValuerBlock = strValuerBlock & m_strFieldSeparator
            End If
        End If
        
        'Valuer code
        strValuerBlock = strValuerBlock & strPanelNo
        
        'Years at address, income, purchase price are not applicable.
        If m_blnFixedLengthRecords Then
            strValuerBlock = strValuerBlock & String(15, cstrLeadingCharForNumbers)
            strValuerBlock = strValuerBlock & String(15, cstrLeadingCharForNumbers)
            strValuerBlock = strValuerBlock & String(10, cstrLeadingCharForNumbers)
        Else
            strValuerBlock = strValuerBlock & MakeString(3, m_strFieldSeparator)
        End If
        
        'Add/Update
        strValuerBlock = strValuerBlock & "A"
        
        ' Add this block to the string that is written to txt file
        
        ''BBG454 - remove bank details from header block
        strNewHeader = RebuildHeader(strRecordHeader)
        
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                        m_strHeaderSuffix & strValuerBlock & cstrRecordSeparator
        
    End If
        
    ValuerDataBlock = True
    
ValuerDataBlockExit:
    
    Set xmlValuer = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function IntermediaryDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                       ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
    On Error GoTo IntermediaryDataBlockExit
        
    Const cstrFunctionName As String = "IntermediaryDataBlock"
    
    Dim xmlIntermediaryOrg As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    
    Dim strIntermediaryOrg As String
    Dim strTemp As String
    Dim strNewHeader As String
    
    Dim intValue As Integer
    
' BBG1764 IK 04/11/2004
' Set xmlIntermediaryOrg = vxmlResponse.selectSingleNode("//INTERMEDIARY/INTERMEDIARYORGANISATION[@USAGE='10']")
    Set xmlIntermediaryOrg = vxmlResponse.selectSingleNode("//INTERMEDIARY/INTERMEDIARYORGANISATION")
' BBG1764 IK 04/11/2004 ends
    
    If Not xmlIntermediaryOrg Is Nothing Then
        
        strIntermediaryOrg = "IN" & m_strFieldSeparator
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
         If m_blnFixedLengthRecords Then
            strIntermediaryOrg = strIntermediaryOrg & String(4, cstrTrailingCharForStrings) 'Title
            strIntermediaryOrg = strIntermediaryOrg & String(2, cstrTrailingCharForStrings) 'Initials
            strIntermediaryOrg = strIntermediaryOrg & String(20, cstrTrailingCharForStrings) 'Surname
            strIntermediaryOrg = strIntermediaryOrg & String(8, cstrLeadingCharForNumbers) 'DOB
            strIntermediaryOrg = strIntermediaryOrg & String(1, cstrTrailingCharForStrings) 'Sex
        Else
            strIntermediaryOrg = strIntermediaryOrg & MakeString(5, m_strFieldSeparator)
        End If
                
        ' Company Name
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, "//INTERMEDIARY/INTERMEDIARYORGANISATION")
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "NAME")
        
        Set xmlFormat = Nothing
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, "//INTERMEDIARY/INTERMEDIARYORGANISATION/ADDRESS")
        
' ik_E2EM00002919 - use standard address routinr
        strIntermediaryOrg = strIntermediaryOrg & BuildAddress(xmlIntermediaryOrg, xmlFormat, True)
' ik_E2EM00002919_ends
        
        'Contact Telephone Number
' BBG1764 IK 04/11/2004
        strTemp = ""
        Set xmlFormat = Nothing
'        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, "//INTERMEDIARY/INTERMEDIARYORGANISATION/CONTACTDETAILS")
        Set xmlFormat = xmlGetNode(vxmlFormatTemplate, "//INTERMEDIARY/INTERMEDIARYORGANISATION/CONTACTDETAILS")
        intValue = xmlGetAttributeAsInteger(xmlIntermediaryOrg, "USAGE")
        If Not xmlFormat Is Nothing Then
            If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
' ik_E2EM00002919
                strTemp = GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "AREACODE", False)
                strTemp = strTemp & " " & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "TELENUMBER")
                strIntermediaryOrg = strIntermediaryOrg & Trim$(strTemp)
' ik_E2EM00002919_ends
            End If
        End If
        If Len(strTemp) = 0 Then
            If m_blnFixedLengthRecords Then
                strIntermediaryOrg = strIntermediaryOrg & String$(15, cstrTrailingCharForStrings)
            Else
                strIntermediaryOrg = strIntermediaryOrg & MakeString(1, m_strFieldSeparator)
            End If
        End If
' BBG1764 IK 04/11/2004 ends
        
        'Valuer code, Years at address, income are not applicable.
        If m_blnFixedLengthRecords Then
            strIntermediaryOrg = strIntermediaryOrg & String(8, cstrTrailingCharForStrings) 'Valuer Code
            strIntermediaryOrg = strIntermediaryOrg & String(4, cstrLeadingCharForNumbers)  'Person Spare 1
            strIntermediaryOrg = strIntermediaryOrg & String(15, cstrLeadingCharForNumbers) 'Person Spare 2
            strIntermediaryOrg = strIntermediaryOrg & String(10, cstrLeadingCharForNumbers) 'Person Spare 3
        Else
            strIntermediaryOrg = strIntermediaryOrg & MakeString(4, m_strFieldSeparator)
        End If
        
         'Add/Update
        strIntermediaryOrg = strIntermediaryOrg & "A"
        
        'BBG454 - remove bank details from header
        strNewHeader = RebuildHeader(strRecordHeader)
        
        ' Add this block to the string that is written to txt file
        strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strIntermediaryOrg & cstrRecordSeparator
        
        
    End If
    
    IntermediaryDataBlock = True
    
IntermediaryDataBlockExit:
    
    Set xmlIntermediaryOrg = Nothing
    Set xmlFormat = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function ValuationArrangementDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
     
     On Error GoTo ValuationArrangementDataBlockExit
    
    Const cstrFunctionName As String = "ValuationArrangementDataBlock"
    
    Dim xmlNPVendor As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    
    Dim strNPVendor As String
    Dim strTemp As String
    Dim intValue As Integer
    
    Set xmlNPVendor = vxmlResponse.selectSingleNode(".//NEWPROPERTY/NEWPROPERTYVENDOR")
    
    If Not xmlNPVendor Is Nothing Then
        
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//NEWPROPERTY/NEWPROPERTYVENDOR")
        strNPVendor = ""
        
        ' Application Number
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "APPLICATIONNUMBER")
        
        'Product  - Use Fixed Value from FormatTemplate (No value from Response)
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "PRODUCT")
        strNPVendor = strNPVendor & strApplicationDate ' Application Date
        strNPVendor = strNPVendor & "ZV"  'Record Type
         
        'Data No 1
        strNPVendor = strNPVendor & "01"
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "COMPANYNAME")
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//NEWPROPERTY/NEWPROPERTYVENDOR/ADDRESS")
        
        'Data No 2
        strNPVendor = strNPVendor & "02"
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "BUILDINGORHOUSENAME")
        
        'Data No 3
        strNPVendor = strNPVendor & "03"
        strTemp = Trim(GetFieldValueString(xmlNPVendor, xmlFormat, "BUILDINGORHOUSENUMBER")) & _
                  Trim(GetFieldValueString(xmlNPVendor, xmlFormat, "FLATNUMBER")) & _
                  Trim(GetFieldValueString(xmlNPVendor, xmlFormat, "STREET"))
        If Len(strTemp) > 20 Then
            strTemp = Left(strTemp, 20)
        Else
            strTemp = strTemp & String(20 - Len(strTemp), cstrTrailingCharForStrings)
        End If
        strNPVendor = strNPVendor & Trim$(strTemp)
        
        ' Data No 4
        strNPVendor = strNPVendor & "04"
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "DISTRICT")
        
        'Data No 5
        strNPVendor = strNPVendor & "05"
        strTemp = Trim(GetFieldValueString(xmlNPVendor, xmlFormat, "TOWN")) & _
                  Trim(GetFieldValueString(xmlNPVendor, xmlFormat, "COUNTY"))
        If Len(strTemp) > 20 Then
            strTemp = Left(strTemp, 20)
        Else
            strTemp = strTemp & String(20 - Len(strTemp), cstrTrailingCharForStrings)
        End If
                  
        strNPVendor = strNPVendor & Trim$(strTemp)
        
        'Data No 6
        strNPVendor = strNPVendor & "06"
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "POSTCODE")
        
        'Data No 7 - Telephone Number
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//NEWPROPERTY/NEWPROPERTYVENDOR/CONTACTDETAILS")
        strNPVendor = strNPVendor & "07"
        
        intValue = xmlGetAttributeAsInteger(xmlNPVendor, "USAGE")
        If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
            strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "AREACODE")
            strNPVendor = strNPVendor & cstrFiller & GetFieldValueString(xmlNPVendor, xmlFormat, "TELENUMBER")
        End If
        
        strNPVendor = strNPVendor & "A"  ' Add / Update
        
        ' Add this block to the string that is written to txt file
        strAllDataBlocks = strAllDataBlocks & strNPVendor & cstrRecordSeparator
        
    End If
    
    ValuationArrangementDataBlock = True
    
ValuationArrangementDataBlockExit:
    
    Set xmlNPVendor = Nothing
    Set xmlFormat = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
#If BMIDS Then
Private Function IntroducerDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                       ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    IntroducerDataBlock
' author:       Mark Coleman
' date:         06/06/2002
' description:  Calls BMids Introducer System (through omAdmin) and builds Introducer block
' pass:         vxmlResponse        XML data extracted using Request Broker
'               vxmlFormatTemplate  XML Formatting data
' return:       True/False
'------------------------------------------------------------------------------------------
    On Error GoTo IntroducerDataBlockExit
        
    Const cstrFunctionName As String = "IntroducerDataBlock"
    
    Dim xmlIntroducer As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim objAdminBO As AdminInterfaceBO
    
    Dim strResponse As String
    Dim strIntroducer As String
    Dim sBusinessSource As String
    Dim bIsIndirect As String
    Dim strValue As String
    Dim strNewHeader As String
    
    Dim intIndex As Integer
    
    'Get index of the lowest level Introducer
    intIndex = GetLowestLevelIntroducerIndexNumber(vxmlResponse)
    If intIndex > 0 Then
        
        Set xmlApplication = xmlGetMandatoryNode(vxmlResponse, ".//APPLICATION")
        'Call omAdmin to get Introducer data from BMids Introducer System
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        
        ' Input the correct request - Start
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetIntroducerData"
        Set xmlTempNode = xmlDoc.createElement("INTRODUCER")
        xmlCopyMandatoryAttributeValue xmlApplication, xmlTempNode, "INTRODUCERIDLEVEL" & CStr(intIndex), "BMREFERENCE"
        xmlTempRequest.appendChild xmlTempNode
        
        Set objAdminBO = gobjContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        strResponse = objAdminBO.Request(xmlTempRequest.xml)
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlDoc.documentElement, , True
        
        'Find Introducer details returned
        Set xmlIntroducer = xmlGetNode(xmlDoc.documentElement, ".//INTRODUCER")
        If Not xmlIntroducer Is Nothing Then
            
            Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//INTRODUCER")
            strIntroducer = "IN"
            ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
            If m_blnFixedLengthRecords Then
                strIntroducer = strIntroducer & String(5, cstrTrailingCharForStrings)
                strIntroducer = strIntroducer & String(3, cstrTrailingCharForStrings)
                strIntroducer = strIntroducer & String(20, cstrTrailingCharForStrings)
                strIntroducer = strIntroducer & String(10, cstrLeadingCharForNumbers)
                strIntroducer = strIntroducer & String(1, cstrTrailingCharForStrings)
            Else
                strIntroducer = strIntroducer & MakeString(5, m_strFieldSeparator)
            End If
            ' Company Name
            strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "COMPANYNAME")
            
            bIsIndirect = False
            sBusinessSource = xmlGetAttributeText(xmlGetNode(vxmlResponse, "APPLICATION"), "DIRECTINDIRECTBUSINESS")
            If Len(sBusinessSource) > 0 Then
                If comboAssistEx.IsValidationType("Direct/Indirect", CInt(sBusinessSource), "I") Then
                    bIsIndirect = True
                End If
            End If
            
            If bIsIndirect Then
                ' Build the correct address lines into the xmlIntroducer data - Start
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE1", xmlGetAttributeText(xmlIntroducer, "BUILDINGNAME")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE2", xmlGetAttributeText(xmlIntroducer, "BUILDINGNO") & " " & xmlGetMandatoryAttributeText(xmlIntroducer, "STREET")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE3", xmlGetAttributeText(xmlIntroducer, "DISTRICT")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE4", xmlGetAttributeText(xmlIntroducer, "TOWN")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE5", xmlGetAttributeText(xmlIntroducer, "COUNTY")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE6", xmlGetAttributeText(xmlIntroducer, "POSTCODE")
                
                'Address
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE1")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE2")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE3")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE4")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE5")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE6")
                        
            Else
                strIntroducer = strIntroducer & MakeString(6, m_strFieldSeparator)
            End If
            
            'Contact Telephone Number
            strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "TELEPHONENO")
            
            If m_blnFixedLengthRecords Then
                strIntroducer = strIntroducer & String(8, cstrTrailingCharForStrings)   'Valuers Code
                strIntroducer = strIntroducer & String(15, cstrLeadingCharForNumbers)   'Person Spare 1
                strIntroducer = strIntroducer & String(15, cstrLeadingCharForNumbers)   'Person Spare 2
                strIntroducer = strIntroducer & String(10, cstrLeadingCharForNumbers)   'Person Spare 3
            Else
                strIntroducer = strIntroducer & MakeString(4, m_strFieldSeparator)
            End If
            
            'Add/Update
            strIntroducer = strIntroducer & "A"
                                        
            ' remove bank details from header
            strNewHeader = RebuildHeader(strRecordHeader)
                
            strAllDataBlocks = strAllDataBlocks & strNewHeader & m_strFieldSeparator & _
                                        m_strHeaderSuffix & strIntroducer & cstrRecordSeparator
                
        End If
    End If
    
    IntroducerDataBlock = True
    
IntroducerDataBlockExit:

    Set xmlIntroducer = Nothing
    Set xmlFormat = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempNode = Nothing
    Set xmlApplication = Nothing
    Set objAdminBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
#End If
Private Function GetLowestLevelIntroducerIndexNumber(ByVal vxmlResponse As IXMLDOMNode) As Integer
' header ----------------------------------------------------------------------------------
' procedure:    GetLowestLevelIntroducerIndexNumber
' author:       Mark Coleman
' date:         06/06/2002
' description:  Find the index number of the lowest level introducer
' pass:         vxmlResponse        XML data extracted using Request Broker
' return:       Index Number
'------------------------------------------------------------------------------------------
On Error GoTo GetLowestLevelIntroducerIndexNumber_Exit
Dim strFunctionName As String
strFunctionName = "GetLowestLevelIntroducerIndexNumber"
    
Dim xmlApplication As IXMLDOMNode
Dim strValue As String
Dim intLoop As Integer
Dim intIndex As Integer
    Set xmlApplication = xmlGetMandatoryNode(vxmlResponse, ".//APPLICATION")
    For intLoop = 3 To 1 Step -1
        ' MO - 28/10/2002 - BMIDS00752 - Use the correct field
        strValue = xmlGetAttributeText(xmlApplication, "INTRODUCERIDLEVEL" & CStr(intLoop))
        If Len(Trim$(strValue)) > 0 Then
            intIndex = intLoop
            Exit For
        End If
    Next
    GetLowestLevelIntroducerIndexNumber = intIndex
GetLowestLevelIntroducerIndexNumber_Exit:
    Set xmlApplication = Nothing
    errCheckError strFunctionName, TypeName(Me)
End Function
Public Function GenerateRecords(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo GenerateRecordsErr
    
    Const cstrFunctionName As String = "GenerateRecords"
    
    Dim vxmlResponse As IXMLDOMElement
    Dim xmlFormatTemplate As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlApplicantList As IXMLDOMNodeList
    
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem
    Set vxmlResponse = GetData(vstrXMLRequest)
    
    'Load template for formatting
    Set xmlDoc = LoadFormatTemplate()
    Set xmlFormatTemplate = xmlDoc.documentElement
    
    If GetMandatoryGlobalParamBoolean("HunterFormat") Then
        m_blnFixedLengthRecords = False
        m_strFieldSeparator = "|"
    Else
        m_blnFixedLengthRecords = True
        m_strFieldSeparator = ""
    End If
    
    ' Generate data blocks
    If (RecordHeader(vxmlResponse, xmlFormatTemplate)) Then
        ''BBG05 - generate applicant list outside and pass it through to datablock functions
        Set xmlApplicantList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1'or @CUSTOMERROLETYPE='3']")
        
        '' BBG05 - add applicant list to function call
        If ApplicantDataBlock(vxmlResponse, xmlFormatTemplate, xmlApplicantList) Then
            ''BBG05 - add applicant list to function call
            If PreviousAddressDataBlock(vxmlResponse, xmlFormatTemplate, xmlApplicantList) Then
                If GuarantorDataBlock(vxmlResponse, xmlFormatTemplate) Then
                    If ExistingMortgageDataBlock(vxmlResponse, xmlFormatTemplate) Then
                        If SecurityAddressDataBlock(vxmlResponse, xmlFormatTemplate) Then
                            If FinancialInstitutionDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                If KeyHolderDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                    If SolicitorDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                        If ValuerDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                            ''BBG05 - add applicant list to fucntion call
                                            If EmployerDataBlock(vxmlResponse, xmlFormatTemplate, xmlApplicantList) Then
#If BMIDS Then
                                                If IntroducerDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                                    Call WriteToTextFile
                                                End If
#End If
                                               If IntermediaryDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                                    'If ValuationArrangementDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                                        ' Write the whole string on to text file
                                                        Call WriteToTextFile
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
      
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    GenerateRecords = xmlOutDoc.xml
    
GenerateRecordsExit:
    
    Set vxmlResponse = Nothing
    Set xmlFormatTemplate = Nothing
    Set xmlDoc = Nothing
        
    Exit Function
    
GenerateRecordsErr:
        
    GenerateRecords = errCreateErrorResponse()
    GoTo GenerateRecordsExit
    
End Function

Private Function WriteToTextFile() As Boolean
    
    On Error GoTo WriteToTextFileExit
    
    Const cstrFunctionName As String = "WriteToTextFile"
    
    Dim blnAppendMode As Boolean
    
    Dim intFileNo As Integer
    
    Dim strFileName As String
    Dim strFileLoc As String
    Dim strTemp As String
    Dim strResponse As String
    
    Dim objCRM As omCRMWrapper
    
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    
    Set xmlDoc = xmlCreateDOMObject
    
    blnAppendMode = GetMandatoryGlobalParamBoolean("HunterMode")
    If blnAppendMode Then
        'Append to a single output file
        strFileName = GetMandatoryGlobalParamString("HunterFile")
    Else
        'Write to a file for the application
        strFileName = strApplicationNumber & ".txt"
    End If
    
    strFileLoc = GetMandatoryGlobalParamString("HunterOutputFolder")
    
    nMaxRetries = GetMandatoryGlobalParamAmount("LockTextFileRetries")
    nSleepMilliSecs = GetMandatoryGlobalParamAmount("LockTextFileSleep")
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("TEXTOUT")
    xmlElement.setAttribute "PATHANDFILENAME", strFileLoc & "\" & strFileName
    xmlElement.setAttribute "MAXRETRIES", nMaxRetries
    xmlElement.setAttribute "SLEEPMILLISECS", nSleepMilliSecs
    
    'WRITEMODE is either "Append" or OverWrite"
    xmlElement.setAttribute "WRITEMODE", "Append"
    xmlRequest.appendChild xmlElement
    
    Set objCRM = gobjContext.CreateInstance(gstrCRM_COMPONENT & ".omCRMWrapper")
    
    'pass string to be written direct, so we don't lose our CR/LF
    strResponse = objCRM.AddText(xmlRequest.xml, strAllDataBlocks)
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
WriteToTextFileExit:
     
    Set objCRM = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlRequest = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
Private Function MakeString(ByVal intLength As Integer, ByVal strChar As String) As String
    If intLength > 0 Then
        MakeString = String(intLength, strChar)
    Else
        MakeString = strChar
    End If
End Function
Private Function BuildAddress(ByVal vxmlTable As IXMLDOMNode, ByVal vxmlFormat As IXMLDOMNode, Optional ByVal blnEmptyvxmlTableAllowed As Boolean = False) As String
' header ----------------------------------------------------------------------------------
' procedure:    BuildAddress
' author:       Mark Coleman
' date:         06/06/2002
' description:  Builds the address part of a record
' pass:         vxmlTable  XML node holding the address details
'               vxmlFormat XML node holding formatting information
' return:       Formatted address block to be added to a record by the calling subroutine
'------------------------------------------------------------------------------------------
    On Error GoTo BuildAddress_Exit
    
    Const strFunctionName As String = "BuildAddress"
    
    Dim strAddress As String
    Dim strTemp As String
    
    
    
    If vxmlFormat Is Nothing Or (vxmlTable Is Nothing And (blnEmptyvxmlTableAllowed = False)) Then
        errThrowError strFunctionName, oeMissingElement, "Table or Format node not specified."
    End If
    
    'Build address block fragment
    
    ' Address Line 1
' ik_E2EM00002919
    strAddress = Left( _
                    Trim( _
                            Trim(GetFieldValueString(vxmlTable, vxmlFormat, "FLATNUMBER", False)) _
                            & " " & _
                            Trim(GetFieldValueString(vxmlTable, vxmlFormat, "BUILDINGORHOUSENAME", False)) _
                        ), 35)
' ik_E2EM00002919_ends
    If m_blnFixedLengthRecords Then
        strAddress = strAddress & MakeString(35 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strAddress = strAddress & m_strFieldSeparator
    End If
        
    'Address Line 2
' ik_E2EM00002919
    strTemp = Left(Trim(Trim(GetFieldValueString(vxmlTable, vxmlFormat, "BUILDINGORHOUSENUMBER", False)) & " " & _
                        Trim(GetFieldValueString(vxmlTable, vxmlFormat, "STREET", False))), 35)
' ik_E2EM00002919_ends
                        
    If m_blnFixedLengthRecords Then
        strAddress = strAddress & strTemp & MakeString(35 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strAddress = strAddress & strTemp & m_strFieldSeparator
    End If
        
    'Address Line 3, 4, 5, 6
    strAddress = strAddress & GetFieldValueString(vxmlTable, vxmlFormat, "DISTRICT")
    strAddress = strAddress & GetFieldValueString(vxmlTable, vxmlFormat, "TOWN")
    strAddress = strAddress & GetFieldValueString(vxmlTable, vxmlFormat, "COUNTY")
    strAddress = strAddress & GetFieldValueString(vxmlTable, vxmlFormat, "POSTCODE")
        
    BuildAddress = strAddress
    
BuildAddress_Exit:
    
    errCheckError strFunctionName, TypeName(Me)
    
End Function
Private Sub CreateFolderEx(ByVal strPath)
    
    On Error GoTo CreateFolderExExit
    
    Const cstrFunctionName As String = "CreateFolderEx"
    
    Dim arrFolders() As String
    Dim intLevels As Integer
    Dim strLowest As String
    Dim strWorking As String
    
    
    If m_objFSO.FolderExists(strPath) = False Then
        
        'Replace strPath, "/", "\"
        arrFolders = Split(strPath, "\")
        intLevels = UBound(arrFolders)
        
        If intLevels > 0 Then
            strLowest = arrFolders(intLevels)
            ReDim Preserve arrFolders(intLevels - 1)
            strWorking = Join(arrFolders, "\")
            CreateFolderEx strWorking
            m_objFSO.CreateFolder strWorking & "\" & strLowest
        End If
        
    End If
    
CreateFolderExExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub Class_Initialize()

    ObjectControl_Activate
    
End Sub
Private Sub ObjectControl_Activate()
    
    Set gobjContext = GetObjectContext()
    Set m_objFSO = New Scripting.FileSystemObject
    
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    
    Set gobjContext = Nothing
    Set m_objFSO = Nothing
    
End Sub
