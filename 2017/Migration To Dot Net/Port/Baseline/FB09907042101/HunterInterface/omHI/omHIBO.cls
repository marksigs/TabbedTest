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
'MARS History:
'Prog   Date        Description
'GHun   15/08/2005  MAR43 Amended BuildApplicant to change &gt; to > in XPath
'-------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext
Private strRecordHeader As String
Private strApplicationNumber As String, strApplicationDate As String
Private strProductType As String
Private strPurchasePrice As String, strAllDataBlocks As String

'BMIDS00025 MDC 23/05/2002
' The field separator should be set to "" when using fixed length records
'Const m_strFieldSeparator = "|" '""
'BMIDS00025 MDC 23/05/2002 - End
'BMIDS00260 MDC 30/07/2002
Private m_strFieldSeparator As String
'BMIDS00260 MDC 30/07/2002

Const cstrRecordSeparator = vbCrLf
Const cstrFiller = " "

'BMIDS00025 MDC 23/05/2002
Const cstrLeadingCharForNumbers = vbNullChar '"0"
Const cstrTrailingCharForStrings = vbNullChar '" "
'BMIDS00025 MDC 23/05/2002 - End

'JR - Omiplus24
Const cstrContactPhoneGrpName = "ContactTelephoneUsage"
Const cstrContactPhoneValType = "W"

'BMIDS00025 MDC 24/05/2002
Private m_blnFixedLengthRecords As Boolean
Private m_strHeaderSuffix As String
'BMIDS00025 MDC 24/05/2002 - End

'BMIDS00260 MDC 29/07/2002
Private m_objFSO As Scripting.FileSystemObject
Private m_blnNewAdvance As Boolean
'BMIDS00260 MDC 29/07/2002 - End

'BS BM0526 16/04/2003
Private m_blnBuyToLet As Boolean

'------------------------------------------------------------------------------------------
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
'INR    14/01/2003  BMIDS646 Call Compensating Resource Manager to write to text files within a
'                   transaction, allowing roll back if the transaction fails after the text write.
'JD     07/12/2004      BMIDS956 include flatnumber in address line 1, ensure telephone number and buildingorhousename are correct length
'-------------------------------------------------------------------------------------------------

Private Function GetData(ByVal vstrXMLRequest As String) As IXMLDOMElement

On Error GoTo GetDataExit

    Const cstrFunctionName As String = "GetData"
    
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
    
    Dim xmlApplication As IXMLDOMNode, xmlQuotation As IXMLDOMNode
    Dim strQuoteNumber As String
    
    '--------------------------------------------------------------------------------------------
    'if an accepted quote exists, return it else the node corresponding to the ActiveQuote number
    '--------------------------------------------------------------------------------------------
    Set xmlApplication = vxmlResponse.selectSingleNode(".//APPLICATION")
    strQuoteNumber = xmlGetAttributeText(xmlApplication, "ACCEPTEDQUOTENUMBER")
    
    If strQuoteNumber = "" Then
        strQuoteNumber = xmlGetAttributeText(xmlApplication, "ACTIVEQUOTENUMBER")
    End If
        
    ' MO - 08/11/2002 - BMIDS00752 - If there is no quotation, raise an error
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
    
    Const cstrFunctionName As String = "LoadRequest"

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
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
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
    
    Const cstrFunctionName As String = "LoadFormatTemplate"

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
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.Load (strFileName)
    
    Set LoadFormatTemplate = xmlDoc
    
    If xmlDoc.parseError.errorCode <> 0 Then
        errThrowError cstrFunctionName, oeXMLParserError
    End If
    
LoadFormatTemplateExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Function RecordHeader(ByVal vxmlResponse As IXMLDOMElement, _
                              ByVal vxmlFormatTemplate As IXMLDOMElement) As Boolean

On Error GoTo RecordHeaderExit
    
    Const cstrFunctionName As String = "RecordHeader"
    
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
    
    'BM0104 MDC 03/12/2002 - Buy To Let
    intTemp = xmlGetAttributeAsInteger(xmlApplication, "SPECIALSCHEME", "0")
    If intTemp <> 0 Then
        If IsValidationType("SpecialSchemes", intTemp, "BTL") Then
            strApplicationType = "B2L"
            'BS BM0526 16/04/2003
            m_blnBuyToLet = True
        End If
    End If
    
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
    
    strPurchasePrice = GetFieldValueString(xmlTable, xmlFormat, "PRESENTVALUATION", , , False)
    If CSafeLng(strPurchasePrice) = 0 Then
        strPurchasePrice = GetFieldValueString(xmlApplication, xmlApplicationFormat, "PURCHASEPRICEORESTIMATEDVALUE", , , False)
    End If
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
        strRecordHeader = strRecordHeader & strScheme & m_strFieldSeparator
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
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
'BMIDS00025 MDC 11/06/2002 - Add error handler & VSA
On Error GoTo BuildApplicantExit

Const cstrFunctionName As String = "ApplicantDataBlock"

Dim strApplicantBlock As String
Dim xmlApplicantList As IXMLDOMNodeList
Dim xmlApplicant As IXMLDOMNode

Dim intApplicants As Integer
Dim intCurrentApplicant As Integer
    
'BMIDS00025 MDC 11/06/2002 - End
    
    Set xmlApplicantList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']")
    intApplicants = IIf(xmlApplicantList.length > 4, 4, xmlApplicantList.length)
    
    If intApplicants > 0 Then
        For intCurrentApplicant = 1 To intApplicants
            Set xmlApplicant = xmlApplicantList.Item(intCurrentApplicant - 1)
            
            ' Build Applicant Data for this applicant
            'BMIDS00025 MDC 24/05/2002
            'strApplicantBlock = "A" + Trim(Str(intCurrentApplicant))
            strApplicantBlock = "A" + Trim(Str(intCurrentApplicant)) & m_strFieldSeparator
            'BMIDS00025 MDC 24/05/2002
            BuildApplicant xmlApplicant, vxmlFormatTemplate, strApplicantBlock
            
        Next intCurrentApplicant
    Else
        'BMIDS00142 MDC 11/06/2002
        errThrowError cstrFunctionName, oeMissingApplicantsForApplication '127     'Missing Applicants For Application
        'BMIDS00142 MDC 11/06/2002 - End
    End If
    
    ApplicantDataBlock = True

'BMIDS00025 MDC 11/06/2002 - Add error handler & VSA
BuildApplicantExit:
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
'BMIDS00025 MDC 11/06/2002 - End
    
End Function

Private Function BuildApplicant(ByVal vxmlApplicant As IXMLDOMNode, _
                                ByVal vxmlFormatTemplate As IXMLDOMNode, _
                                ByRef strApplicantBlock As String) As String

On Error GoTo BuildApplicantExit

    Const cstrFunctionName As String = "BuildApplicant"
    
    Dim xmlTable As IXMLDOMElement, xmlFormat As IXMLDOMElement
    Dim strTemp As String
        
    'BMIDS00025 MDC 07/06/2002
    Dim xmlIncome As IXMLDOMNode
    Dim xmlIncomeList As IXMLDOMNodeList
    Dim dblTotalAnnualIncome As Double
    Dim dblIncome As Double
    Dim intFrequency As Integer
    'BMIDS00025 MDC 07/06/2002 - End
        
    Dim strIncome As String 'BMIDS00285 MDC 27/09/2002
    
    ' Customer Title - Get the desc for TITLE, if it is empty use value of 'TITLEOTHER'
    Set xmlTable = xmlGetNode(vxmlApplicant, ".//CUSTOMER")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT")
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "TITLE")
    If strTemp <> "" Then
        strApplicantBlock = strApplicantBlock & strTemp
    Else
        strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "TITLEOTHER")
    End If
    
    'Initials - FirstForeName[1] + SecondForeName[1]
    'BMIDS00025 MDC 23/05/2002
'    strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME") & _
              GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME")
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", , , False) & _
              GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME")
    
    strTemp = Trim(strTemp)
    
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & strTemp & MakeString(3 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strApplicantBlock = strApplicantBlock & strTemp
    End If
    'strApplicantBlock = strApplicantBlock & strTemp & String(2 - Len(strTemp), cstrTrailingCharForStrings)
    'BMIDS00025 MDC 23/05/2002 - End

    ' Sur Name
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "SURNAME")
    
    'Date Of Birth
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "DATEOFBIRTH")

    'Sex
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "GENDER")
    
    'BMIDS00025 MDC 24/05/2002
    'Filler
    'strApplicantBlock = strApplicantBlock & String(1, cstrFiller)
    
    'Company Name
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & Space(30)
    Else
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
    End If
    'BMIDS00025 MDC 24/05/2002 - End
    
    ' Choose Current Address, and Address Format mentioned below Customer node
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    Set xmlTable = vxmlApplicant.selectSingleNode(".//CUSTOMERADDRESS[@ADDRESSTYPE='1']")
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT/ADDRESS")
        
    ' Address Line 1
    'JD BMIDS956 include flatnumber
    strApplicantBlock = strApplicantBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlTable, xmlFormat, "FLATNUMBER", , , False)) & " " & _
                Trim$(GetFieldValueString(xmlTable, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35) & m_strFieldSeparator
        
    'Address Line 2
    'BMIDS00025 MDC 23/05/2002
'    strTemp = Trim(GetFieldValueString(xmlTable, xmlFormat, "BUILDINGORHOUSENUMBER")) & Space(1) & _
              Trim(GetFieldValueString(xmlTable, xmlFormat, "STREET"))
    strTemp = Left$(Trim$(Trim(GetFieldValueString(xmlTable, xmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & " " & _
                Trim(GetFieldValueString(xmlTable, xmlFormat, "STREET"))), 35)
'    strApplicantBlock = strApplicantBlock & strTemp & String(32 - Len(strTemp), cstrTrailingCharForStrings)
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & strTemp & MakeString(35 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strApplicantBlock = strApplicantBlock & strTemp
    End If
    'BMIDS00025 MDC 23/05/2002 - End
        
    'Address Line 3, 4, 5, 6
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "DISTRICT")
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "TOWN")
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "COUNTY")
    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "POSTCODE")
    
    ' Get DateMovedIn and find years at address
    Dim dtDateMovedIn As Date, intNumberOfYears As Integer
    
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
    Set xmlTable = Nothing: Set xmlFormat = Nothing
    'BMIDS00260 MDC 29/07/2002 - Use Work or mobile if home no not present
    Dim strValueId As String
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
    'BMIDS00260 MDC 29/07/2002 - End
    
    Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//APPLICANT/CUSTOMERTELEPHONENUMBER")
    'JR - Omiplus24
    'BMIDS00025 MDC 24/05/2002
    strTemp = GetFieldValueString(xmlTable, xmlFormat, "AREACODE", , , False)
    If Len(Trim$(strTemp)) > 0 Then
        strTemp = strTemp & " "
    End If
    strTemp = strTemp & GetFieldValueString(xmlTable, xmlFormat, "TELEPHONENUMBER")
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & strTemp & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strApplicantBlock = strApplicantBlock & Left$(strTemp, 15) ' BMIDS956 limit to 15
    End If
'    strApplicantBlock = strApplicantBlock & GetFieldValueString(xmlTable, xmlFormat, "AREACODE", , , False)
'    strApplicantBlock = strApplicantBlock & Space(1) & GetFieldValueString(xmlTable, xmlFormat, "TELEPHONENUMBER")
    'BMIDS00025 MDC 24/05/2002 - End
        
    'Valuer Code
    'BMIDS00025 MDC 24/05/2002
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
    'BMIDS00260 MDC 29/07/2002 - Find all currently employment records for this applicant with earned income
'    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYEDDETAIL[@MAINSTATUS='1']/EARNEDINCOME[@EARNEDINCOMEAMOUNT > 0]")
    'MAR43 GHun Changed &gt; to >
    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]/EMPLOYEDDETAIL/EARNEDINCOME[@EARNEDINCOMEAMOUNT > 0]")
    'BMIDS00260 MDC 29/07/2002 - End
    For Each xmlIncome In xmlIncomeList
        dblIncome = xmlGetAttributeAsDouble(xmlIncome, "EARNEDINCOMEAMOUNT")
        intFrequency = CSafeInt(xmlGetMandatoryAttributeText(xmlIncome, "PAYMENTFREQUENCYTYPE"))
        dblTotalAnnualIncome = dblTotalAnnualIncome + (dblIncome * intFrequency)
    Next
    
    'BMIDS00285 MDC 27/09/2002 - Include Self Employed Income
    'MAR43 GHun Changed &gt; to >
    Set xmlIncomeList = vxmlApplicant.selectNodes(".//EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]/SELFEMPLOYEDDETAIL[@YEAR1AMOUNT > 0 or @YEAR2AMOUNT > 0 or @YEAR2AMOUNT > 0]")
    For Each xmlIncome In xmlIncomeList
        dblIncome = CalculateAverageNetProfit(xmlIncome)
        dblTotalAnnualIncome = dblTotalAnnualIncome + dblIncome
    Next
    
    strIncome = Format$(dblTotalAnnualIncome, "########0.00")
    
    If m_blnFixedLengthRecords Then
'        strApplicantBlock = strApplicantBlock & CStr(dblTotalAnnualIncome) & MakeString(15 - Len(dblIncome), cstrLeadingCharForNumbers)
        strApplicantBlock = strApplicantBlock & strIncome & MakeString(15 - Len(strIncome), cstrLeadingCharForNumbers)
    Else
'        strApplicantBlock = strApplicantBlock & CStr(dblTotalAnnualIncome) & m_strFieldSeparator
        strApplicantBlock = strApplicantBlock & strIncome & m_strFieldSeparator
    End If
    'BMIDS00285 MDC 27/09/2002 - End
    
    'Person Spare 3
    If m_blnFixedLengthRecords Then
        strApplicantBlock = strApplicantBlock & String(10, cstrLeadingCharForNumbers)
    Else
        strApplicantBlock = strApplicantBlock & m_strFieldSeparator
    End If
    
    'Add/Update
    strApplicantBlock = strApplicantBlock & "A"
    
    'BMIDS00025 MDC 24/05/2002 - End
    
    'Years at Address
    strTemp = Trim(Str(intNumberOfYears))
    
'    'BMIDS00025 MDC 23/05/2002
'    'strApplicantBlock = strApplicantBlock & String(4 - Len(strTemp), cstrLeadingCharForNumbers) & strTemp
'    strApplicantBlock = strApplicantBlock & MakeString(4 - Len(strTemp), cstrLeadingCharForNumbers) & strTemp
'    'BMIDS00025 MDC 23/05/2002 - End
'
'    'Income **************
'    strApplicantBlock = strApplicantBlock & String(7, cstrLeadingCharForNumbers)
'
'    'Purchase Price *******
'    strApplicantBlock = strApplicantBlock & strPurchasePrice
'
'    'Filler
'    strApplicantBlock = strApplicantBlock & String(11, cstrFiller)
'
'    'Add/Update
'    strApplicantBlock = strApplicantBlock & "A"
    
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
        
    'BMIDS00025 MDC 05/06/2002
'    strNewRecordHeader = Left(strRecordHeader, 9) & "MORTGAGE" & Mid(strRecordHeader, 18)
'    strNewRecordHeader = Left(strNewRecordHeader, Len(strNewRecordHeader) - 8) & strProductType
    'BMIDS00025 MDC 05/06/2002 - End
    
    strAllDataBlocks = strAllDataBlocks & strNewRecordHeader & strApplicantBlock & cstrRecordSeparator
    
BuildApplicantExit:
    
    Set xmlTable = Nothing
    Set xmlFormat = Nothing
    'BMIDS00025 MDC 07/06/2002
    Set xmlIncome = Nothing
    Set xmlIncomeList = Nothing
    'BMIDS00025 MDC 07/06/2002 - End
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

'BMIDS00285 MDC 27/09/2002 - Include Self Employed Income
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

Private Function GetFieldValueString(ByVal vxmlDataElement As IXMLDOMElement, _
                                     ByVal vxmlFormatElement As IXMLDOMElement, _
                                     ByVal strFieldName As String, _
                                     Optional ByVal blnFormatted As Boolean = True, _
                                     Optional blnLastField As String = False, _
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
    Dim blnContinue As Boolean
    Dim strRet As String    'BMIDS00025 MDC 24/05/2002
    
    strSeparator = IIf(blnLastField, cstrRecordSeparator, m_strFieldSeparator)
    
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
        
    ' Return the value, if formatting is not required
    If Not blnFormatted Then
        GetFieldValueString = strFieldValue
        Exit Function
    End If
    
    'If fixed value is mentioned, just return it
    If strFixedValue <> "" Then
        GetFieldValueString = strFixedValue & strSeparator
        Exit Function
    End If
    
    'If Empty return the string containing spaces or zeros
    'BMIDS00025 MDC 24/05/2002 - Only pad if fixed length format
    If UCase(strIsEmpty) = "Y" And m_blnFixedLengthRecords Then
        If strDataType = "STRING" Or strDataType = "COMBO" Or strDataType = "DATE" Then
            GetFieldValueString = Space(intFieldSize) & strSeparator
        Else
            GetFieldValueString = String(intFieldSize, "0") & strSeparator
        End If
        
        Exit Function
    End If

    ' Format the value, if required
    'BMIDS00025 MDC 24/05/2002
    strRet = GetFormattedValue(strFieldValue, intFieldSize, strDataType)
    If blnAddFieldSeparator Then
        strRet = strRet & strSeparator
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
                
                'BMIDS00025 MDC 23/05/2002
                If intValueLength >= intFieldSize Then
                    GetFormattedValue = Left(strFieldValue, intFieldSize)
                Else
                    'Only pad field if using fixed length format
                    If m_blnFixedLengthRecords Then
                        'GetFormattedValue = strFieldValue & String(intFieldSize - intValueLength, cstrTrailingCharForStrings)
                        GetFormattedValue = strFieldValue & MakeString(intFieldSize - intValueLength, cstrTrailingCharForStrings)
                    Else
                        GetFormattedValue = strFieldValue
                    End If
                End If
                'BMIDS00025 MDC 23/05/2002 - End
            
            Case "NUMBER"
                If intValueLength < intFieldSize Then  'actual length is less than the expected in record
                    'BMIDS00025 MDC 23/05/2002
                    'Only pad field if using fixed length format
                    If m_blnFixedLengthRecords Then
                        'Pad with leading zeros
                        'GetFormattedValue = String(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                        GetFormattedValue = MakeString(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                    Else
                        GetFormattedValue = strFieldValue
                    End If
                    'BMIDS00025 MDC 23/05/2002 - End
                ElseIf intValueLength > intFieldSize Then 'actual length is more than the expected in record
                    'Trim right side
                    GetFormattedValue = Left(strFieldValue, intFieldSize)
                Else  'actual length is same as expected
                    GetFormattedValue = strFieldValue
                End If
                
            Case "DATE"
                GetFormattedValue = Left(strFieldValue, 2) & Mid(strFieldValue, 4, 2) & Mid(strFieldValue, 7, 4)

            Case Else
                errThrowError "omHIBO.GetFormattedValue", oeInValidDataTypeValue
        End Select
    End If
    
End Function

Private Function GuarantorDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
    
On Error GoTo GuarantorDataBlockExit

    Const cstrFunctionName As String = "GuarantorDataBlock"

    Dim xmlGuarantorList As IXMLDOMNodeList, xmlGuarantor As IXMLDOMNode, xmlNode As IXMLDOMNode
    Dim xmlTable As IXMLDOMNode, xmlFormat As IXMLDOMNode
    
    Dim strGuarantorBlock As String, intCustomerOrder As Integer, strCustomerOrder As String
    Dim strTemp As String
    
    Set xmlGuarantorList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='2']")
    
    ' Find the first guarantor
    intCustomerOrder = 999  'Initialise Customer Order with high value
    For Each xmlNode In xmlGuarantorList
        strCustomerOrder = xmlGetAttributeText(xmlNode, "CUSTOMERORDER")
        If strCustomerOrder <> "" Then
            'BMIDS00025 MDC 05/06/2002 - Correction to search
            If CSafeInt(strCustomerOrder) < intCustomerOrder Then
            'If Val(strCustomerOrder) < intCustomerOrder Then
                Set xmlGuarantor = xmlNode
                intCustomerOrder = CSafeInt(strCustomerOrder)
            End If
            'BMIDS00025 MDC 05/06/2002 - End
        End If
    Next xmlNode

    'If guarantor record is found, build the data block
    If Not xmlGuarantor Is Nothing Then
        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMER")
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//GUARANTOR")
        
        'BMIDS00025 MDC 05/06/2002
        'strGuarantorBlock = "GA"
        strGuarantorBlock = "GA" & m_strFieldSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "TITLE")   'Title

        'BMIDS00025 MDC 05/06/2002
        'Initials
'        strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME") & _
'                  GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME")
        strTemp = GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", , , False) & _
                  GetFieldValueString(xmlTable, xmlFormat, "SECONDFORENAME")
        strTemp = Trim(strTemp)
        'BMIDS00025 MDC 05/06/2002 - End

        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & strTemp & MakeString(3 - Len(strTemp), cstrTrailingCharForStrings)
        Else
            strGuarantorBlock = strGuarantorBlock & strTemp
        End If
'        If Len(strTemp) > 2 Then
'            strGuarantorBlock = strGuarantorBlock & Left(strTemp, 2)
'        Else
'            strGuarantorBlock = strGuarantorBlock & strTemp & String(2 - Len(strTemp), cstrTrailingCharForStrings)
'        End If

        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "SURNAME")  'Surname
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "DATEOFBIRTH") 'Date of Birth
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "GENDER")   'Sex
'        If m_blnFixedLengthRecords Then
'            strGuarantorBlock = strGuarantorBlock & MakeString(5, cstrTrailingCharForStrings)   'Title
'            strGuarantorBlock = strGuarantorBlock & MakeString(3, cstrTrailingCharForStrings)   'Initials
'            strGuarantorBlock = strGuarantorBlock & MakeString(20, cstrTrailingCharForStrings)  'Surname
'            strGuarantorBlock = strGuarantorBlock & MakeString(10, cstrTrailingCharForStrings)  'Date of Birth
'            strGuarantorBlock = strGuarantorBlock & MakeString(1, cstrTrailingCharForStrings)   'Gender
'        Else
'            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
'        End If
        'BMIDS00025 MDC 05/06/2002 - End
        
        
        'BMIDS00025 MDC 05/06/2002
        'Filler
        'strGuarantorBlock = strGuarantorBlock & String(1, cstrFiller)
        
        'BM0214
        'Company Name
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & Space(30)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If
        
        'strGuarantorBlock = strGuarantorBlock & _
        '        Trim$(Left$(GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", False, , False) & " " & _
        '        GetFieldValueString(xmlTable, xmlFormat, "SURNAME", False), 30)) & m_strFieldSeparator
        'BM0214 End
        
'        strTemp = Left(GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME", False), 1) & _
                  GetFieldValueString(xmlTable, xmlFormat, "SURNAME", False)
        
        'ADDRESS
        Set xmlTable = Nothing: Set xmlFormat = Nothing
'        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMER/CUSTOMERADDRESS")
        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMERADDRESS[@ADDRESSTYPE='1']")
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//GUARANTOR/ADDRESS")
        
        ' Address Line 1
        'strGuarantorBlock = strGuarantorBlock & Trim(strTemp) & String(32 - Len(strTemp), cstrTrailingCharForStrings)
        strGuarantorBlock = strGuarantorBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlTable, xmlFormat, "FLATNUMBER", , , False)) & " " & _
                Trim$(GetFieldValueString(xmlTable, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35) & m_strFieldSeparator
                    
        'BMIDS00025 MDC 23/05/2002 - End
        
        ' Address Line 2
'        Set xmlTable = Nothing: Set xmlFormat = Nothing
'        Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMERADDRESS[@ADDRESSTYPE='1']")
'        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//GUARANTOR/ADDRESS")
        strGuarantorBlock = strGuarantorBlock & _
                Left$(Trim$(Trim(GetFieldValueString(xmlTable, xmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & _
                " " & Trim(GetFieldValueString(xmlTable, xmlFormat, "STREET", , , False))), 35) & m_strFieldSeparator
        
        'Address Line 3, 4, 5, 6
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "DISTRICT")
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "TOWN")
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "COUNTY")
        'BMIDS00025 MDC 23/05/2002
        'strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "COUNTRY")
        'BMIDS00025 MDC 23/05/2002 - End
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "POSTCODE")
        
        'Telephone
        'Set xmlTable = Nothing: Set xmlFormat = Nothing
        'Set xmlTable = xmlGuarantor.selectSingleNode(".//CUSTOMER/CUSTOMERTELEPHONENUMBER")
        'Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//GUARANTOR/CUSTOMERTELEPHONENUMBER")
                
        'JR - Omiplus24
        'BMIDS00025 MDC 05/06/2002
'        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "AREACODE")
'        strGuarantorBlock = strGuarantorBlock & Space(1) & GetFieldValueString(xmlTable, xmlFormat, "TELEPHONENUMBER")
        
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & MakeString(15, cstrTrailingCharForStrings)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If
        
        'BM0214
        'Valuer Code
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & Space(8)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If
        
        'Person Spare 1 - Forename
        Set xmlTable = xmlGetNode(xmlGuarantor, ".//CUSTOMER")
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//FIRSTNAME")
        strGuarantorBlock = strGuarantorBlock & GetFieldValueString(xmlTable, xmlFormat, "FIRSTFORENAME")
        
        'Person Spare 2 - Income
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & String(15, cstrLeadingCharForNumbers)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If
        
        'Person Spare 3
        If m_blnFixedLengthRecords Then
            strGuarantorBlock = strGuarantorBlock & String(10, cstrLeadingCharForNumbers)
        Else
            strGuarantorBlock = strGuarantorBlock & m_strFieldSeparator
        End If
        
        'BM0214 End
        
        'Valuer Code, Years at Address, Income, Purchase Price are not applicable
        'BMIDS00025 MDC 05/06/2002 - Telephone No also not required
'        If m_blnFixedLengthRecords Then
'            strGuarantorBlock = strGuarantorBlock & MakeString(15, cstrTrailingCharForStrings)
'            strGuarantorBlock = strGuarantorBlock & String(6, cstrTrailingCharForStrings)
'            strGuarantorBlock = strGuarantorBlock & String(4, cstrLeadingCharForNumbers)
'            strGuarantorBlock = strGuarantorBlock & String(7, cstrLeadingCharForNumbers)
'            strGuarantorBlock = strGuarantorBlock & String(15, cstrLeadingCharForNumbers)
'        Else
'            strGuarantorBlock = strGuarantorBlock & String$(5, m_strFieldSeparator)
'        End If
        
        'Filler
        'strGuarantorBlock = strGuarantorBlock & String(11, cstrFiller)
        'BMIDS00025 MDC 05/06/2002 - End
        
        'Add/Update
        strGuarantorBlock = strGuarantorBlock & "A"
        
        ' Add to the string that is writte to the txt file
        'BMIDS00025 MDC 05/06/2002
        'strAllDataBlocks = strAllDataBlocks & strRecordHeader & strGuarantorBlock & cstrRecordSeparator
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
                                    m_strHeaderSuffix & strGuarantorBlock & cstrRecordSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
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
    
    Dim xmlNewProperty As IXMLDOMElement, xmlFormat As IXMLDOMElement
    Dim strSecurityAddrBlock As String, strTemp As String
    'BMIDS00260 MDC 29/07/2002
    Dim strValueId As String
    'BMIDS00260 MDC 29/07/2002 - End
    
    'BMIDS00260 MDC 29/07/2002
    'BS BM0526 16/04/03
    'If m_blnNewAdvance Then
    If m_blnNewAdvance Or m_blnBuyToLet Then
        Set xmlNewProperty = vxmlResponse.selectSingleNode(".//NEWPROPERTY")
    Else
        'Get valueid for Current/Home Address
        strValueId = GetFirstComboValueId("CustomerAddressType", "H")
        Set xmlNewProperty = vxmlResponse.selectSingleNode(".//CUSTOMER/CUSTOMERADDRESS[@ADDRESSTYPE='" & strValueId & "']")
    End If
    'BMIDS00260 MDC 29/07/2002 - End
    
    If Not xmlNewProperty Is Nothing Then
        
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//NEWPROPERTY")
        
        'BMIDS00025 MDC 05/06/2002
        'strSecurityAddrBlock = "SA"
        strSecurityAddrBlock = "SA" & m_strFieldSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
        ' Title, Initials, Surname, DateofBirth, Sex - not applicable to NewProperty..
        'BMIDS00025 MDC 05/06/2002 - Cater for variable length records too
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
        
        'Filler
        'strSecurityAddrBlock = strSecurityAddrBlock & String(1, cstrFiller)
        
        'Company Name
        'strSecurityAddrBlock = strSecurityAddrBlock & String(30, cstrTrailingCharForStrings)
        'BMIDS00025 MDC 05/06/2002 - End
        
        'Address
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//NEWPROPERTY/ADDRESS")
                
        'Address Line 1
        'BMIDS00025 MDC 23/05/2002
'        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENAME")
        strSecurityAddrBlock = strSecurityAddrBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlNewProperty, xmlFormat, "FLATNUMBER", , , False)) & " " & _
                Trim$(GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35) & m_strFieldSeparator
        
        'Address Line 2
'        strTemp = Trim(GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENUMBER")) & Space(1) & _
                  Trim(GetFieldValueString(xmlNewProperty, xmlFormat, "STREET"))
        strSecurityAddrBlock = strSecurityAddrBlock & _
                Left$(Trim$(Trim(GetFieldValueString(xmlNewProperty, xmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & _
                " " & Trim(GetFieldValueString(xmlNewProperty, xmlFormat, "STREET", , , False))), 35) & m_strFieldSeparator
        
        'strSecurityAddrBlock = strSecurityAddrBlock & strTemp & MakeString(32 - Len(strTemp), cstrTrailingCharForStrings)
        'BMIDS00025 MDC 23/05/2002 - End
        
        'Address Line 3, 4, 5, 6
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "DISTRICT")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "TOWN")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "COUNTY")
        strSecurityAddrBlock = strSecurityAddrBlock & GetFieldValueString(xmlNewProperty, xmlFormat, "POSTCODE")
        
        'BMIDS00025 MDC 05/06/2002
        If m_blnFixedLengthRecords Then
            ' Telephone Number
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrTrailingCharForStrings)
                    
            'Valuer Code ******
            strSecurityAddrBlock = strSecurityAddrBlock & String(8, cstrTrailingCharForStrings)
            
            'Years at address, Income are zeros
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrLeadingCharForNumbers)
            strSecurityAddrBlock = strSecurityAddrBlock & String(15, cstrLeadingCharForNumbers)
            
            'Purchase Price ******
            'strSecurityAddrBlock = strSecurityAddrBlock & strPurchasePrice
            strSecurityAddrBlock = strSecurityAddrBlock & String$(10, cstrLeadingCharForNumbers)
        Else
            strSecurityAddrBlock = strSecurityAddrBlock & MakeString(5, m_strFieldSeparator)
        End If
        'BMIDS00025 MDC 05/06/2002 - End
        
        'BMIDS00025 MDC 05/06/2002
        'strSecurityAddrBlock = strSecurityAddrBlock & String(11, cstrFiller) 'Filler
        'BMIDS00025 MDC 05/06/2002 - End
        
        strSecurityAddrBlock = strSecurityAddrBlock & "A" ' Add/Update
        
        ' Add this block to the string that is written to txt file
        'BMIDS00025 MDC 05/06/2002
        'strAllDataBlocks = strAllDataBlocks & strRecordHeader & strSecurityAddrBlock & cstrRecordSeparator
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strSecurityAddrBlock & cstrRecordSeparator
        'BMIDS00025 MDC 05/06/2002
        
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
    
    Dim xmlSolicitor As IXMLDOMNode, xmlFormat As IXMLDOMNode
    Dim strSolicitorBlock As String, strTemp As String
    Dim xmlAddressOrTelephone As IXMLDOMNode    'BM0127 MDC 03/12/2002
    
    'BM0127 MDC 03/12/2002
    'Set xmlSolicitor = vxmlResponse.selectSingleNode(".//LEGALREP")
    Set xmlSolicitor = vxmlResponse.selectSingleNode(".//LEGALREP/NAMEANDADDRESSDIRECTORY")
    If xmlSolicitor Is Nothing Then
        Set xmlSolicitor = vxmlResponse.selectSingleNode(".//LEGALREP/THIRDPARTY")
    End If
    'BM0127 MDC 03/12/2002 - End
    
    If Not xmlSolicitor Is Nothing Then
        'BMIDS00025 MDC 05/06/2002
        'strSolicitorBlock = "ST"
        strSolicitorBlock = "ST" & m_strFieldSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP")
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
        'BMIDS00025 MDC 05/06/2002
        If m_blnFixedLengthRecords Then
            strSolicitorBlock = strSolicitorBlock & String(4, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(2, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(20, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(8, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(1, cstrTrailingCharForStrings)
        Else
            strSolicitorBlock = strSolicitorBlock & MakeString(5, m_strFieldSeparator)
        End If
        'BMIDS00025 MDC 05/06/2002 - End
        
        ' Filler
        'BMIDS00025 MDC 05/06/2002
'        strSolicitorBlock = strSolicitorBlock & String(1, cstrFiller)
        'BMIDS00025 MDC 05/06/2002 - End
        
        'Company Name
        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "COMPANYNAME")
                
        'Address
        Set xmlFormat = Nothing
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP/ADDRESS")
        
        'BM0127 MDC 03/12/2002
        Set xmlAddressOrTelephone = xmlGetNode(xmlSolicitor, "ADDRESS")
        'GD BM0346 - added optional parameter
        strSolicitorBlock = strSolicitorBlock & BuildAddress(xmlAddressOrTelephone, xmlFormat, True)

        'BMIDS00025 MDC 05/06/2002
'        'Address Line 1
'        'strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "BUILDINGORHOUSENAME")
'        strSolicitorBlock = strSolicitorBlock & _
'                Left$(Trim$(Trim$(GetFieldValueString(xmlSolicitor, xmlFormat, "FLATNUMBER", , , False)) & " " & _
'                Trim$(GetFieldValueString(xmlSolicitor, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35) & m_strFieldSeparator
'
'        'Address Line 2
''        strTemp = Trim(GetFieldValueString(xmlSolicitor, xmlFormat, "BUILDINGORHOUSENUMBER")) & Space(1) & _
''                  Trim(GetFieldValueString(xmlSolicitor, xmlFormat, "STREET"))
''        strSolicitorBlock = strSolicitorBlock & strTemp & String(32 - Len(strTemp), cstrTrailingCharForStrings)
'        strSolicitorBlock = strSolicitorBlock & _
'                Left$(Trim$(Trim(GetFieldValueString(xmlSolicitor, xmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & _
'                " " & Trim(GetFieldValueString(xmlSolicitor, xmlFormat, "STREET", , , False))), 35) & m_strFieldSeparator
'        'BMIDS00025 MDC 23/05/2002 - End
'
'        'Address Line 3, 4, 5, 6
'        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "DISTRICT")
'        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "TOWN")
'        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "COUNTY")
'        strSolicitorBlock = strSolicitorBlock & GetFieldValueString(xmlSolicitor, xmlFormat, "POSTCODE")
        'BM0127 MDC 03/12/2002 - End
        
        'Telephone Number
        Set xmlFormat = Nothing
        Set xmlFormat = vxmlFormatTemplate.selectSingleNode(".//APPLICATIONLEGALREP/CONTACTDETAILS")
        
        'BM0127 MDC 03/12/2002
        strTemp = GetFirstComboValueId("ContactTelephoneUsage", "W")
        Set xmlAddressOrTelephone = xmlGetNode(xmlSolicitor, "CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE='" & strTemp & "']")
        
        strSolicitorBlock = strSolicitorBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlAddressOrTelephone, xmlFormat, "AREACODE", , , False)) & " " & _
                Trim$(GetFieldValueString(xmlAddressOrTelephone, xmlFormat, "TELENUMBER", , , False))), 15) & m_strFieldSeparator

        'BM0127 MDC 03/12/2002 - End
                
        'Valuer code, Years at address, income, purchase price are not applicable.
        'BMIDS00025 MDC 23/05/2002
        If m_blnFixedLengthRecords Then
            strSolicitorBlock = strSolicitorBlock & String(8, cstrTrailingCharForStrings)
            strSolicitorBlock = strSolicitorBlock & String(15, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(15, cstrLeadingCharForNumbers)
            strSolicitorBlock = strSolicitorBlock & String(10, cstrLeadingCharForNumbers)
        Else
            strSolicitorBlock = strSolicitorBlock & MakeString(4, m_strFieldSeparator)
        End If
        
        'Filler
'        strSolicitorBlock = strSolicitorBlock & String(11, cstrFiller)
        'BMIDS00025 MDC 23/05/2002 - End
        
        'Add/Update
        strSolicitorBlock = strSolicitorBlock & "A"
            
        ' Add this block to the string that is written to txt file
        'BMIDS00025 MDC 05/06/2002
'        strAllDataBlocks = strAllDataBlocks & strRecordHeader & strSolicitorBlock & cstrRecordSeparator
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
                            m_strHeaderSuffix & strSolicitorBlock & cstrRecordSeparator
        'BMIDS00025 MDC 05/06/2002 - End
    End If
                                    
    SolicitorDataBlock = True
                    
SolicitorDataBlockExit:
    
    Set xmlSolicitor = Nothing
    Set xmlFormat = Nothing
    Set xmlAddressOrTelephone = Nothing 'BM0127 MDC 03/12/2002
    
    errCheckError cstrFunctionName, TypeName(Me)
                    
End Function

Private Function EmployerDataBlock(ByVal vxmlResponse As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
'-------------------------------------------------------------------------------
' For each applicant, build Employment records using the latest employment data
'--------------------------------------------------------------------------------
        
On Error GoTo EmployerDataBlockExit
    
    Const cstrFunctionName As String = "EmployerDataBlock"
    
    Dim xmlApplicantList As IXMLDOMNodeList, xmlApplicant As IXMLDOMNode
    Dim xmlEmployment As IXMLDOMNode, xmlEmployer As IXMLDOMNode
    Dim xmlFormat As IXMLDOMNode
'    Dim xmlSelfEmployedDetail As IXMLDOMNode
    'BMIDS00260 MDC 30/07/2002
    Dim xmlEmploymentList As IXMLDOMNodeList
    'BMIDS00260 MDC 30/07/2002 - End
    
    Dim intNoOfApplicants As Integer, intCounter As Integer, intValue As Integer
    Dim intCustomerOrder As Integer
    Dim strEmployerBlock As String, strTemp As String
    
    
    'BMIDS00260 MDC 30/07/2002
    'Find all current employments
    Set xmlEmploymentList = vxmlResponse.selectNodes(".//CUSTOMER/INCOME/EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]")
    For Each xmlEmployment In xmlEmploymentList
        
        'BM0104 MDC 29/11/2002
        Set xmlEmployer = xmlGetNode(xmlEmployment, ".//EMPLOYER")
        If Not xmlEmployer Is Nothing Then
        'BM0104 MDC 29/11/2002 - End
            intCounter = intCounter + 1
            If intCounter <= 4 Then
                strEmployerBlock = "E" & Left(Trim(Str(intCounter)), 1)
                
                ' Title, Initials, Surname, DateOfBirth, Sex are not applicable
                If m_blnFixedLengthRecords Then
                    strEmployerBlock = strEmployerBlock & String(5, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(3, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(20, cstrTrailingCharForStrings)
                    strEmployerBlock = strEmployerBlock & String(10, cstrLeadingCharForNumbers)
                    strEmployerBlock = strEmployerBlock & String(1, cstrTrailingCharForStrings)
                Else
                    strEmployerBlock = strEmployerBlock & MakeString(6, m_strFieldSeparator)
                End If
                
                'Company Name
                Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER")
                Set xmlEmployer = xmlGetMandatoryNode(xmlEmployment, ".//EMPLOYER") 'BMIDS00279 MDC 01/08/2002
                
                strEmployerBlock = strEmployerBlock & GetFieldValueString(xmlEmployer, xmlFormat, "COMPANYNAME")
                
                'BMIDS00279 MDC 01/08/2002
                'Set xmlEmployer = xmlGetMandatoryNode(xmlEmployment, ".//EMPLOYER")
                'BMIDS00279 MDC 01/08/2002 - End
    
                'Address
                strEmployerBlock = strEmployerBlock & BuildAddress(xmlEmployer, xmlFormat)
                
                'Telephone Number
                Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/CONTACTDETAILS")
                'JR - Omiplus24
                intValue = xmlGetAttributeAsInteger(xmlEmployer, "USAGE")
                If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
                    'BMIDS00286 MDC 30/09/2002
                    strTemp = Left$(Trim$(Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "AREACODE", , , False)) _
                                & " " & Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "TELENUMBER", , , False))), 15)
                    strEmployerBlock = strEmployerBlock & strTemp
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
                    'BMIDS00286 MDC 30/09/2002 - End
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
                
                strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
                                        m_strHeaderSuffix & strEmployerBlock & cstrRecordSeparator
        
                'Determine if Self Employed
                intValue = CSafeInt(xmlGetMandatoryAttributeAsLong(xmlEmployment, "EMPLOYMENTSTATUS"))
                If IsValidationType("EmploymentStatus", intValue, "S") Then
                    'Self Employed
                    If Not AccountantDataBlock(xmlEmployment, vxmlFormatTemplate) Then
                        EmployerDataBlock = False
                        Exit For
                    End If
                End If
            End If
        End If  'BM0104 MDC 29/11/2002
    Next
    'BMIDS00260 MDC 30/07/2002 - End
    
'    Set xmlApplicantList = vxmlResponse.selectNodes(".//APPLICATIONCUSTOMERROLE[@CUSTOMERROLETYPE='1']")
'
'    'BMIDS00025 MDC 06/06/2002 - Limit to maximum of 4 records
'    'intNoOfApplicants = xmlApplicantList.length
'    intNoOfApplicants = IIf(xmlApplicantList.length > 4, 4, xmlApplicantList.length)
'    'BMIDS00025 MDC 06/06/2002 - End
'
'    For intCounter = 0 To intNoOfApplicants - 1
'        Set xmlApplicant = xmlApplicantList.Item(intCounter)
'        'BMIDS00025 MDC 06/06/2002 - Get MAIN employment
'        'Set xmlEmployment = xmlApplicant.selectSingleNode(".//CUSTOMER/INCOME/EMPLOYMENT")
'        Set xmlEmployment = xmlApplicant.selectSingleNode(".//CUSTOMER/INCOME/EMPLOYMENT[@MAINSTATUS='1']")
'        'BMIDS00025 MDC 06/06/2002 - End
'
'        intCustomerOrder = xmlGetMandatoryAttributeText(xmlApplicant, "CUSTOMERORDER")
'
'        If Not xmlEmployment Is Nothing Then
'            Set xmlEmployer = xmlGetMandatoryNode(xmlEmployment, ".//EMPLOYER")
'            'BMIDS00025 MDC 06/06/2002
''            Set xmlSelfEmployedDetail = xmlEmployment.selectSingleNode(".//SELFEMPLOYEDDETAIL")
'
'            strEmployerBlock = "E" & Left(Trim(Str(intCounter + 1)), 1)
'
'            ' Title, Initials, Surname, DateOfBirth, Sex are not applicable
'            If m_blnFixedLengthRecords Then
'                strEmployerBlock = strEmployerBlock & String(5, cstrTrailingCharForStrings)
'                strEmployerBlock = strEmployerBlock & String(3, cstrTrailingCharForStrings)
'                strEmployerBlock = strEmployerBlock & String(20, cstrTrailingCharForStrings)
'                strEmployerBlock = strEmployerBlock & String(10, cstrLeadingCharForNumbers)
'                strEmployerBlock = strEmployerBlock & String(1, cstrTrailingCharForStrings)
'            Else
'                strEmployerBlock = strEmployerBlock & MakeString(6, m_strFieldSeparator)
'            End If
'
'            'Company Name
'            Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER")
'            strEmployerBlock = strEmployerBlock & GetFieldValueString(xmlEmployer, xmlFormat, "COMPANYNAME")
'
'            'Address
'            strEmployerBlock = strEmployerBlock & BuildAddress(xmlEmployer, xmlFormat)
'
'            'Telephone Number
'            Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/CONTACTDETAILS")
'            'JR - Omiplus24
'            intValue = xmlGetAttributeAsInteger(xmlEmployer, "USAGE")
'            If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
'                strEmployerBlock = strEmployerBlock & _
'                            Left$(Trim$(Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "AREACODE", , , False)) _
'                            & " " & Trim$(GetFieldValueString(xmlEmployer, xmlFormat, "TELENUMBER", , , False))), 15)
'                If m_blnFixedLengthRecords Then
'                    strEmployerBlock = strEmployerBlock & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings)
'                Else
'                    strEmployerBlock = strEmployerBlock & m_strFieldSeparator
'                End If
'            End If
'
'            'Valuers code, Person Spare 1, Person Spare 2, Person Spare 3 not required
'            If m_blnFixedLengthRecords Then
'                strEmployerBlock = strEmployerBlock & String(8, cstrTrailingCharForStrings) 'Valuer Code
'                strEmployerBlock = strEmployerBlock & String(15, cstrLeadingCharForNumbers) 'Person Spare 1
'                strEmployerBlock = strEmployerBlock & String(15, cstrLeadingCharForNumbers) 'Person Spare 2
'                strEmployerBlock = strEmployerBlock & String(10, cstrLeadingCharForNumbers) 'Person Spare 3
'            Else
'                strEmployerBlock = strEmployerBlock & MakeString(4, m_strFieldSeparator)
'            End If
'
'            'Add/Update
'            strEmployerBlock = strEmployerBlock & "A"  'Add/Update
'
'            strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
'                                    m_strHeaderSuffix & strEmployerBlock & cstrRecordSeparator
'
'            'Determine if Self Employed
'            intValue = CSafeInt(xmlGetMandatoryAttributeAsLong(xmlEmployment, "EMPLOYMENTSTATUS"))
'            If IsValidationType("EmploymentStatus", intValue, "S") Then
'                'Self Employed
'                If Not AccountantDataBlock(xmlEmployment, vxmlFormatTemplate) Then
'                    EmployerDataBlock = False
'                    Exit For
'                End If
'            End If
'            'BMIDS00025 MDC 06/06/2002 - End
'        End If
'    Next
                                        
    EmployerDataBlock = True

EmployerDataBlockExit:
    
    Set xmlApplicantList = Nothing
    Set xmlApplicant = Nothing
    Set xmlEmployment = Nothing
    Set xmlEmployer = Nothing
    Set xmlFormat = Nothing
'    Set xmlSelfEmployedDetail = Nothing
    
    'BMIDS00260 MDC 30/07/2002 - End
    Set xmlEmploymentList = Nothing
    'BMIDS00260 MDC 30/07/2002 - End

    errCheckError cstrFunctionName, TypeName(Me)
                                        
End Function

Private Function AccountantDataBlock(ByVal xmlEmployment As IXMLDOMNode, _
                                    ByVal vxmlFormatTemplate As IXMLDOMNode) As Boolean
'-----------------------------------------------------------
' Build Accountant data block for self-employed applicant
'-----------------------------------------------------------
On Error GoTo AccountantDataBlockExit
    
    Const cstrFunctionName As String = "AccountantDataBlock"
    
    Dim xmlAccountant As IXMLDOMNode, xmlFormat As IXMLDOMNode
    Dim strAccountantBlock As String, strTemp As String
    Dim intValue As Integer 'JR - Omiplus24
    Static intCounter As Integer    'BMIDS00025 MDC 06/06/2002
    
    Set xmlAccountant = xmlEmployment.selectSingleNode(".//ACCOUNTANT")
    
    If Not xmlAccountant Is Nothing Then
        
        'BMIDS00025 MDC 06/06/2002
        'strAccountantBlock = "CT"
        intCounter = intCounter + 1
        strAccountantBlock = "C" & CStr(intCounter)
        'BMIDS00025 MDC 06/06/2002 - End
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable
        If m_blnFixedLengthRecords Then
            strAccountantBlock = strAccountantBlock & String(5, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(3, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(20, cstrTrailingCharForStrings)
            strAccountantBlock = strAccountantBlock & String(10, cstrLeadingCharForNumbers)
            strAccountantBlock = strAccountantBlock & String(1, cstrTrailingCharForStrings)
        Else
            strAccountantBlock = strAccountantBlock & MakeString(6, m_strFieldSeparator)
        End If
        
        'Filler
        'strAccountantBlock = strAccountantBlock & String(1, cstrFiller)

        'Accountant Data
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT")
        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "COMPANYNAME")
        
        ' Address
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT/ADDRESS")
        strAccountantBlock = strAccountantBlock & BuildAddress(xmlAccountant, xmlFormat)
        
'        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "BUILDINGORHOUSENAME")
'
'        ' House Number, street
'        strTemp = Trim(GetFieldValueString(xmlAccountant, xmlFormat, "BUILDINGORHOUSENUMBER")) & _
'                  Trim(GetFieldValueString(xmlAccountant, xmlFormat, "STREET"))
'
'        If Len(strTemp) > 32 Then
'            strTemp = Left(strTemp, 32)
'        Else
'            'BMIDS00025 MDC 23/05/2002
'            'strTemp = strTemp & String(32 - Len(strTemp), cstrTrailingCharForStrings)
'            strTemp = strTemp & MakeString(32 - Len(strTemp), cstrTrailingCharForStrings)
'            'BMIDS00025 MDC 23/05/2002 - End
'        End If
'
'        strAccountantBlock = strAccountantBlock & strTemp
'
'        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "DISTRICT")
'        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "TOWN")
'        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "COUNTY")
'        strAccountantBlock = strAccountantBlock & GetFieldValueString(xmlAccountant, xmlFormat, "POSTCODE")
        
        'Telephone Number
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//EMPLOYER/ACCOUNTANT/CONTACTDETAILS")
        'JR - Omiplus24
        intValue = xmlGetAttributeAsInteger(xmlAccountant, "USAGE")
        If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
            strAccountantBlock = strAccountantBlock & _
                        Left$(Trim$(Trim$(GetFieldValueString(xmlAccountant, xmlFormat, "AREACODE", , , False)) _
                        & " " & Trim$(GetFieldValueString(xmlAccountant, xmlFormat, "TELENUMBER", , , False))), 15)
        End If
        
        If m_blnFixedLengthRecords Then
            strAccountantBlock = strAccountantBlock & MakeString(15 - Len(strTemp), cstrTrailingCharForStrings) 'Fill out telephone number
            strAccountantBlock = strAccountantBlock & String(8, cstrTrailingCharForStrings) 'Valuer Code
            strAccountantBlock = strAccountantBlock & String(15, cstrLeadingCharForNumbers) 'Years at address
            strAccountantBlock = strAccountantBlock & String(15, cstrLeadingCharForNumbers) 'Income
            strAccountantBlock = strAccountantBlock & String(10, cstrLeadingCharForNumbers) 'Purchase Price
        Else
            strAccountantBlock = strAccountantBlock & MakeString(5, m_strFieldSeparator)
        End If
'        strAccountantBlock = strAccountantBlock & String(11, cstrFiller) 'Filler
        
        strAccountantBlock = strAccountantBlock & "A"  'Add/Update
        
        ' Add this block to the string that is written to txt file
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
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
    
    Dim xmlValuer As IXMLDOMNode, xmlFormat As IXMLDOMElement
    Dim strValuerBlock As String, strTemp As String
    Dim intValue As Integer 'JR - Omiplus24
    Dim strPanelNo As String    'BMIDS00025 MDC 05/06/2002
    
    ' Response will have valuerInstruction records sorted in descending order. Take the first one for our purpose
    Set xmlValuer = vxmlResponse.selectSingleNode(".//VALUERINSTRUCTION")
    
    If Not xmlValuer Is Nothing Then
        
        'BMIDS00025 MDC 05/06/2002
        'strValuerBlock = "VA"
        strValuerBlock = "VA" & m_strFieldSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
        'BMIDS00025 MDC 05/06/2002
        If m_blnFixedLengthRecords Then
            strValuerBlock = strValuerBlock & String(4, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(2, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(20, cstrTrailingCharForStrings)
            strValuerBlock = strValuerBlock & String(8, cstrLeadingCharForNumbers)
            strValuerBlock = strValuerBlock & String(1, cstrTrailingCharForStrings)
        Else
            strValuerBlock = strValuerBlock & MakeString(5, m_strFieldSeparator)
        End If
'        strValuerBlock = strValuerBlock & String(1, cstrFiller) 'Filler
        'BMIDS00025 MDC 05/06/2002 - End
                
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER")
        
        'BMIDS00025 MDC 05/06/2002
        strPanelNo = GetFieldValueString(xmlValuer, xmlFormat, "VALUERPANELNO")
        'BMIDS00025 MDC 05/06/2002 - End
        
        'Company Name
        Set xmlValuer = vxmlResponse.selectSingleNode(".//VALUERINSTRUCTION/VALUER")
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "COMPANYNAME")
                
        '------------------
        'Address
        '------------------
        
        'BMIDS00025 MDC 05/06/2002
        'Set xmlFormat = Nothing: Set xmlValuer = Nothing
        'Set xmlValuer = vxmlResponse.selectSingleNode(".//VALUERINSTRUCTION/VALUER")
        'BMIDS00025 MDC 05/06/2002 - End
        
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER/ADDRESS")
        
        'BMIDS00025 MDC 05/06/2002
        'Address Line 1
        'strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "BUILDINGORHOUSENAME")
        strValuerBlock = strValuerBlock & _
                Left$(Trim$(Trim$(GetFieldValueString(xmlValuer, xmlFormat, "FLATNUMBER", , , False)) & " " & _
                Trim$(GetFieldValueString(xmlValuer, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35) & m_strFieldSeparator
        
        'Address Line 2
        'strTemp = Trim(GetFieldValueString(xmlValuer, xmlFormat, "BUILDINGORHOUSENUMBER")) & Space(1) & _
                  Trim(GetFieldValueString(xmlValuer, xmlFormat, "STREET"))
        'strValuerBlock = strValuerBlock & strTemp & String(32 - Len(strTemp), cstrTrailingCharForStrings)
        strValuerBlock = strValuerBlock & _
                Left$(Trim$(Trim(GetFieldValueString(xmlValuer, xmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & _
                " " & Trim(GetFieldValueString(xmlValuer, xmlFormat, "STREET", , , False))), 35) & m_strFieldSeparator
        'BMIDS00025 MDC 23/05/2002 - End
        
        'Address Line 3, 4, 5, 6
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "DISTRICT")
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "TOWN")
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "COUNTY")
        strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "POSTCODE")
        
        'Telephone Number
        Set xmlFormat = Nothing
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//VALUER/CONTACTDETAILS")
        
        'JR - Omiplus24
        
        intValue = xmlGetAttributeAsInteger(xmlValuer, "USAGE")
        If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
        'BMIDS00025 MDC 05/06/2002
'            strValuerBlock = strValuerBlock & GetFieldValueString(xmlValuer, xmlFormat, "AREACODE")
'            strValuerBlock = strValuerBlock & cstrFiller & GetFieldValueString(xmlValuer, xmlFormat, "TELENUMBER")
            strValuerBlock = strValuerBlock & _
                    Left$(Trim$(Trim$(GetFieldValueString(xmlValuer, xmlFormat, "AREACODE", , , False)) & " " & _
                    Trim$(GetFieldValueString(xmlValuer, xmlFormat, "TELENUMBER", , , False))), 15) & m_strFieldSeparator
        Else
            If m_blnFixedLengthRecords Then
                strValuerBlock = strValuerBlock & String$(15, cstrTrailingCharForStrings)
            Else
                strValuerBlock = strValuerBlock & m_strFieldSeparator
            End If
        'BMIDS00025 MDC 05/06/2002 - End
        End If
        
        'BMIDS00025 MDC 05/06/2002
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
        'strValuerBlock = strValuerBlock & String(11, cstrFiller) 'Filler
        'BMIDS00025 MDC 05/06/2002 - End
        
        strValuerBlock = strValuerBlock & "A" 'Add/Update
        
        ' Add this block to the string that is written to txt file
        'BMIDS00025 MDC 05/06/2002
        'strAllDataBlocks = strAllDataBlocks & strRecordHeader & strValuerBlock & cstrRecordSeparator
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
                        m_strHeaderSuffix & strValuerBlock & cstrRecordSeparator
        'BMIDS00025 MDC 05/06/2002 - End
        
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
    
    Dim xmlIntermediaryOrg As IXMLDOMNode, xmlFormat As IXMLDOMNode
    Dim strIntermediaryOrg As String, strTemp As String
    Dim intValue As Integer 'JR - Omiplus24
    
    Set xmlIntermediaryOrg = vxmlResponse.selectSingleNode(".//INTERMEDIARY/INTERMEDIARYORGANISATION")
    
    If Not xmlIntermediaryOrg Is Nothing Then
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//INTERMEDIARY/INTERMEDIARYORGANISATION")
        
        strIntermediaryOrg = "IN"
        
        ' Title, Initials, Surname, DateOfBirth, Sex are not applicable. So empty
        strIntermediaryOrg = strIntermediaryOrg & String(4, cstrTrailingCharForStrings)
        strIntermediaryOrg = strIntermediaryOrg & String(2, cstrTrailingCharForStrings)
        strIntermediaryOrg = strIntermediaryOrg & String(20, cstrTrailingCharForStrings)
        strIntermediaryOrg = strIntermediaryOrg & String(8, cstrLeadingCharForNumbers)
        strIntermediaryOrg = strIntermediaryOrg & String(1, cstrTrailingCharForStrings)
                
        strIntermediaryOrg = strIntermediaryOrg & String(1, cstrFiller) 'Filler
        
        ' Company Name
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//INTERMEDIARY/INTERMEDIARYORGANISATION")
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "NAME")
    
        ' Address
        Set xmlFormat = Nothing
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//INTERMEDIARY/INTERMEDIARYORGANISATION/ADDRESS")
        
        'BMIDS956 include flatnumber
        strTemp = Left$(Trim$(Trim$(GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "FLATNUMBER", , , False)) & " " & _
            Trim$(GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "BUILDINGORHOUSENAME", , , False))), 35)
        If m_blnFixedLengthRecords Then
            Dim nLen As Integer
            nLen = Len(strTemp)
            strTemp = strTemp & MakeString(35 - nLen, cstrTrailingCharForStrings)
        Else
            strTemp = strTemp & m_strFieldSeparator
        End If
        'strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "BUILDINGORHOUSENAME")
        strIntermediaryOrg = strIntermediaryOrg & strTemp
        
        'Address Line 2
        strTemp = Trim(GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "BUILDINGORHOUSENUMBER")) & Space(1) & _
                  Trim(GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "STREET"))
        
        strIntermediaryOrg = strIntermediaryOrg & strTemp & String(32 - Len(strTemp), cstrTrailingCharForStrings)
        
        'Address Line 3, 4, 5, 6
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "DISTRICT")
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "TOWN")
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "COUNTY")
        strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "POSTCODE")

        'Contact Telephone Number
        Set xmlFormat = Nothing
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//INTERMEDIARY/INTERMEDIARYORGANISATION/CONTACTDETAILS")
        'JR - Omiplus24
        intValue = xmlGetAttributeAsInteger(xmlIntermediaryOrg, "USAGE")
        If IsValidationType(cstrContactPhoneGrpName, intValue, cstrContactPhoneValType) Then
            strIntermediaryOrg = strIntermediaryOrg & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "AREACODE")
            strIntermediaryOrg = strIntermediaryOrg & cstrFiller & GetFieldValueString(xmlIntermediaryOrg, xmlFormat, "TELENUMBER")
        End If
        
        'Valuer code, Years at address, income are not applicable.
        strIntermediaryOrg = strIntermediaryOrg & String(6, cstrTrailingCharForStrings)
        strIntermediaryOrg = strIntermediaryOrg & String(4, cstrLeadingCharForNumbers)
        strIntermediaryOrg = strIntermediaryOrg & String(7, cstrLeadingCharForNumbers)
        
        'Purchase Price
        strIntermediaryOrg = strIntermediaryOrg & "DIRECT BUSINESS"
        
        strIntermediaryOrg = strIntermediaryOrg & String(11, cstrFiller) 'Filler
        strIntermediaryOrg = strIntermediaryOrg & "A" 'Add/Update
        
        ' Add this block to the string that is written to txt file
        strAllDataBlocks = strAllDataBlocks & strRecordHeader & strIntermediaryOrg & cstrRecordSeparator
    
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
   
    Dim xmlNPVendor As IXMLDOMNode, xmlFormat As IXMLDOMNode
    Dim strNPVendor As String, strTemp As String
    Dim intValue As Integer 'JR - Omiplus24

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
        
        strNPVendor = strNPVendor & strTemp
        
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
                  
        strNPVendor = strNPVendor & strTemp
        
        'Data No 6
        strNPVendor = strNPVendor & "06"
        strNPVendor = strNPVendor & GetFieldValueString(xmlNPVendor, xmlFormat, "POSTCODE")
        
        'Data No 7 - Telephone Number
        Set xmlFormat = xmlGetMandatoryNode(vxmlFormatTemplate, ".//NEWPROPERTY/NEWPROPERTYVENDOR/CONTACTDETAILS")
        strNPVendor = strNPVendor & "07"
        'JR - Omiplus24
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

'BMIDS00025 MDC 06/06/2002
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
Dim strValue As String
Dim intIndex As Integer
'BM0283
Dim sBusinessSource As String
Dim bIsIndirect     As String
'BM0283 End

    'Get index of the lowest level Introducer
    intIndex = GetLowestLevelIntroducerIndexNumber(vxmlResponse)
    
    If intIndex > 0 Then
        
        Set xmlApplication = xmlGetMandatoryNode(vxmlResponse, ".//APPLICATION")
        
        'Call omAdmin to get Introducer data from BMids Introducer System
        Set objAdminBO = gobjContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        ' MO - 28/10/2002 - BMIDS00752 - Input the correct request - Start
        Set xmlTempRequest = xmlDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetIntroducerData"
        Set xmlTempNode = xmlDoc.createElement("INTRODUCER")
        xmlCopyMandatoryAttributeValue xmlApplication, xmlTempNode, "INTRODUCERIDLEVEL" & CStr(intIndex), "BMREFERENCE"
        xmlTempRequest.appendChild xmlTempNode
        ' MO - 28/10/2002 - BMIDS00752 - End
        
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
                strIntroducer = strIntroducer & MakeString(6, m_strFieldSeparator)
            End If
            
            ' Company Name
            strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "COMPANYNAME")
            
            'BM0283
            bIsIndirect = False
            sBusinessSource = xmlGetAttributeText(xmlGetNode(vxmlResponse, "APPLICATION"), "DIRECTINDIRECTBUSINESS")
            If Len(sBusinessSource) > 0 Then
                If comboAssistEx.IsValidationType("Direct/Indirect", CInt(sBusinessSource), "I") Then
                    bIsIndirect = True
                End If
            End If
            
            If bIsIndirect Then
            'BM0283 End
            
                ' MO - 28/10/2002 - BMIDS00752 - Build the correct address lines into the xmlIntroducer data - Start
                
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE1", xmlGetAttributeText(xmlIntroducer, "BUILDINGNAME")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE2", xmlGetAttributeText(xmlIntroducer, "BUILDINGNO") & " " & xmlGetMandatoryAttributeText(xmlIntroducer, "STREET")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE3", xmlGetAttributeText(xmlIntroducer, "DISTRICT")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE4", xmlGetAttributeText(xmlIntroducer, "TOWN")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE5", xmlGetAttributeText(xmlIntroducer, "COUNTY")
                xmlSetAttributeValue xmlIntroducer, "ADDRESSLINE6", xmlGetAttributeText(xmlIntroducer, "POSTCODE")
                
                ' MO - 28/10/2002 - BMIDS00752 - End
                
                ' Address
                
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE1")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE2")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE3")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE4")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE5")
                strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "ADDRESSLINE6")
                        
            'BM0283
            Else
                strIntroducer = strIntroducer & MakeString(6, m_strFieldSeparator)
            End If
            'BM0283 End
            
            'Contact Telephone Number
            strIntroducer = strIntroducer & GetFieldValueString(xmlIntroducer, xmlFormat, "TELEPHONENO")
            
            'Valuer code, Person Spare 1, Person Spare 2, Person Spare 3 are not applicable.
            If m_blnFixedLengthRecords Then
                strIntroducer = strIntroducer & String(8, cstrTrailingCharForStrings)   'Valuers Code
                strIntroducer = strIntroducer & String(15, cstrLeadingCharForNumbers)   'Person Spare 1
                strIntroducer = strIntroducer & String(15, cstrLeadingCharForNumbers)   'Person Spare 2
                strIntroducer = strIntroducer & String(10, cstrLeadingCharForNumbers)   'Person Spare 3
            Else
                strIntroducer = strIntroducer & MakeString(4, m_strFieldSeparator)
            End If
            
            strIntroducer = strIntroducer & "A" 'Add/Update
            
            ' Add this block to the string that is written to txt file
            strAllDataBlocks = strAllDataBlocks & strRecordHeader & m_strFieldSeparator & _
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

Const strFunctionName As String = "GetLowestLevelIntroducerIndexNumber"
    
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
'BMIDS00025 MDC 06/06/2002 - End

Public Function GenerateRecords(ByVal vstrXMLRequest As String) As String

On Error GoTo GenerateRecordsErr
    
    Const cstrFunctionName As String = "GenerateRecords"

    Dim vxmlResponse As IXMLDOMElement, xmlFormatTemplate As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40, xmlOutDoc As FreeThreadedDOMDocument40
    
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem

    Set vxmlResponse = GetData(vstrXMLRequest)
    
    'Load template for formatting
    Set xmlDoc = LoadFormatTemplate()
    Set xmlFormatTemplate = xmlDoc.documentElement
    
    'BMIDS00260 MDC 30/07/2002 - Parameterise
    If GetMandatoryGlobalParamBoolean("HunterFormat") Then
        m_blnFixedLengthRecords = False
        m_strFieldSeparator = "|"
    Else
        m_blnFixedLengthRecords = True
        m_strFieldSeparator = ""
    End If
'    'BMIDS00025 MDC 24/05/2002
'    If Len(Trim$(m_strFieldSeparator)) = 0 Then
'        m_blnFixedLengthRecords = True
'    End If
'    'BMIDS00025 MDC 24/05/2002 - End
    
    ' Generate data blocks
    If (RecordHeader(vxmlResponse, xmlFormatTemplate)) Then
        If ApplicantDataBlock(vxmlResponse, xmlFormatTemplate) Then
            If GuarantorDataBlock(vxmlResponse, xmlFormatTemplate) Then
                If SecurityAddressDataBlock(vxmlResponse, xmlFormatTemplate) Then
                    If SolicitorDataBlock(vxmlResponse, xmlFormatTemplate) Then
                        If ValuerDataBlock(vxmlResponse, xmlFormatTemplate) Then
                            If EmployerDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                'BMIDS00025 MDC 06/06/2002
                                If IntroducerDataBlock(vxmlResponse, xmlFormatTemplate) Then
                                    Call WriteToTextFile
                                End If
'                                If IntermediaryDataBlock(vxmlResponse, xmlFormatTemplate) Then
'                                    If ValuationArrangementDataBlock(vxmlResponse, xmlFormatTemplate) Then
'                                        ' Write the whole string on to text file
'                                        Call WriteToTextFile
'                                    End If
'                                End If
                                'BMIDS00025 MDC 06/06/2002 - End
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
    
    'BMIDS00260 29/07/2002
'    Dim objFSO As Scripting.FileSystemObject    'BMIDS00025 MDC 23/05/2002
    Dim blnAppendMode As Boolean
    'BMIDS00260 29/07/2002 - End
    
    Const cstrFunctionName As String = "WriteToTextFile"
    
    Dim intFileNo As Integer
    Dim strFileName As String, strFileLoc As String
    'BMIDS646
    Dim objCRM As omCRMWrapper
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Dim strResponse As String
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    'Remove last character (Record separator) of the full string- this might cause blank lines in the output
    Dim strTemp As String
    'BM0273 Remove trailing CRLF, if it exists
    'BMIDS646 Don't remove trailing CRLF for omCRM, we need them as no longer using print
    'which previously added them.
'    If Right(strAllDataBlocks, 2) = vbCrLf Then
'        strTemp = Left(strAllDataBlocks, Len(strAllDataBlocks) - 2)
'    Else
        strTemp = strAllDataBlocks
'        strTemp = strAllDataBlocks & Chr$(13) & Chr$(10)
'    End If
    'strTemp = Left(strAllDataBlocks, Len(strAllDataBlocks) - 1)
    'BM0273 End
    
    'BMIDS00260 MDC 29/07/2002
    blnAppendMode = GetMandatoryGlobalParamBoolean("HunterMode")
    If blnAppendMode Then
        'Append to a single output file
        strFileName = GetMandatoryGlobalParamString("HunterFile")
    Else
        'Write to a file for the application
        strFileName = strApplicationNumber & ".txt"
    End If
    'BMIDS00260 MDC 29/07/2002 - End
    
    'BMIDS00025 MDC 11/06/2002 - Parameterise output folder
    'strFileLoc = App.Path & "\HunterInterfaceResults"              ' Any parameterisation reqd ??
    strFileLoc = GetMandatoryGlobalParamString("HunterOutputFolder")
    'BMIDS00025 MDC 11/06/2002 - End
       
    'BMIDS00025 MDC 23/05/2002
    'Ensure that target folder exists
    'BMIDS00260 29/07/2002
'    Set objFSO = New Scripting.FileSystemObject
'    If Not objFSO.FolderExists(strFileLoc) Then
'        MkDir strFileLoc
'        'objFSO.CreateFolder strFileLoc
'    End If
'BMIDS646 omCRM will create The folder if necessary and write to file
'    CreateFolderEx strFileLoc
'    'BMIDS00025 MDC 23/05/2002 - End
'    'BMIDS00260 MDC 29/07/2002 - End
'
'    ' Open the required file for appending data
'    intFileNo = FreeFile
'
'    'BMIDS00260 MDC 29/07/2002
'    If blnAppendMode Then
'        Open strFileLoc & "\" & strFileName For Append As #intFileNo
'    Else
'        Open strFileLoc & "\" & strFileName For Output As #intFileNo
'    End If
'    'BMIDS00260 MDC 29/07/2002 - End
'
'    Print #intFileNo, strTemp
'    Close #intFileNo
    'BMIDS646
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
    strResponse = objCRM.AddText(xmlRequest.xml, strTemp)
                                                
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True

WriteToTextFileExit:

    'BMIDS00260 MDC 29/07/2002
    'Set objFSO = Nothing    'BMIDS00025 MDC 23/05/2002
    'BMIDS00260 MDC 29/07/2002 - End
    'BMIDS646
    Set objCRM = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlRequest = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function

'BMIDS00025 MDC 23/05/2002
Private Function MakeString(ByVal intLength As Integer, ByVal strChar As String) As String

    If intLength > 0 Then
        MakeString = String(intLength, strChar)
    Else
        MakeString = strChar
    End If
    
End Function

'BMIDS00025 MDC 23/05/2002 - End

'BMIDS00025 MDC 06/06/2002
'GD BM0346 added optional parameter
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

    'Validate parameters ----------------------------------------------------------------------------
    'GD BM0346
    If vxmlFormat Is Nothing Or (vxmlTable Is Nothing And (blnEmptyvxmlTableAllowed = False)) Then
        errThrowError strFunctionName, oeMissingElement, "Table or Format node not specified."
    End If
    
    'Build address block fragment -------------------------------------------------------------------
    ' Address Line 1
    strAddress = Left$(Trim$(Trim$(GetFieldValueString(vxmlTable, vxmlFormat, "FLATNUMBER", , , False)) & " " & _
            Trim$(GetFieldValueString(vxmlTable, vxmlFormat, "BUILDINGORHOUSENAME", , , False))), 35)
    If m_blnFixedLengthRecords Then
        strAddress = strAddress & MakeString(35 - Len(strTemp), cstrTrailingCharForStrings)
    Else
        strAddress = strAddress & m_strFieldSeparator
    End If
        
    'Address Line 2
    strTemp = Left$(Trim$(Trim$(GetFieldValueString(vxmlTable, vxmlFormat, "BUILDINGORHOUSENUMBER", , , False)) & " " & _
                Trim$(GetFieldValueString(vxmlTable, vxmlFormat, "STREET", , , False))), 35)
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
'BMIDS00025 MDC 06/06/2002 - End

'BMIDS00260 MDC 29/07/2002 - Create folder n layers deep
Private Sub CreateFolderEx(ByVal strPath)

On Error GoTo CreateFolderExExit

Dim arrFolders() As String
Dim intLevels As Integer
Dim strLowest As String
Dim strWorking As String
Const cstrFunctionName As String = "CreateFolderEx"

    If m_objFSO.FolderExists(strPath) = False Then
        'Replace strPath, "/", "\"
        arrFolders = Split(strPath, "\")
        intLevels = UBound(arrFolders)
        If intLevels > 0 Then
            strLowest = arrFolders(intLevels)
            ReDim Preserve arrFolders(intLevels - 1)
            strWorking = Join(arrFolders, "\")
            CreateFolderEx strWorking
            'On Error Resume Next
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
    'BMIDS00260 MDC 29/07/2002
    Set m_objFSO = New Scripting.FileSystemObject
    'BMIDS00260 MDC 29/07/2002 - End
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
    'BMIDS00260 MDC 29/07/2002
    Set m_objFSO = Nothing
    'BMIDS00260 MDC 29/07/2002 - End
End Sub

