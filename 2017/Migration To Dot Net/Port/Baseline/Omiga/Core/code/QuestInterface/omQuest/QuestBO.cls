VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuestBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        AQR             Description
'MV     16/03/2004  BBG107          WP019 - Created
'MV     20/04/2004  BBG240          Amended CreateDataBlock() to include Selling Agent Names
'MV     29/04/2004  BB240           Amended GetData() to include the template Attribute
'MSla   01/06/2004  BBG614          Correct instantiation Request Broker object
'MV     05/06/2004  BBG48           Amended WriteToTextFile() to include omCRM Wrapper
'MV     09/06/2004  BBG663          Amended CreateDataBlock()
'MV     09/06/2004  BBG48           Amended WriteToTextFile()
'TK     25/06/2004  BBG663          Amended CreateDataBlock()
'MV     30/07/2004  E2EM00000627    Amended CreateDataBlock()
'MV     10/09/2004  E2EM00001470    Amended CreateQuestReportData() , CreateDataBlock()
'MV     20/09/2004  E2EM00001676/E2EM00001655 Amended CreateDataBlock()
'------------------------------------------------------------------------------------------
Option Explicit


Implements ObjectControl
Private m_objContext As ObjectContext
Private strAllDataBlocks As String
Private strApplicationNumber  As String
'Private m_objFSO As Scripting.FileSystemObject
Private m_blnFixedLengthRecords As Boolean
Const cstrLeadingCharForNumbers = "0" 'vbNullChar '"0"
Const cstrTrailingCharForStrings = " " ' vbNullChar '" "
Const cstrRecordSeparator = vbCrLf
Private m_strFieldSeparator As String
Private strDatablock  As String
Public Function CreateDataBlock(ByVal xmlRequest As IXMLDOMNode, ByVal xmlCaseTaskNode As IXMLDOMNode) As Boolean
    
    On Error GoTo CreateDataBlockErr
        
    Const cstrFunctionName As String = "OmQuest.QuestBO.CreateDataBlock"
    
    Dim xmlTempNode As IXMLDOMNode
    
    Dim strFlatNo  As String
    Dim strHouseName  As String
    Dim strHouseNo As String
    Dim strStreetName As String
    Dim strDistrict As String
    Dim strTown As String
    Dim strCounty As String
    Dim strPostCode As String
    Dim strTenure As String
    Dim strYearBuilt As String
    Dim strPurchasePrice  As String
    Dim strAdvanceAmount As String
    Dim strApp1Title   As String
    Dim strApp1FirstForeName As String
    Dim strApp1SecondForeName As String
    Dim strApp1OtherForeNames As String
    Dim strApp1SurName As String
    Dim strApp2Title   As String
    Dim strApp2FirstForeName As String
    Dim strApp2SecondForeName As String
    Dim strApp2OtherForeNames As String
    Dim strApp2SurName As String
    Dim strValuationFee As String
    Dim strAppCorrespHouseNo  As String
    Dim strAppCorrespFlatNo  As String
    Dim strAppCorrespHouseName  As String
    Dim strAppCorrespPostCode  As String
    Dim strAppCorrespCounty  As String
    Dim strAppCorrespTown  As String
    Dim strAppCorrespDistrict As String
    Dim strAppCorrespStreetName  As String
    Dim strMortgagePurpose  As String
    Dim strPropertyInScotland  As String
    Dim strApp1HomeSTDCode  As String
    Dim strApp1HomeTeleNo As String
    Dim strApp1WorkSTDCode  As String
    Dim strApp1WorkTeleNo  As String
    Dim strArrgAddrssFlatNo  As String
    Dim strArrgAddrssHouseName As String
    Dim strArrgAddrssHouseNo  As String
    Dim strArrgAddrssStreetName  As String
    Dim strArrgAddrssDistrict  As String
    Dim strArrgAddrssTown  As String
    Dim strArrgAddrssCounty As String
    Dim strArrgAddrssPostCode As String
    Dim strLenderStaffContactTitle As String
    Dim strLenderStaffContactForeName As String
    Dim strLenderStaffContactSurName As String
    Dim strLenderStaffContactInitials As String
    Dim strLenderStaffContactTeleNo  As String
    Dim strLenderStaffContactSTDCode  As String
    Dim strMonthlyRentalIncome As String
    Dim strSellingAgentName  As String
    Dim strSellingAgentContactSTDCode As String
    Dim strSellingAgentContactTeleNo As String
    Dim strValuationType As String
    Dim strValuationTypeText As String
    Dim strDetSurveyQuoteInd As String
    Dim strApp2strSurname As String
    Dim strArrgAddrssCompanyName As String
    Dim strArrgAddrssAreaCode  As String
    Dim strLenderStaffContact  As String
    Dim strLenderContactTeleNo As String
    Dim strLenderContactAreaCode  As String
    Dim strUnitCode As String
    Dim strRemortgageInd As String
    Dim strSellingAgent  As String
    Dim strSpecialGroup As String
    Dim strArrgAddrssAreaTeleNo As String
    Dim strArrgContactTitleText As String
    Dim strArrgContactForeName  As String
    Dim strArrgContactSurname As String
    Dim strArrgContactName As String
    Dim strReinspectionInstruct As String
    Dim strCaseTaskId As String
    Dim strArrangementAccess  As String
    
    Dim blnSpecialGroup  As Boolean
    
    strReinspectionInstruct = GetGlobalParamString("ReinspectionInstruct")
    strCaseTaskId = xmlGetAttributeText(xmlCaseTaskNode, "TASKID")
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/NEWPROPERTY")
    
    If Not xmlTempNode Is Nothing Then
        
        'Flat Number & HouseName - Char(30)
        strFlatNo = xmlGetAttributeText(xmlTempNode, "FLATNUMBER")
        strHouseName = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENAME")
         
        'House Number - Char(5)
        strHouseNo = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENUMBER")
        
        'Street Name - Char((30)
        strStreetName = xmlGetAttributeText(xmlTempNode, "STREET")
        
        'District - Char(30)
        strDistrict = xmlGetAttributeText(xmlTempNode, "DISTRICT")
        
        'Town - Char(30)
        strTown = xmlGetAttributeText(xmlTempNode, "TOWN")
        
        'County - Char(30)
        strCounty = xmlGetAttributeText(xmlTempNode, "COUNTY")
        
        'PostCode - Char(9)
        strPostCode = xmlGetAttributeText(xmlTempNode, "POSTCODE")
        
        'Year Built - Char(4)
        strYearBuilt = xmlGetAttributeText(xmlTempNode, "YEARBUILT")
        
        'Tenure Type - Number(2)
        strTenure = xmlGetAttributeText(xmlTempNode, "TENURETYPE")
        
        'Property In Scotland
        strPropertyInScotland = xmlGetAttributeText(xmlTempNode, "PROPERTYLOCATION_TEXT")
        If UCase(strPropertyInScotland) = "SCOTLAND" Then
            strPropertyInScotland = "Y"
        Else
            strPropertyInScotland = "N"
        End If
        
        'MonthlyRentalIncome
        strMonthlyRentalIncome = xmlGetAttributeText(xmlTempNode, "MONTHLYRENTALINCOME")
        
        'Valuation Type
        strValuationType = xmlGetAttributeText(xmlTempNode, "VALUATIONTYPE")
        
        If Trim$(strReinspectionInstruct) = Trim$(strCaseTaskId) Then
            strValuationType = "40"
        End If
            
        'Valuationtype Text
        If strValuationType = "40" Then
            strValuationTypeText = GetComboText("ValuationType", strValuationType)
        Else
            strValuationTypeText = xmlGetAttributeText(xmlTempNode, "VALUATIONTYPE_TEXT")
        End If
        
        If UCase$(strValuationTypeText) = "HOMEBUYER" Then
            strDetSurveyQuoteInd = "Y"
        Else
            strDetSurveyQuoteInd = "N"
        End If
    
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION")
    
    If Not xmlTempNode Is Nothing Then
        'Purchase Price - Number(7)
        strPurchasePrice = xmlGetAttributeText(xmlTempNode, "PURCHASEPRICEORESTIMATEDVALUE")
    
        'MortgagePurpose
        strMortgagePurpose = xmlGetAttributeText(xmlTempNode, "TYPEOFAPPLICATION")
    
        'Special Group
        blnSpecialGroup = False
        
        strSpecialGroup = xmlGetAttributeText(xmlTempNode, "SPECIALGROUP")
        If strSpecialGroup <> "" Then
            If IsValidationType("SpecialGroup", CInt(strSpecialGroup), "BTL") Then
                blnSpecialGroup = True
            End If
        End If
        
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/QUOTATION/MORTGAGESUBQUOTE")
    
    'Advacne Amount - Number(7)
    strAdvanceAmount = xmlGetAttributeText(xmlTempNode, "AMOUNTREQUESTED")
    
    'Applicant 1 Details
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMER")
    If Not xmlTempNode Is Nothing Then
        
        'Applicant 1 Title - Char(8)
        strApp1Title = xmlGetAttributeText(xmlTempNode, "TITLE_TEXT")
        
        'Applicant 1 First ForeName - Char(1)
        strApp1FirstForeName = xmlGetAttributeText(xmlTempNode, "FIRSTFORENAME")
        
        'Applicant 1 Second ForeName - Char(1)
        strApp1SecondForeName = xmlGetAttributeText(xmlTempNode, "SECONDFORENAME")
        
        'Applicant 1 Other ForeNames - Char(1)
        strApp1OtherForeNames = xmlGetAttributeText(xmlTempNode, "OTHERFORENAMES")
        
        'Surname - Char(25)
        strApp1SurName = xmlGetAttributeText(xmlTempNode, "SURNAME")
    End If
    
    'Applicant 2 Details
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMER")
    If Not xmlTempNode Is Nothing Then
        'Applicant 2 Title - Char(8)
        strApp2Title = xmlGetAttributeText(xmlTempNode, "TITLE_TEXT")
    
        'Applicant 2 First ForeName - Char(1)
        strApp2FirstForeName = xmlGetAttributeText(xmlTempNode, "FIRSTFORENAME")
    
        'Applicant 2 Second ForeName - Char(1)
        strApp2SecondForeName = xmlGetAttributeText(xmlTempNode, "SECONDFORENAME")
    
        'Applicant 2 Other ForeNames - Char(1)
        strApp2OtherForeNames = xmlGetAttributeText(xmlTempNode, "OTHERFORENAMES")
        
        'Applicant 2 Surname - Char(25)
        strApp2SurName = xmlGetAttributeText(xmlTempNode, "SURNAME")
    End If
    
    'Valuation Fee - Number(7)
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/FEES/MORTGAGEONEOFFCOST[@MORTGAGEONEOFFCOSTTYPE='3']")
    If Not xmlTempNode Is Nothing Then
        If IsValidationType("OneOffCost", 3, "VAL") Then
            strValuationFee = xmlGetAttributeText(xmlTempNode, "AMOUNT")
        End If
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='1' and @CUSTOMERROLETYPE='1']/CUSTOMER/CUSTOMERADDRESS")
    If Not xmlTempNode Is Nothing Then
        
        'Flat Number & HouseName - Char(30)
        strAppCorrespFlatNo = xmlGetAttributeText(xmlTempNode, "FLATNUMBER")
        strAppCorrespHouseName = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENAME")
        
        'House Number - Char(5)
        strAppCorrespHouseNo = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENUMBER")
        
        'Street Name - Char(30)
        strAppCorrespStreetName = xmlGetAttributeText(xmlTempNode, "STREET")
        
        'District - Char(30)
        strAppCorrespDistrict = xmlGetAttributeText(xmlTempNode, "DISTRICT")
    
        'Town - Char(30)
        strAppCorrespTown = xmlGetAttributeText(xmlTempNode, "TOWN")
        
        'County - Char(30)
        strAppCorrespCounty = xmlGetAttributeText(xmlTempNode, "COUNTY")
        
        'PostCode - Char(9)
        strAppCorrespPostCode = xmlGetAttributeText(xmlTempNode, "POSTCODE")
        
    End If

    'Home Telephone Details of Applicant 1
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMER/CUSTOMERTELEPHONENUMBER[@USAGE='1']")
    If Not xmlTempNode Is Nothing Then
        strApp1HomeSTDCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        strApp1HomeTeleNo = xmlGetAttributeText(xmlTempNode, "TELEPHONENUMBER")
   End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='1' and @CUSTOMERROLETYPE='1']/CUSTOMER/ACCOUNTRELATIONSHIP/MORTGAGEACCOUNT")
    If Not xmlTempNode Is Nothing Then
        If xmlGetAttributeText(xmlTempNode, "REMORTGAGEINDICATOR") = "1" Then
            strRemortgageInd = "Y"
        Else
            strRemortgageInd = "N"
        End If
    End If
    
    'Work Telephone Details of Applicant 1
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/APPLICATIONCUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMER/CUSTOMERTELEPHONENUMBER[@USAGE='2']")
    If Not xmlTempNode Is Nothing Then
        strApp1WorkSTDCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        strApp1WorkTeleNo = xmlGetAttributeText(xmlTempNode, "TELEPHONENUMBER")
    End If
    
    'Selling Agent
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/ESTATEAGENT")
    If Not xmlTempNode Is Nothing Then
        strSellingAgent = xmlGetAttributeText(xmlTempNode, "COMPANYNAME")
    End If
    
    'ARRANGEMENTSFORACCESS  = 1 (Vendor ) or  2 ( Others)
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/NEWPROPERTY[@ARRANGEMENTSFORACCESS='1']/NEWPROPERTYVENDOR[@USAGE='10']")
                                                                               
    'If ARRANGEMENTSFORACCESS = 1 not found then find for 2 where 1 = Vendor and 2 = Others
    If Not xmlTempNode Is Nothing Then
        
        strArrangementAccess = "1"
        
        'Arranger Address Flat Number & HouseName - Char(30)
        strArrgAddrssFlatNo = xmlGetAttributeText(xmlTempNode, "FLATNUMBER")
        strArrgAddrssHouseName = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENAME")
        
        'Arranger Address House Number - Char(5)
        strArrgAddrssHouseNo = xmlGetAttributeText(xmlTempNode, "BUILDINGORHOUSENUMBER")
        
        'Arranger Address Street Name - Char(30)
        strArrgAddrssStreetName = xmlGetAttributeText(xmlTempNode, "STREET")
        
        'Arranger Address District - Char(30)
        strArrgAddrssDistrict = xmlGetAttributeText(xmlTempNode, "DISTRICT")
    
        'Arranger Address Town - Char(30)
        strArrgAddrssTown = xmlGetAttributeText(xmlTempNode, "TOWN")
        
        'Arranger Address County - Char(30)
        strArrgAddrssCounty = xmlGetAttributeText(xmlTempNode, "COUNTY")
        
        'Arranger Address PostCode - Char(9)
        strArrgAddrssPostCode = xmlGetAttributeText(xmlTempNode, "POSTCODE")
        
        'Arranger Contact Title
        strArrgAddrssCompanyName = xmlGetAttributeText(xmlTempNode, "COMPANYNAME")
        
        'Arranger Contact Telephone STD Code
        strArrgAddrssAreaCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        
        'Arranger Contact Telephone Number
        strArrgAddrssAreaTeleNo = xmlGetAttributeText(xmlTempNode, "TELENUMBER")
        
        'Arranger Contact Title
        strArrgContactTitleText = xmlGetAttributeText(xmlTempNode, "CONTACTTITLE_TEXT")
        
        'Arranger Conact Title
        strArrgContactForeName = xmlGetAttributeText(xmlTempNode, "CONTACTFORENAME")
        
        'Arranger Conact Title
        strArrgContactSurname = xmlGetAttributeText(xmlTempNode, "CONTACTSURNAME")
        
    Else
        If xmlTempNode Is Nothing Then
            Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/NEWPROPERTY[@ARRANGEMENTSFORACCESS='99']")
        End If
        
        If Not xmlTempNode Is Nothing Then
            strArrgContactName = xmlGetAttributeText(xmlTempNode, "ACCESSCONTACTNAME")
            strArrgAddrssAreaCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
            strArrgAddrssAreaTeleNo = xmlGetAttributeText(xmlTempNode, "ACCESSTELEPHONENUMBER")
        End If
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/NEWPROPERTY[@ARRANGEMENTSFORACCESS='2']/NEWPROPERTYADDRESS")
    If Not xmlTempNode Is Nothing Then
        strSellingAgentName = xmlGetAttributeText(xmlTempNode, "ACCESSCONTACTNAME")
        strSellingAgentContactSTDCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        strSellingAgentContactTeleNo = xmlGetAttributeText(xmlTempNode, "ACCESSTELEPHONENUMBER")
    End If
        
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/MORTGAGEPRODUCT/MORTGAGELENDER/MORTGAGELENDERDIRECTORY/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@TELEPHONESEQNUM='1']")
    If Not xmlTempNode Is Nothing Then
        strLenderContactAreaCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        strLenderContactTeleNo = xmlGetAttributeText(xmlTempNode, "TELENUMBER")
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/USERHISTORY")
    If Not xmlTempNode Is Nothing Then
        strUnitCode = xmlGetAttributeText(xmlTempNode, "UNITID")
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/USERHISTORY/USERROLE/OMIGAUSER/ORGANISATIONUSER")
    If Not xmlTempNode Is Nothing Then
        strLenderStaffContactTitle = xmlGetAttributeText(xmlTempNode, "USERTITLE_TEXT")
        strLenderStaffContactForeName = xmlGetAttributeText(xmlTempNode, "USERFORENAME")
        strLenderStaffContactInitials = xmlGetAttributeText(xmlTempNode, "USERINITIALS")
        strLenderStaffContactSurName = xmlGetAttributeText(xmlTempNode, "USERSURNAME")
    Else
    
    End If
    
    Set xmlTempNode = xmlRequest.selectSingleNode("//RESPONSE/APPLICATION/USERHISTORY/USERROLE/UNIT/UNITCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@TELEPHONESEQNUM='1']")
    If Not xmlTempNode Is Nothing Then
        strLenderStaffContactSTDCode = xmlGetAttributeText(xmlTempNode, "AREACODE")
        strLenderStaffContactTeleNo = xmlGetAttributeText(xmlTempNode, "TELENUMBER")
    End If
        
    'Flat No + House Name
    strDatablock = strDatablock & GetFormattedValue(strFlatNo & strHouseName, 30, "STRING")
    
    'House Number
    strDatablock = strDatablock & GetFormattedValue(strHouseNo, 5, "STRING")
    
    'Street Name
    strDatablock = strDatablock & GetFormattedValue(strStreetName, 30, "STRING")
    
    'District
    strDatablock = strDatablock & GetFormattedValue(strDistrict, 30, "STRING")
    
    'Town
    strDatablock = strDatablock & GetFormattedValue(strTown, 30, "STRING")
    
    'County
    strDatablock = strDatablock & GetFormattedValue(strCounty, 30, "STRING")
    
    'PostCode
    strDatablock = strDatablock & GetFormattedValue(strPostCode, 9, "STRING")
    
    'Year Built
    strDatablock = strDatablock & GetFormattedValue(strYearBuilt, 4, "STRING")
    
    'Tenure
    strDatablock = strDatablock & GetFormattedValue(strTenure, 2, "NUMBER")
    
    'Purchase Price
    strDatablock = strDatablock & GetFormattedValue(strPurchasePrice, 7, "NUMBER")
    
    'Advance Amount
    strDatablock = strDatablock & GetFormattedValue(strAdvanceAmount, 7, "NUMBER")
    
    'Applicant 1 Details
    strDatablock = strDatablock & GetFormattedValue(strApp1Title, 8, "STRING")
    
    strApp1FirstForeName = Left(strApp1FirstForeName, 1)
    strApp1OtherForeNames = Left(strApp1OtherForeNames, 1)
    strDatablock = strDatablock & GetFormattedValue(strApp1FirstForeName & " " & strApp1OtherForeNames, 3, "STRING")
    
    strDatablock = strDatablock & GetFormattedValue(strApp1SurName, 25, "STRING")
    
    'Applicant 2 Details
    strDatablock = strDatablock & GetFormattedValue(strApp2Title, 8, "STRING")
    
    strApp2FirstForeName = Left(strApp2FirstForeName, 1)
    strApp2OtherForeNames = Left(strApp2OtherForeNames, 1)
    strDatablock = strDatablock & GetFormattedValue(strApp2FirstForeName & " " & strApp2OtherForeNames, 3, "STRING")
    
    strDatablock = strDatablock & GetFormattedValue(strApp2SurName, 25, "STRING")
    
    'Survey Type
    strDatablock = strDatablock & GetFormattedValue(strValuationType, 2, "NUMBER")
    
    'Valuation Fee
    strDatablock = strDatablock & GetFormattedValue(strValuationFee, 7, "NUMBER")
    
    'Selling Agent
    strDatablock = strDatablock & GetFormattedValue(strSellingAgent, 30, "STRING")
    
    'Contact Details
    strDatablock = strDatablock & GetFormattedValue(strSellingAgentName & " " & strSellingAgentContactSTDCode & strSellingAgentContactTeleNo, 58, "STRING")
    
    'Right To buy Ind
    strDatablock = strDatablock & Space$(1)
    
    'Remortgage Indicator
    strDatablock = strDatablock & GetFormattedValue(strRemortgageInd, 1, "STRING")
    
    'Applicant 1 Corrrespondece Address
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespFlatNo & strAppCorrespHouseName, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespHouseNo, 5, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespStreetName, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespDistrict, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespTown, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespCounty, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strAppCorrespPostCode, 9, "STRING")
    
    'Applicant 1 Work Telephone
    strDatablock = strDatablock & GetFormattedValue(strApp1WorkSTDCode & " " & strApp1WorkTeleNo, 14, "STRING")
    
    'Applicant 1 Home Telephone ' TK 25/06/2004 BBG663
    strDatablock = strDatablock & GetFormattedValue(strApp1HomeSTDCode & " " & strApp1HomeTeleNo, 14, "STRING")
    
    'Set Unit Code
    strDatablock = strDatablock & GetFormattedValue(strUnitCode, 4, "STRING")
    
    'Rental Assessment Required
    If blnSpecialGroup Then
        strDatablock = strDatablock & "Y"
    Else
        strDatablock = strDatablock & "N"
    End If
    
    'Expected Monthly Rental Income
    strDatablock = strDatablock & GetFormattedValue(strMonthlyRentalIncome, 7, "NUMBER")
    
    'Arranger Address
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssFlatNo & strArrgAddrssHouseName, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssHouseNo, 5, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssStreetName, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssDistrict, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssTown, 30, "STRING")
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssCounty, 30, "STRING")
    
    'Arranger Post Code
    strDatablock = strDatablock & GetFormattedValue(strArrgAddrssPostCode, 9, "STRING")

    If strArrangementAccess = "1" Then
        strDatablock = strDatablock & GetFormattedValue(strArrgContactTitleText & " " & strArrgContactForeName & " " & strArrgContactSurname & " " & strArrgAddrssAreaCode & " " & strArrgAddrssAreaTeleNo, 58, "STRING")
    Else
        strDatablock = strDatablock & GetFormattedValue(strArrgContactName & " " & strArrgAddrssAreaCode & " " & strArrgAddrssAreaTeleNo, 58, "STRING")
    End If
    'Detailed Survey Quote Indicator
    strDatablock = strDatablock & GetFormattedValue(strDetSurveyQuoteInd, 1, "STRING")
    
    'Mortgage Purpose
    strDatablock = strDatablock & GetFormattedValue(strMortgagePurpose, 2, "NUMBER")
    
    'Valuation Fee Indicator
    If Len(strValuationFee) <= 0 Then
        strDatablock = strDatablock & GetFormattedValue("Y", 1, "STRING")
    Else
        strDatablock = strDatablock & GetFormattedValue("N", 1, "STRING")
    End If
    
    'Is Property In Scotland Indicator
    strDatablock = strDatablock & GetFormattedValue(strPropertyInScotland, 1, "STRING")
        
    'Lender Contact Telephone
    strDatablock = strDatablock & GetFormattedValue(strLenderContactAreaCode & " " & strLenderContactTeleNo, 14, "STRING")
    
    'Lender Staff Contact
    strLenderStaffContact = strLenderStaffContactTitle & " " & strLenderStaffContactForeName & " " & strLenderStaffContactInitials & " " & strLenderStaffContactSurName & " " & strLenderStaffContactSTDCode & " " & strLenderStaffContactTeleNo
    strDatablock = strDatablock & GetFormattedValue(strLenderStaffContact, 58, "STRING")
        
    CreateDataBlock = True
    
CreateDataBlockErr:

    Set xmlTempNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)

End Function


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
    Dim strFieldValue As String
    Dim strIsEmpty As String
    Dim strSeparator As String
    Dim intFieldSize As Integer
    Dim strFixedValue As String
    Dim strDataType As String
    Dim blnContinue As Boolean
    Dim strRet As String
    
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
    
    'If Empty return the string containing spaces or zeros -  Only pad if fixed length format
    If UCase(strIsEmpty) = "Y" And m_blnFixedLengthRecords Then
        If strDataType = "STRING" Or strDataType = "COMBO" Or strDataType = "DATE" Then
            GetFieldValueString = Space(intFieldSize) & strSeparator
        Else
            GetFieldValueString = String(intFieldSize, "0") & strSeparator
        End If
        Exit Function
    End If
    
    ' Format the value, if required
    strRet = GetFormattedValue(strFieldValue, intFieldSize, strDataType)
    If blnAddFieldSeparator Then
        strRet = strRet & strSeparator
    End If
    
    GetFieldValueString = strRet
    
End Function
Private Function WriteToTextFile() As Boolean
    
    On Error GoTo WriteToTextFileExit
    Const cstrFunctionName As String = "WriteToTextFile"
    
    Dim blnAppendMode As Boolean
    Dim intFileNo As Integer
    Dim strfilename As String
    Dim strFileLoc As String
    
    Dim objCRM As omCRMWrapper
    
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    
    Set xmlDoc = xmlCreateDOMObject
    
    blnAppendMode = GetMandatoryGlobalParamBoolean("QuestDataFileMode")
    If blnAppendMode Then
        'Append to a single output file
        strfilename = GetMandatoryGlobalParamString("QuestDataFileName")
    Else
        'Write to a file for the application
        strfilename = strApplicationNumber & ".txt"
    End If
    
    strFileLoc = GetMandatoryGlobalParamString("QuestDataOutputFolder")
    
    nMaxRetries = GetMandatoryGlobalParamAmount("LockTextFileRetries")
    nSleepMilliSecs = GetMandatoryGlobalParamAmount("LockTextFileSleep")
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("TEXTOUT")
    xmlElement.setAttribute "PATHANDFILENAME", strFileLoc & "\" & strfilename
    xmlElement.setAttribute "MAXRETRIES", nMaxRetries
    xmlElement.setAttribute "SLEEPMILLISECS", nSleepMilliSecs
    'WRITEMODE is either "Append" or OverWrite"
    xmlElement.setAttribute "WRITEMODE", "Append"
    xmlRequest.appendChild xmlElement
    If m_objContext Is Nothing Then
        Set objCRM = New omCRMWrapper
    Else
        Set objCRM = m_objContext.CreateInstance(gstrCRM_COMPONENT & ".omCRMWrapper")
    End If
    'pass string to be written direct, so we don't lose our CR/LF
    strResponse = objCRM.AddText(xmlRequest.xml, strDatablock)
                                               
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
WriteToTextFileExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function


Public Function CreateQuestReportData(ByVal vxmlRequestNode As String) As String
    
    On Error GoTo CreateQuestReportDataErr
    Const cstrFunctionName  As String = "CreateQuestReportData "
    
    Dim xmlResponse As IXMLDOMElement
    Dim xmlFormatTemplate As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Set xmlInDoc = xmlCreateDOMObject
    xmlInDoc.loadXML (vxmlRequestNode)
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Set xmlCaseTaskNode = xmlInDoc.selectSingleNode(".//CASETASK")
    
    'Get the actual data
    Set xmlResponse = GetData(vxmlRequestNode)
    errCheckXMLResponseNode xmlResponse, , True
    
    'Format the output
    If GetMandatoryGlobalParamBoolean("QuestFormat") Then
        m_blnFixedLengthRecords = False
        m_strFieldSeparator = "|"
    Else
        m_blnFixedLengthRecords = True
        m_strFieldSeparator = ""
    End If

    ' Generate data blocks
    If CreateDataBlock(xmlResponse, xmlCaseTaskNode) Then
        Call WriteToTextFile
    End If
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    CreateQuestReportData = xmlOutDoc.xml

CreateQuestReportDataExit:
    
    Set xmlResponse = Nothing
    Set xmlFormatTemplate = Nothing
    Set xmlResponseElem = Nothing
    Set xmlOutDoc = Nothing
        
    Exit Function
CreateQuestReportDataErr:

    CreateQuestReportData = errCreateErrorResponse()
    GoTo CreateQuestReportDataExit
    
End Function

Private Function LoadFormatTemplate() As FreeThreadedDOMDocument40
    
    On Error GoTo LoadFormatTemplateExit
    
    Const cstrFunctionName = "LoadFormatTemplate"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim hFile As Integer
    Dim strfilename As String
    
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strfilename = App.Path & "\QuestDataFormatTemplate.xml"
    strfilename = Replace(strfilename, "DLL", "XML", 1, 1, vbTextCompare)
    hFile = FreeFile
    
    Open strfilename For Input Access Read As hFile
    Close hFile
    
    'If anything now breaks it will be XML related.
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    xmlDoc.Load (strfilename)
    
    Set LoadFormatTemplate = xmlDoc
    
    If xmlDoc.parseError.errorCode <> 0 Then
        errThrowError cstrFunctionName, oeXMLParserError
    End If
    
LoadFormatTemplateExit:
    
    
    errCheckError cstrFunctionName, TypeName(Me)
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
    Else
        Select Case strDataType
            Case "STRING"  'add trailing spaces
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
            Case "NUMBER"
                If intValueLength < intFieldSize Then  'actual length is less than the expected in record
                    'Only pad field if using fixed length format
                    If m_blnFixedLengthRecords Then
                        'Pad with leading zeros
                        'GetFormattedValue = String(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                        GetFormattedValue = MakeString(intFieldSize - intValueLength, cstrLeadingCharForNumbers) & strFieldValue
                    Else
                        GetFormattedValue = strFieldValue
                    End If
                ElseIf intValueLength > intFieldSize Then 'actual length is more than the expected in record
                    'Trim right side
                    GetFormattedValue = Left(strFieldValue, intFieldSize)
                Else  'actual length is same as expected
                    GetFormattedValue = strFieldValue
                End If
            Case "DATE"
                GetFormattedValue = Left(strFieldValue, 2) & Mid(strFieldValue, 4, 2) & Mid(strFieldValue, 7, 4)
            Case Else
                errThrowError "omQuestBO.GetFormattedValue", oeInValidDataTypeValue
        End Select
    End If
End Function

Private Function MakeString(ByVal intLength As Integer, ByVal strChar As String) As String
    If intLength > 0 Then
        MakeString = String(intLength, strChar)
    Else
        MakeString = strChar
    End If
End Function

Private Function GetData(ByVal vstrxmlRequest As String) As IXMLDOMNode
    
    On Error GoTo GetDataExit
    
    Const cstrFunctionName As String = "GetData"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim vxmlRequest  As IXMLDOMNode
    Dim xmlApplication As IXMLDOMElement
    Dim objRBDO As OmRequestDO
    
    Dim strAFFNumber As String
    Dim strResponse As String
    
    'Get Application Number from Request
    Set xmlDoc = xmlLoad(vstrxmlRequest, cstrFunctionName)
    Set xmlApplication = xmlGetMandatoryNode(xmlDoc.documentElement, ".//APPLICATION")
    
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
        
    Set vxmlRequest = xmlDoc.documentElement
    
    xmlSetAttributeValue vxmlRequest, "RB_TEMPLATE", "QuestDataTemplate"
    
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequest, ".//APPLICATION")
    xmlSetAttributeValue xmlApplication, "APPLICATIONNUMBER", strApplicationNumber
    xmlSetAttributeValue xmlApplication, "APPLICATIONFACTFINDNUMBER", strAFFNumber
    
    ' Call Request Broker component omRB
    'MSla 01/06/2004 BBG614
    If m_objContext Is Nothing Then
        Set objRBDO = New OmRequestDO
    Else
        Set objRBDO = m_objContext.CreateInstance(gstrREQUEST_BROKER_COMPONENT & ".OmRequestDO")
    End If
    'Set objRBDO = m_objContext.CreateInstance(gstrREQUEST_BROKER_COMPONENT & ".OmRequestDO")
    'MSla 01/06/2004 BBG614 - End
    
    ' Note the Reuqest is run with UserId and and other attrubutes as mentioned RequestXML file
    strResponse = objRBDO.OmDataRequest(vxmlRequest.xml)
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    ' PRW BBG726 - Code removed for this AQR
    'Testing - To save the output to a file
'    Dim strfilename As String
'    strfilename = "C:\temp\WP19\Output\Output_" & Format(Now(), "HH_MM") & ".xml"
'    xmlDoc.save (strfilename)
    ' End BBG726
    
    Set GetData = xmlDoc.documentElement
    
GetDataExit:
    
    Set xmlDoc = Nothing
    Set vxmlRequest = Nothing
    Set xmlApplication = Nothing
    Set objRBDO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

Private Sub CreateFolderEx(ByVal strPath)

'    On Error GoTo CreateFolderExExit
'    Const cstrFunctionName As String = "CreateFolderEx"
'
'    Dim arrFolders() As String
'    Dim intLevels As Integer
'    Dim strLowest As String
'    Dim strWorking As String
'
'    If m_objFSO.FolderExists(strPath) = False Then
'
'        'Replace strPath, "/", "\"
'        arrFolders = Split(strPath, "\")
'        intLevels = UBound(arrFolders)
'
'        If intLevels > 0 Then
'            strLowest = arrFolders(intLevels)
'            ReDim Preserve arrFolders(intLevels - 1)
'            strWorking = Join(arrFolders, "\")
'            CreateFolderEx strWorking
'            'On Error Resume Next
'            m_objFSO.CreateFolder strWorking & "\" & strLowest
'        End If
'
'    End If
'
'CreateFolderExExit:
'
'    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo DoRequestExit
    Const strFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
        ' debug only
        Case "CREATEQUESTREPORTDATA"
            'CreateQuestReportData    vxmlRequestNode, vxmlResponseNode
    End Select
DoRequestExit:
    
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Public Function omRequest(ByVal vstrXmlIn As String) As String
    
    On Error GoTo omRequestVbErr
    Const strFunctionName As String = "omRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' will be superceded by any subsequent SetAbort
    m_objContext.SetComplete
    
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    omRequest = xmlOutDoc.xml
    
omRequestExit:
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function
    
omRequestVbErr:
    
   ' omRequest  errCreateErrorResponse()
    GoTo omRequestExit
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
   ' Set m_objFSO = New Scripting.FileSystemObject
End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    'Set m_objFSO = Nothing
End Sub


