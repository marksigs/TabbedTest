VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThirdPartyClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IThirdPartyClassDef
'GD ADDED

'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'MDC    27/11/2002  BMIDS00900 - Add DXID and DXLocation to ThirdParty
'GD     21/05/2003  BM0370 Change Type of DXID
'------------------------------------------------------------------------------------------


Private Function IThirdPartyClassDef_LoadPanelValuerListData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelValuerListData"
    
    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>VA_PANELVALUERLIST"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NAMEANDADDRESSTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>COMPANYNAME<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>HEADOFFICEINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>PANELID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ASSOCIATEDPANELID<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "<OTHERS>PAYMENTMETHOD<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PaymentMethod</COMBO></OTHERS>"

    
    
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "<OTHERS>COUNTRY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Country</COMBO></OTHERS>"
    
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    
    'JR - Omiplus24
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TELENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactTelephoneUsage</COMBO></OTHERS>"
    
    'GD ADDED
    'strXML = strXML & "<PRIMARYKEY>VALUERTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<PRIMARYKEY>VALUERTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ValuerType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>QUALIFICATIONS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>VALUATIONTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>VALUATIONTYPETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IThirdPartyClassDef_LoadPanelValuerListData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IThirdPartyClassDef_LoadDirectoryData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IThirdPartyClassDef_LoadDirectoryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>NAMEANDADDRESSDIRECTORY"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>NAMEANDADDRESSTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>ThirdPartyType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HEADOFFICEINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NAMEANDADDRESSACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NAMEANDADDRESSACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>"
    'GD BM0370 START
    'strXML = strXML & "<OTHERS>DXID<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DXID<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BM0370 END

    strXML = strXML & "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NOTES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORGANISATIONTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>OrganisationType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>BRANCHNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NAMEANDADDRESSBANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function


Private Function IThirdPartyClassDef_LoadDirectoryDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadDirectoryDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objAddressClassDef As IAddressClassDef, _
        objContactClassDef As IContactClassDef
    Dim objXmlAssist As New XMLAssist

    Set objAddressClassDef = New AddressClassDef
    Set objContactClassDef = New ContactClassDef

    strXML = strXML & "<TABLENAME>DIRECTORYDETAILS"
    strXML = strXML & IThirdPartyClassDef_LoadDirectoryData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    strXML = strXML & objContactClassDef.LoadContactData.xml
    strXML = strXML & IThirdPartyClassDef_LoadPanelData.xml
    strXML = strXML & IThirdPartyClassDef_LoadPanelBankAccountData.xml
    strXML = strXML & IThirdPartyClassDef_LoadPanelValuerData.xml
    strXML = strXML & IThirdPartyClassDef_LoadPanelValuerValuationTypeData.xml
    strXML = strXML & IThirdPartyClassDef_LoadPanelLegalRepData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadDirectoryDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objAddressClassDef = Nothing
    Set objContactClassDef = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IThirdPartyClassDef_LoadDirectorySummaryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadDirectorySummaryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objAddressClassDef As IAddressClassDef, _
        objContactClassDef As IContactClassDef
    Dim objXmlAssist As New XMLAssist
    
    Set objAddressClassDef = New AddressClassDef
    Set objContactClassDef = New ContactClassDef

    strXML = strXML & "<TABLENAME>DIRECTORYSUMMARY"
    strXML = strXML & IThirdPartyClassDef_LoadDirectoryData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    strXML = strXML & objContactClassDef.LoadContactData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadDirectorySummaryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objAddressClassDef = Nothing
    Set objContactClassDef = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IThirdPartyClassDef_LoadMortgageLenderDirectoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadMortgageLenderDirectoryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>MORTGAGELENDERDIRECTORY"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>MAINMORTGAGELENDERIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadMortgageLenderDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IThirdPartyClassDef_LoadPanelBankAccountData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelBankAccountData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PANELBANKACCOUNT"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BANKACCOUNTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>BANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTCURRENCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>AccountCurrency</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTCURRENCYTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPanelBankAccountData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadPanelData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PANEL"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>PANELID<TYPE>dbdtString</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>ASSOCIATEDPANELID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTMETHOD<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>PaymentMethod</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPanelData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadPanelLegalRepData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelLegalRepData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PANELLEGALREP"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>TYPEOFLEGALREP<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>TypeOfLegalRep</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFPARTNERS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SENIORPARTNERDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INDEMNITYINSURANCEEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYOFPRACTICE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>CountryOfPractice</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYOFORIGIN<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>CountryOfOrigin</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TEMPORARYAPPOINTMENTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    '~ MF 12/08/2005 MAR20
    strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPanelLegalRepData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadPanelValuerData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelValuerData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PANELVALUER"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>QUALIFICATIONS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPanelValuerData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadPanelValuerValuationTypeData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPanelValuerValuationTypeData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PANELVALUERVALUATIONTYPE"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>VALUATIONTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>ValuationType</COMBO></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>VALUATIONTYPETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPanelValuerValuationTypeData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadThirdPartyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadThirdPartyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>THIRDPARTY"
    strXML = strXML & "<PRIMARYKEY>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>THIRDPARTYTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>ThirdPartyType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORGANISATIONTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "  <COMBO>OrganisationType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>BRANCHNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYBANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    'BMIDS00900 MDC 27/11/2002
    
    'GD BM0370 START
    'strXML = strXML & "<OTHERS>DXID<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DXID<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BM0370 END
    
    
    
    strXML = strXML & "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00900 MDC 27/11/2002 - End
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadThirdPartyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IThirdPartyClassDef_LoadThirdPartyDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadThirdPartyDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objAddressClassDef As IAddressClassDef, _
        objContactClassDef As IContactClassDef
    Dim objXmlAssist As New XMLAssist
    
    Set objAddressClassDef = New AddressClassDef
    Set objContactClassDef = New ContactClassDef

    strXML = strXML & "<TABLENAME>THIRDPARTYDETAILS"
    strXML = strXML & IThirdPartyClassDef_LoadThirdPartyData.xml
    strXML = strXML & objAddressClassDef.LoadAddressData.xml
    'strXML = strXML & objContactClassDef.LoadContactDetailsData.xml
    strXML = strXML & objContactClassDef.LoadContactData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadThirdPartyDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objAddressClassDef = Nothing
    Set objContactClassDef = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IThirdPartyClassDef_LoadMainMortgageLenderData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadMainMortgageLenderData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>MAINMORTGAGELENDER"
    strXML = strXML & IThirdPartyClassDef_LoadDirectoryData.xml
    strXML = strXML & IThirdPartyClassDef_LoadMortgageLenderDirectoryData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadMainMortgageLenderData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IThirdPartyClassDef_LoadPaymentDetailsData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IThirdPartyClassDef_LoadPaymentDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PAYMENTDETAILS"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTMETHOD<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ProcurationFeePaymentMethod</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>BANKSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IThirdPartyClassDef_LoadPaymentDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

