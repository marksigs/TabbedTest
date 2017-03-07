VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definitions for omCust
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/12/99    Created
'JR     03/09/01    Omiplus24 - Modified IClassDef_LoadCustomerTelephoneData
'------------------------------------------------------------------------------------------

Option Explicit

Implements IClassDef

Private Function IClassDef_LoadAliasData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : DOM Document
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadAliasData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "ALIAS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ALIASSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ALIASTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFCHANGE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>METHODOFCHANGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasMethodOfChange</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadAliasData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadAreasOfInterestData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IClassDef_LoadAreasOfInterestData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>AREASOFINTEREST"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>INTERESTAREA<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"

    Set IClassDef_LoadAreasOfInterestData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadCustomerAddressData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : Dom Document (Table Definition)
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadCustomerAddressData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMERADDRESS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CustomerAddressType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDIN<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDOUT<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATUREOFOCCUPANCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>NatureOfOccupancy</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadCustomerAddressData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadCustomerTelephoneData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
    Const strFunctionName As String = "IClassDef_LoadCustomerTelephoneData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>CUSTOMERTELEPHONENUMBER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TELEPHONESEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTIME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadCustomerTelephoneData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadAliasPersonData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : XML string
' **************************************************************************
     Const strFunctionName As String = "IClassDef_LoadAliasPersonData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FINDALIASLIST"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ALIASSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ALIASTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFCHANGE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>METHODOFCHANGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasMethodOfChange</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
'Person Data
    strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GENDER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Sex</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MARITALSTATUS<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>MaritalStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALITY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Nationality</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERFORENAMES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadAliasPersonData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadCurrentPropertyData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : DOM Document
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadCurrentPropertyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CURRENTPROPERTY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyDescription</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ESTIMATEDCURRENTVALUE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTENDEDACTION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CurrentPropertyIntendedAction</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUATIONAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TENURE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEARBUILT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REINSTATEMENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTVALUERNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadCurrentPropertyData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadHomeInsuranceData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - HomeInsurance Table
'Pass        : None
'Return      : XML Document
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadHomeInsuranceData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "HOMEINSURANCE"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HOMEINSURANCETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>InsuranceType</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadHomeInsuranceData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadTenancyData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : XML string
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadTenancyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "TENANCY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TENANCYTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>LandlordType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYRENTAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadTenancyData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadCustomerData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - Customer Table
'Pass        : None
'Return      : XML string
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadCustomerData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>OTHERSYSTEMCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERSYSTEMTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadCustomerData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IClassDef_LoadCustomerVersionData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - Customer Version Table
'Pass        : None
'Return      : XML string
' **************************************************************************
    Const strFunctionName As String = "IClassDef_LoadCustomerVersionData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMERVERSION"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CAROWNER<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTEMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GENDER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Sex</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MAILSHOTREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MARITALSTATUS<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>MaritalStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MEMBEROFSTAFF<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALINSURANCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALITY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Nationality</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERFORENAMES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SMOKERSTATUS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEARADDEDTOVOTERSROLL<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTSYNCJOBTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MOTHERSMAIDENNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NORMALLYRESIDENTINUK<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GOODHEALTH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALIASINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UKRESIDENTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DEPENDANTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadCustomerVersionData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    
End Function

Private Function IClassDef_LoadDependantDetailsData() As FreeThreadedDOMDocument40
'Header------------------------------------------------------------------------------------
'Description : Creates the template required for getting XML for Dependant and Person data
'              Builds  template by adding the definitions of Dependant and Person
'Pass        : None
'Return      : XML string
' *****************************************************************************************
    Const strFunctionName As String = "IClassDef_LoadDependantDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objIPersonClassDef As IPersonClassDef

    strXML = strXML & "<TABLENAME>DEPENDANTDETAILS"
    
    Set xmlClassDefDoc = IClassDef_LoadDependantData
    strXML = strXML & xmlClassDefDoc.xml
    
    Set objIPersonClassDef = New PersonClassdef
    Set xmlClassDefDoc = objIPersonClassDef.LoadPersonData
    strXML = strXML & xmlClassDefDoc.xml
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadDependantDetailsData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objIPersonClassDef = Nothing

End Function

Private Function IClassDef_LoadDependantData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
    Const strFunctionName As String = "IClassDef_LoadDependantData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>DEPENDANT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>DEPENDANTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DEPENDANTRELATIONSHIP<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>DependantRelationship</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadDependantData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    
End Function



Private Function IClassDef_LoadUnitData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'history:
'BG SYS1860 22/01/01 Added Method for omApp.ApplicationDO.GetUnitFromApplication
'
' *****************************************************************************
    Const strFunctionName As String = "IClassDef_LoadUnitData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>USERHISTORY"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>USERHISTORYDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IClassDef_LoadUnitData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

