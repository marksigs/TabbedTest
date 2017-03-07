VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omCustClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definitions for omCust
'
'-----------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/12/99    Created
'SR     17/04/00    Added enumerations to build classdef with DateOfBirth alone for
'                   CustomerVersion
'IW     19/04/00    Added Work Extension Number
'IW     05/05/00    Added PreferredMethod
'IW     25/05/00    Added EmailPreferred to CustomerVersion class def.
'MC     01/06/00    Added Tax Details option to LoadCustomerVersionData
'MC     14/06/00    SYS0935 Amended LoadCustomerTelephoneData to use PREFERREDMETHODOFCONTACT
'JR     28/06/01    Omiplus24 Amended LoadCustomerTelephoneData to include AREA AND COUNTRY CODE
'DRC    09/10/01    SYS2754 ammended LoadHomeInsuranceData to include all Primary keys
'JLD    12/12/01    SYS3366 dbdttype wrong for telenumbers
'MC     21/05/2004  BMIDS747    VERIFICATION record change
'                   New Fields added WhoSeenCustomer, HowCustomerIDSeen to "CustomerVersion" DB Table
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
' Prog  Date        Description
' PSC   28/06/2005  MAR5 - Added LoadCustomerAddressAndRoleData and LoadCustomerNameData
' MF    01/08/2005  MARS019 New combo added: Special needs. Added to CustomerVersion Table
'                   Change in IomCustClassDef_LoadCustomerVersionData
' MF    09/08/2005  MAR20 New fields added: TIMEATBANKYEARS & TIMEATBANKMONTHS to CutomerVersion
'HM     08/09/2005  MAR29   IomCustClassDef_LoadCustomerVersionData was extended
'                           IomCustClassDef_LoadWrapUpDetailsData was added
'PSC    12/10/2005  MAR57 added CUSTOMERSTATUS and CUSTOMERCATEGORY
'HMA    20/10/2005  MAR23 Amended LoadCustomerVersionData and added LoadKYCData
'PSC    12/01/2006  MAR1033 Amended LoadCustomerAddressData to add LASTAMENDEDDATE
'PSC    24/01/2006  MAR1102 Amended LoadCustomerAddressData to make LASTAMENDEDDATE datetime
'MAH    14/11/2006  E2CR35  Amended LoadCstomerVersionData adding HousingBenefit, and Residency
'------------------------------------------------------------------------------------------

Option Explicit

Implements IomCustClassDef

Private Function IomCustClassDef_LoadAliasData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : DOM Document
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadAliasData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "ALIAS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    14/10/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>ALIASSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ALIASTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFCHANGE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>METHODOFCHANGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>AliasMethodOfChange</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    ' DRC 10/02/2004  BMIDS693
    strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadAliasData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function


Private Function IomCustClassDef_LoadAreasOfInterestData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCustClassDef_LoadAreasOfInterestData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>AREASOFINTEREST"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>INTERESTAREA<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadAreasOfInterestData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCustClassDef_LoadContactHistoryData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "IomCustClassDef_LoadContactHistoryData"
    
    Dim objXmlAssist As New XMLAssist
    
    ' PSC 21/10/2002 BMIDS00030
    Const cstrXML As String = "<TABLENAME>CONTACTHISTORY" & _
    "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
    "<PRIMARYKEY>CONTACTHISTORYDATETIME<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>" & _
    "<OTHERS>CONTACTREASONCODE<TYPE>dbdtComboId</TYPE>" & _
    "<COMBO>ContactHistoryReason</COMBO></OTHERS>" & _
    "<OTHERS>OTHERREASONTEXT<TYPE>dbdtString</TYPE></OTHERS>" & _
    "<OTHERS>CONTACTTEXT<TYPE>dbdtString</TYPE></OTHERS>" & _
    "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" & _
    "<OTHERS>UNITID<TYPE>dbdtString</TYPE></OTHERS>" & _
    "<OTHERS>STATUSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
    "<OTHERS>INTERFACEDTOEXTERNALSYSTEM<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
    "</TABLENAME>"

    Set IomCustClassDef_LoadContactHistoryData = objXmlAssist.load(cstrXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing

End Function



Private Function IomCustClassDef_LoadCustomerAddressData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : Dom Document (Table Definition)
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadCustomerAddressData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMERADDRESS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    14/10/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CustomerAddressType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDIN<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDOUT<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATUREOFOCCUPANCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>NatureOfOccupancy</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    ' PSC 12/01/2006 MAR1033
    ' PSC 24/01/2006 MAR1102
    strXML = strXML & "<OTHERS>LASTAMENDEDDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadCustomerAddressData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadCustomerTelephoneData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerTelephoneNumber Table
'Pass        : None
'Return      : Dom Document (Table Definition)
' *****************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadCustomerTelephoneData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>CUSTOMERTELEPHONENUMBER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    14/10/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>TELEPHONESEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTIME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PREFERREDMETHODOFCONTACT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadCustomerTelephoneData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function



Private Function IomCustClassDef_LoadAliasPersonData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : XML string
' **************************************************************************
     Const strFunctionName As String = "IomCustClassDef_LoadAliasPersonData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FINDALIASLIST"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    14/10/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>ALIASSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    ' DRC 10/02/2004  BMIDS693
    strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
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

    Set IomCustClassDef_LoadAliasPersonData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadCurrentPropertyData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : DOM Document
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadCurrentPropertyData"

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

    Set IomCustClassDef_LoadCurrentPropertyData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadDefinition(ByVal strTableName As String) _
                                                    As FreeThreadedDOMDocument40
'Header------------------------------------------------------------------------------
'Description : Load appropriate class defintion based on the table name passed in.
'              Calls the appropriate load method
'Pass        : strTableName - Name of the table, for which the definition is required
'Return      : XML Document
' ************************************************************************************
On Error GoTo LoadDefinitionVbErr
    
    Const strFunctionName As String = "IomCustClassDef_LoadDefinition"
    
    Dim objErrAssist As New ErrAssist
    
    Select Case strTableName
        Case "CUSTOMER"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadCustomerData
        Case "CUSTOMERVERSION"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadCustomerVersionData
        Case "CUSTOMERADDRESS"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadCustomerAddressData
        Case "CUSTOMERTELEPHONE"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadCustomerTelephoneData
        Case "CUSTOMERWRAPUPDETAILS"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadWrapUpDetailsData
        Case "ALIAS"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadAliasData
        Case "DEPENDANT"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadDependantData
        Case "TENANCY"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadTenancyData
        Case "HOMEINSURANCE"
            Set IomCustClassDef_LoadDefinition = IomCustClassDef_LoadHomeInsuranceData
        Case Else
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 522, _
                                    " For Loading the Class Definition"
    End Select
        
    Set objErrAssist = Nothing
    Exit Function
    
LoadDefinitionVbErr:
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function IomCustClassDef_LoadHomeInsuranceData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - HomeInsurance Table
'Pass        : None
'Return      : XML Document
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadHomeInsuranceData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "HOMEINSURANCE"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HOMEINSURANCETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>InsuranceType</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadHomeInsuranceData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadPropertyInsuranceData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCustClassDef_LoadPropertyInsuranceData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "QUICKQUOTEAPPLICANTDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ANNUALINCOMEORNETPROFIT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CCJDETAILSSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMPLOYMENTSTATUSSEQUENCENUMBER"
    strXML = strXML & "<TYPE>dbdtDouble</TYPE>"
    strXML = strXML & "<FORMATMASK>0.000</FORMATMASK>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>GOODHEALTH"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>SMOKER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMONTHLYOUTGOINGS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALDEBTSOUTSTANDINGBALANCE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALDEBTSMONTHLYPAYMENTS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UKTAXPAYER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadPropertyInsuranceData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IomCustClassDef_LoadTenancyData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerAddress Table
'Pass        : None
'Return      : XML string
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadTenancyData"

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

    Set IomCustClassDef_LoadTenancyData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadCustomerData( _
                            Optional ByVal intCustDataDefType As _
                            omCustDataDefType = omCustVersionDefault) As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - Customer Table
'Pass        : None
'Return      : XML string
' **************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadCustomerData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    Select Case intCustDataDefType
    
    Case omCustVersionDefault

        strXML = strXML & "<TABLENAME>"
        strXML = strXML & "CUSTOMER"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>OTHERSYSTEMCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OTHERSYSTEMTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        strXML = strXML & "</TABLENAME>"
        
    Case omCustVersionNumber
        strXML = strXML & "<TABLENAME>"
        strXML = strXML & "CUSTOMER"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>OTHERSYSTEMCUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "</TABLENAME>"
        
    Case Else:
        ' N/A
    End Select
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadCustomerData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCustClassDef_LoadCustomerVersionData( _
                            Optional ByVal intCustVersionDataDefType As _
                            omCustDataDefType = omCustVersionDefault) As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCustClassDef_LoadCustomerVersionData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    

    Select Case intCustVersionDataDefType
    
    Case omCustVersionDefault
    
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
        '~ MF 01/08/2005 MARS019 New combo added: Special needs
        strXML = strXML & "<OTHERS>SPECIALNEEDS<TYPE>dbdtComboId</TYPE>"
        strXML = strXML & "<COMBO>SpecialNeeds</COMBO></OTHERS>"
        '~ MF 09/08/2005 MAR20 New fields: TIMEATBANKYEARS & TIMEATBANKMONTHS
        strXML = strXML & "<OTHERS>TIMEATBANKYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TIMEATBANKMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
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
        strXML = strXML & "<OTHERS>PREFERREDNAME<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SUFFICIENTREPAYMENTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EMAILPREFERRED<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBEROFDEPENDANTS<TYPE>dbdtInt</TYPE></OTHERS>"
        'BMIDS00336 MDC 27/08/2002 - Credit Check & Bureau Download
        strXML = strXML & "<OTHERS>CREDITSEARCH<TYPE>dbdtBoolean</TYPE></OTHERS>"
        'BMIDS00336 MDC 27/08/2002 - Credit Check & Bureau Download - End
        
        '*=[MC] BMIDS747 - CC063 : HOW VERIFICATION DETAILS DC035
        strXML = strXML & "<OTHERS>WHOSEENCUSTOMER<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>HOWCUSTOMERIDSEEN<TYPE>dbdtInt</TYPE></OTHERS>"
        '*=[MC]SECTION END
        
        'BMIDS758 Add Transfer of Equity indicator
        strXML = strXML & "<OTHERS>NEWTOECUSTOMERIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        
        'JD BMIDS830
        strXML = strXML & "<OTHERS>OWNMORTGAGELENDERIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MORTGAGEREPAID<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MORTGAGELET<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AXESSAPPLICATION<TYPE>dbdtBoolean</TYPE></OTHERS>"
        'MAR29
        strXML = strXML & "<OTHERS>WRAPUPEMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WRAPUPEMAILPREFERRED<TYPE>dbdtBoolean</TYPE></OTHERS>"
        ' PSC 12/10/2005 MAR57 - Start
        strXML = strXML & "<OTHERS>CUSTOMERSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CUSTOMERCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CUSTOMERKYCSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CUSTOMERKYCADDRESSFLAG<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CUSTOMERKYCIDFLAG<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CUSTOMERDORMANTACCOUNTIND<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PROSPECTPASSWORDTAKEN<TYPE>dbdtBoolean</TYPE></OTHERS>"
        ' PSC 12/10/2005 MAR57 - End
        ' MAH 14/11/2006 E2CR35 - START
        strXML = strXML & "<OTHERS>HOUSINGBENEFIT<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>COUNTRYOFRESIDENCY<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>REASONFORCOUNTRYOFRESIDENCY<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RIGHTTOWORK<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DIPLOMATICIMMUNITY<TYPE>dbdtBoolean</TYPE></OTHERS>"
        ' MAH 14/11/2006 E2CR35 - END
        
        strXML = strXML & "</TABLENAME>"
    
    Case omCustVersionDOB
        strXML = strXML & "<TABLENAME>"
        strXML = strXML & "CUSTOMERVERSION"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "</TABLENAME>"
    Case omCustVersionTax
        strXML = strXML & "<TABLENAME>"
        strXML = strXML & "CUSTOMERVERSION"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>NATIONALINSURANCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "</TABLENAME>"
   'MAR23
    Case omCustVersionNumber
        strXML = strXML & "<TABLENAME>"
        strXML = strXML & "CUSTOMERVERSION"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "</TABLENAME>"
    Case Else:
        ' N/A
    End Select
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadCustomerVersionData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    
End Function



Private Function IomCustClassDef_LoadDependantDetailsData() As FreeThreadedDOMDocument40
'Header------------------------------------------------------------------------------------
'Description : Creates the template required for getting XML for Dependant and Person data
'              Builds  template by adding the definitions of Dependant and Person
'Pass        : None
'Return      : XML string
' *****************************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadDependantDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objIPersonClassDef As IPersonClassDef

    strXML = strXML & "<TABLENAME>DEPENDANTDETAILS"
    
    Set xmlClassDefDoc = IomCustClassDef_LoadDependantData
    strXML = strXML & xmlClassDefDoc.xml
    
    Set objIPersonClassDef = New PersonClassdef
    Set xmlClassDefDoc = objIPersonClassDef.LoadPersonData
    strXML = strXML & xmlClassDefDoc.xml
    
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadDependantDetailsData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objIPersonClassDef = Nothing

End Function


Private Function IomCustClassDef_LoadDependantData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadDependantData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>DEPENDANT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    ' DM    14/10/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>DEPENDANTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PERSONGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DEPENDANTRELATIONSHIP<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>DependantRelationship</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALDETAILS<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadDependantData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    
End Function


'BMIDS00005 CRWP2 BM065
Private Function IomCustClassDef_LoadApplicationContactData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindContactsByApplication view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadApplicationContactData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>FINDCONTACTSPERAPPLICATION"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>OMIGACUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CIFNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAGECHANGEDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFLOAN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFLOANCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONSTAGEDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICATIONSTAGENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTREASON<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERREASON<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadApplicationContactData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00005 End

''BMIDS00005 CRWP2 BM065
'Private Function IomCustClassDef_LoadApplicationContactData() As FreeThreadedDOMDocument40
'' header ----------------------------------------------------------------------------------
'' description:  Creates the class template required for getting the XML from a recordset - FindContactsByApplication view
'' pass:         None
'' return:       XML string
''------------------------------------------------------------------------------------------
'
'    Const strFunctionName As String = "IomCustClassDef_LoadApplicationContactData"
'
'    Dim strXML          As String
'    Dim objXmlAssist    As XMLAssist
'
'    Set objXmlAssist = New XMLAssist
'
'    strXML = "<TABLENAME>CONTACT"
'    strXML = strXML & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>OMIGACUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>CIFNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>STAGECHANGEDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>TYPEOFLOAN<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>TYPEOFLOANCODE<TYPE>dbdtInt</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>APPLICATIONSTAGEDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>APPLICATIONSTAGENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>CONTACTREASON<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>CONTACTDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>CONTACTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>OTHERREASON<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "</TABLENAME>"
'
'    Set IomCustClassDef_LoadApplicationContactData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
'
'    Set objXmlAssist = Nothing
'
'End Function
''BMIDS00005 End

'BMIDS00026 IWP2 BM060
Private Function IomCustClassDef_LoadApplicationCustomerData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindCustomersByApplication view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadApplicationCustomerData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>FINDCUSTOMERSPERAPPLICATION"
    strXML = strXML & "<OTHERS>OMIGACUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CIFNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NEWCUSTOMERINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERROLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFAPPLICANTS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLECODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INITIALS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FORENAME1<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FORENAME2<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00205 Address details moved to LoadCRSCustomerAddressData
    strXML = strXML & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MARITALSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MARITALSTATUSCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MEMBEROFSTAFF<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMPLOYEENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GENDER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GENDERCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALINSURANCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALALLOWABLEINCOME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OCCUPATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OCCUPATIONTYPECODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>JOBTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMPLOYMENTSTARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAILSHOTREQUIRED<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALITY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALITYCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadApplicationCustomerData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00026 End

'BMIDS00026 IWP2 BM060
Private Function IomCustClassDef_LoadCustomerPhoneData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindPhoneNumbersPerCustomer view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadCustomerPhoneData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>FINDPHONENUMBERSPERCUSTOMER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGECODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadCustomerPhoneData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00026 End

'BMIDS00026 IWP2 BM060
Private Function IomCustClassDef_LoadCustomerVerificationData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindVerificationPerCustomer view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadCustomerVerificationData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    'BMIDS864 ApplicationNumber is not a primary key
    
    strXML = "<TABLENAME>FINDVERIFICATIONPERCUSTOMER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VERIFICATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IDENTIFICATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IDENTIFICATIONTYPECODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERIDENTIFICATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadCustomerVerificationData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00026 End

Private Function IomCustClassDef_LoadCRSCustomerAddressData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindAddressesPerCustomer view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadCRSCustomerAddressData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>FINDADDRESSESPERCUSTOMER"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAFDELIVERYPOINTSUFFIX<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATEMOVEDIN<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadCRSCustomerAddressData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00443
Private Function IomCustClassDef_LoadLatestCustomerVersionData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadLatestCustomerVersionData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>CUSTOMERVERSION"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadLatestCustomerVersionData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing

End Function
'BMIDS00443 End

'BMIDS758  New function
Private Function IomCustClassDef_LoadRemovedToECustomer() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCustClassDef_LoadRemovedToECustomer"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
     
    strXML = strXML & "<TABLENAME>REMOVEDTOECUSTOMER"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>OMIGACUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"
    
    Set IomCustClassDef_LoadRemovedToECustomer = _
                    objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

'BMIDS864  Add function
Private Function IomCustClassDef_LoadVerificationData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from a recordset - FindVerificationPerCustomer view
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadVerificationData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>VERIFICATION"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>VERIFICATIONSEQUENCENUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>VERIFICATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IDENTIFICATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERIDENTIFICATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ISSUER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VERIFICATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadVerificationData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS864 End
'PSC 28/06/2005 MAR5 - Start
Private Function IomCustClassDef_LoadCustomerAddressAndRoleData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - FindCustomerNameAndAddress View
'Pass        : None
'Return      : Dom Document (Table Definition)
' **************************************************************************
    Dim strFunctionName As String
    strFunctionName = "IomCustClassDef_LoadCustomerAddressAndRoleData"
    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMERADDRESS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERADDRESSSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CustomerAddressType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERROLETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CustomerRoleType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set IomCustClassDef_LoadCustomerAddressAndRoleData = xmlClassDefDoc
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Function IomCustClassDef_LoadCustomerNameData() As FreeThreadedDOMDocument40
'Header--------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - FindCustomerNameAndAddress View
'Pass        : None
'Return      : Dom Document (Table Definition)
' **************************************************************************
    Dim strFunctionName As String
    strFunctionName = "IomCustClassDef_LoadCustomerNameData"
    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CUSTOMERNAME"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CORRESPONDENCESALUTATION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERFORENAMES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set IomCustClassDef_LoadCustomerNameData = xmlClassDefDoc
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function
'PSC 28/06/2005 MAR5 - End


Private Function IomCustClassDef_LoadWrapUpDetailsData() As MSXML2.IXMLDOMDocument2
'------------------------------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset - CustomerWrapUpDetails Table
'Pass        : None
'Return      : Dom Document (Table Definition)
' *****************************************************************************
    Const strFunctionName As String = "IomCustClassDef_LoadWrapUpDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist

    strXML = strXML & "<TABLENAME>CUSTOMERWRAPUPDETAILS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>WRAPUPSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTIME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PREFERREDMETHODOFCONTACT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set IomCustClassDef_LoadWrapUpDetailsData = xmlClassDefDoc

    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

'MAR23 New function for KnowYourCustomerCheck table

Private Function IomCustClassDef_LoadKYCData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates the class template required for getting the XML from
'               a recordset - KnowYourCustomerCheck table
' pass:         None
' return:       XML string
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCustClassDef_LoadKYCData"
    
    Dim strXML          As String
    Dim objXmlAssist    As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = "<TABLENAME>KNOWYOURCUSTOMERCHECK"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>SEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SUCCESSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCustClassDef_LoadKYCData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

