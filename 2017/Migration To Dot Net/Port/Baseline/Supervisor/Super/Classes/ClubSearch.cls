VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClubSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : ClubSearch
' Description   : Stores all search criteria and performs the search against the
'                 intermediary table
'
' Change history
' Prog      Date        Description
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Stores the search criteria specified in the gui
Private Type SearchDetails
    sClubId As String
    sDescription As String
    sBankAccountName As String
    sName As String
    sBankAccountNumber As String
    sBankAccountBranchName As String
    sBankSortCode As String
    sBankWizardIndicator As String
    sBuildingName As String
    sBuildingNumber As String
    sFlatNumber As String
    sStreet As String
    sDistrict As String
    sTown As String
    sCounty As String
    sCountry As String
    sPostcode As String
    sTelephoneCountryCode As String
    sTelephoneAreaCode As String
    sTelephoneNumber As String
    sFaxCountryCode As String
    sFaxAreaCode As String
    sFaxNumber As String
    sVolProcFeeAdjustment As String
    sProcLoadingInd As String
    sPackagerIndicator As String
    sLastUpdatedDate As String
    sLastUpdatedBy As String
    sPaymentType As String
    sAgreedProcRate As String
    sLastUpdateDate As String
    sLastUpdateBy As String
    sAgreedProcFeeRate As String
    sAddress As String
End Type

Private m_SearchDetails As SearchDetails

' Set
Public Sub SetName(ByVal sName As String)
    m_SearchDetails.sName = sName
End Sub
Public Sub SetAddress(ByVal sName As String)
    m_SearchDetails.sAddress = sName
End Sub

Public Sub SetDescription(ByVal sDescription As String)
    m_SearchDetails.sDescription = sDescription
End Sub
Public Sub SetBuildingName(ByVal sBuildingName As String)
    m_SearchDetails.sBuildingName = sBuildingName
End Sub
Public Sub SetBuildingNumber(ByVal sBuildingNumber As String)
    m_SearchDetails.sBuildingNumber = sBuildingNumber
End Sub
Public Sub SetFlatNumber(ByVal sFlatNumber As String)
    m_SearchDetails.sFlatNumber = sFlatNumber
End Sub
Public Sub SetStreet(ByVal sStreet As String)
    m_SearchDetails.sStreet = sStreet
End Sub
Public Sub SetDistrict(ByVal sDistrict As String)
    m_SearchDetails.sDistrict = sDistrict
End Sub
Public Sub SetTown(ByVal sTown As String)
    m_SearchDetails.sDistrict = sTown
End Sub
Public Sub SetPostcode(ByVal sPostcode As String)
    m_SearchDetails.sPostcode = sPostcode
End Sub
Public Sub SetTelephoneCountryCode(ByVal sTelephoneCountryCode As String)
    m_SearchDetails.sTelephoneCountryCode = sTelephoneCountryCode
End Sub
Public Sub SetTelephoneAreaCode(ByVal sTelephoneAreaCode As String)
    m_SearchDetails.sTelephoneAreaCode = sTelephoneAreaCode
End Sub
Public Sub SetTelephoneNumber(ByVal sTelephoneNumber As String)
    m_SearchDetails.sTelephoneNumber = sTelephoneNumber
End Sub
Public Sub SetFaxCountryCode(ByVal sFaxCountryCode As String)
    m_SearchDetails.sFaxCountryCode = sFaxCountryCode
End Sub
Public Sub SetFaxAreaCode(ByVal sFaxAreaCode As String)
    m_SearchDetails.sFaxAreaCode = sFaxAreaCode
End Sub
Public Sub SetFaxNumber(ByVal sFaxNumber As String)
    m_SearchDetails.sFaxNumber = sFaxNumber
End Sub
Public Sub SetVolProcFeeAdjustment(sVolProcFeeAdjustment As String)
    m_SearchDetails.sVolProcFeeAdjustment = sVolProcFeeAdjustment
End Sub
Public Sub SetProcLoadingInd(sProcLoadingInd As String)
    m_SearchDetails.sProcLoadingInd = sProcLoadingInd
End Sub
Public Sub SetPackagerIndicator(sPackagerIndicator As String)
    m_SearchDetails.sPackagerIndicator = sPackagerIndicator
End Sub
Public Sub SetLastUpdateDate(sLastUpdateDate As String)
    m_SearchDetails.sLastUpdateDate = sLastUpdateDate
End Sub
Public Sub SetLastUpdateBy(sLastUpdateBy As String)
    m_SearchDetails.sLastUpdateBy = sLastUpdateBy
End Sub
Public Sub SetPaymentType(sPaymentType As String)
    m_SearchDetails.sPaymentType = sPaymentType
End Sub
Public Sub SetAgreedProcFeeRate(sAgreedProcFeeRate As String)
    m_SearchDetails.sAgreedProcFeeRate = sAgreedProcFeeRate
End Sub

' Get
Public Function GetName() As String
    GetName = m_SearchDetails.sName
End Function
Public Function GetDescription() As String
    GetDescription = m_SearchDetails.sDescription
End Function
Public Function GetBuildingName() As String
    GetBuildingName = m_SearchDetails.sBuildingName
End Function
Public Function GetBuildingNumber() As String
    GetBuildingNumber = m_SearchDetails.sBuildingNumber
End Function
Public Function GetFlatNumber() As String
    GetFlatNumber = m_SearchDetails.sFlatNumber
End Function
Public Function GetStreet() As String
    GetStreet = m_SearchDetails.sStreet
End Function
Public Function GetDistrict() As String
    GetDistrict = m_SearchDetails.sDistrict
End Function
Public Function GetTown() As String
    GetTown = m_SearchDetails.sTown
End Function
Public Function GetPostCode() As String
    GetPostCode = m_SearchDetails.sPostcode
End Function
Public Function GetTelephoneCountryCode() As String
    GetTelephoneCountryCode = m_SearchDetails.sTelephoneCountryCode
End Function
Public Function GetTelephoneAreaCode() As String
    GetTelephoneAreaCode = m_SearchDetails.sTelephoneAreaCode
End Function
Public Function GetTelephoneNumber() As String
    GetTelephoneNumber = m_SearchDetails.sTelephoneNumber
End Function
Public Function GetFaxCountryCode() As String
    GetFaxCountryCode = m_SearchDetails.sFaxCountryCode
End Function
Public Function GetFaxAreaCode() As String
    GetFaxAreaCode = m_SearchDetails.sFaxAreaCode
End Function
Public Function GetFaxNumber() As String
    GetFaxNumber = m_SearchDetails.sFaxNumber
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Find
' Description   : Using the criteria set in SearchDetails, performs the search against the table
' TW        18/12/2006  Remove the use of CLUB view from Supervisor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Find(objClubTable As ClubTable)
    On Error GoTo Failed

    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim strWhereAnd As String
    Dim strCriteria As String
    
    strWhereAnd = " WHERE ISNULL(PACKAGERINDICATOR, 0) = 0 AND "
    
    'Name
    If Len(m_SearchDetails.sName) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " ISNULL(MORTGAGECLUBNETWORKASSOCNAME, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sName)
        strWhereAnd = " AND"
    End If
    
' TW 16/01/2007 EP2_859
    'Address
    If Len(m_SearchDetails.sAddress) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " (STREET LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR DISTRICT LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR TOWN LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR COUNTY LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR POSTCODE LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress) & ") "
        strWhereAnd = " AND"
    End If
    
'    'Description
'    If Len(m_SearchDetails.sDescription) > 0 Then
'        strCriteria = strCriteria & strWhereAnd & " ISNULL(MORTGAGECLUBNETWORKASSOCDESCRIPTION, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sDescription)
'        strWhereAnd = " AND"
'    End If
'
'    'BuildingName
'    If Len(m_SearchDetails.sBuildingName) > 0 Then
'        strCriteria = strCriteria & strWhereAnd & " ISNULL(BUILDINGORHOUSENAME, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sBuildingName)
'        strWhereAnd = " AND"
'    End If
'
'    'BuildingNumber
'    If Len(m_SearchDetails.sBuildingNumber) > 0 Then
'        strCriteria = strCriteria & strWhereAnd & " ISNULL(BUILDINGORHOUSENUMBER, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sBuildingNumber)
'        strWhereAnd = " AND"
'    End If
'
'    'Street
'    If Len(m_SearchDetails.sStreet) > 0 Then
'        strCriteria = strCriteria & strWhereAnd & " ISNULL(STREET, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sStreet)
'        strWhereAnd = " AND"
'    End If
'
'    'Postcode
'    If Len(m_SearchDetails.sPostcode) > 0 Then
'        strCriteria = strCriteria & strWhereAnd & " ISNULL(POSTCODE, '') LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sPostcode)
'        strWhereAnd = " AND"
'    End If
'
'    sSQL = "Select * " & _
'        "From MORTGAGECLUBNETWORKASSOCIATION" & _
'        strCriteria
    
    sSQL = "SELECT CLUBNETWORKASSOCIATIONID, MORTGAGECLUBNETWORKASSOCNAME, " & _
              "REPLACE(REPLACE(LTRIM(" & _
              "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
              "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
              "FROM MORTGAGECLUBNETWORKASSOCIATION" & _
        strCriteria
' TW 16/01/2007 EP2_859 End
    
    Set objRS = g_clsDataAccess.GetTableData(, sSQL)
    TableAccess(objClubTable).SetRecordSet objRS
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub







