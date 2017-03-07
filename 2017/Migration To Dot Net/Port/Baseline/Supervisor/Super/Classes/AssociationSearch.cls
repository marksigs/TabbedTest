VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssociationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : AssociationSearch
' Description   : Stores all search criteria and performs the search against the
'                 intermediary table
'
' Change history
' Prog      Date        Description
' TW        18/12/2006  Remove the use of ASSOCIATION view from Supervisor
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Stores the search criteria specified in the gui
Private Type SearchDetails
    sAssociationId As String
    sName As String
    sDescription As String
    sBuildingName As String
    sBuildingNumber As String
    sStreet As String
    sPostcode As String
    sAddress As String
End Type

Private m_SearchDetails As SearchDetails

' Set
Public Sub SetAddress(ByVal sName As String)
    m_SearchDetails.sAddress = sName
End Sub
Public Sub SetName(ByVal sName As String)
    m_SearchDetails.sName = sName
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
Public Sub SetStreet(ByVal sStreet As String)
    m_SearchDetails.sStreet = sStreet
End Sub
Public Sub SetPostcode(ByVal sPostcode As String)
    m_SearchDetails.sPostcode = sPostcode
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
Public Function GetStreet() As String
    GetStreet = m_SearchDetails.sStreet
End Function
Public Function GetPostCode() As String
    GetPostCode = m_SearchDetails.sPostcode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Find
' Description   : Using the criteria set in SearchDetails, performs the search against the table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Find(objAssociationTable As AssociationTable)
    On Error GoTo Failed

    Dim sSQL As String
    Dim objRS As ADODB.Recordset
    Dim strWhereAnd As String
    Dim strCriteria As String
    
    strWhereAnd = " WHERE PACKAGERINDICATOR = 1 AND"
    
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
    TableAccess(objAssociationTable).SetRecordSet objRS
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub





