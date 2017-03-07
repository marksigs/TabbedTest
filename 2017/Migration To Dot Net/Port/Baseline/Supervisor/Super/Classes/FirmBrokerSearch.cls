VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmBrokerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : FirmBrokerSearch
' Description   : Stores all search criteria and performs the search against the
'                 intermediary table
'
' Change history
' Prog      Date        Description

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Stores the search criteria specified in the gui
Private Type SearchDetails
    sFirmName As String
    sFSARef As String
    sAddress As String
End Type

Private m_SearchDetails As SearchDetails

' Set
Public Sub SetFirmName(ByVal sName As String)
    m_SearchDetails.sFirmName = sName
End Sub
Public Sub SetFSARef(ByVal sName As String)
    m_SearchDetails.sFSARef = sName
End Sub
Public Sub SetAddress(ByVal sAddress As String)
    m_SearchDetails.sAddress = sAddress
End Sub

' Get
Public Function GetFirmName() As String
    GetFirmName = m_SearchDetails.sFirmName
End Function
Public Function GetFSARef() As String
    GetFSARef = m_SearchDetails.sFSARef
End Function
Public Function GetAddress() As String
    GetAddress = m_SearchDetails.sAddress
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Find
' Description   : Using the criteria set in SearchDetails, performs the search against the table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Find(objFirmBrokerTable As FirmBrokerTable)
    On Error GoTo Failed

    Dim sSQL As String
    Dim sCompanyNameSearch As String
    Dim sForenameSearch As String
    Dim sSurnameSearch As String
    Dim sTownSearch As String
    Dim sIntTypeSearch As String
    Dim rs As ADODB.Recordset
    Dim strWhereAnd As String
    Dim strCriteria As String
    
    strWhereAnd = " WHERE"
    
    'Firmname
    If Len(m_SearchDetails.sFirmName) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " NAME LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFirmName)
        strWhereAnd = " AND"
    End If
    
    'FSARef
    If Len(m_SearchDetails.sFSARef) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " FSARef LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFSARef)
        strWhereAnd = " AND"
    End If
    
    'Address
    If Len(m_SearchDetails.sAddress) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " (ADDRESSLINE1 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE2 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE3 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE4 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE5 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE6 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR POSTCODE LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress) & ") "
        strWhereAnd = " AND"
    End If
    
    sSQL = "Select * " & _
        "From FirmBroker" & _
        strCriteria
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    TableAccess(objFirmBrokerTable).SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



