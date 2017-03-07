VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BandedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BandedTable
' Description   : Updates every record with the common fields retrieved from GetUpdateFields.
'                 This is used in "Sets" where there is one table, but that one table defines
'                 sets and also bands. The set is maybe the set name and date (which occurs in
'                 every record) and the rest of each record is the band.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private m_colUpdateValues As Collection
Private m_colUpdateFields As Collection

Private m_sName As String
Private m_sStartDate As String
Private m_sDescription As String

Private m_clsTableAccess As TableAccess
Private m_bUpdateSets  As Boolean
Public Sub SetTableAccess(clsTableAccess As TableAccess)
    Set m_clsTableAccess = clsTableAccess
End Sub

Public Sub SetUpdateValues(colValues As Collection)
    Set m_colUpdateValues = colValues
End Sub
Public Sub SetUpdateFields(colFields As Collection)
    Set m_colUpdateFields = colFields
End Sub
'Public Function GetUpdateFields() As Collection
'
'End Function
'Public Function GetUpdateValues() As Collection
'
'End Function
Public Sub SetUpdateSets(Optional bUpdate As Boolean = True)
    m_bUpdateSets = bUpdate
End Sub
Public Sub DoUpdateSets()
    On Error GoTo Failed
    Dim sField As String
    ' DJP SQL Server port
    Dim sValue As Variant
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim rs As ADODB.Recordset
    Dim nRSCnt As Integer
    
    If m_bUpdateSets = True Then
        Set rs = m_clsTableAccess.GetRecordSet()
        
        m_clsTableAccess.ValidateData
        If m_colUpdateFields.Count = m_colUpdateValues.Count Then
            nCount = m_colUpdateFields.Count
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                ' A for loop is used to navigate through the record set, as VB does not
                ' correctly set the .EOF value to TRUE when the last record of the recordset is hit.
                
                For nRSCnt = 1 To rs.RecordCount
                    For nThisItem = 1 To nCount
                        sField = m_colUpdateFields(nThisItem)
                        ' Convert this if required
                        'sValue = g_clsDataAccess.ConvertDatabaseItem(sValue)
                        
                        sValue = m_colUpdateValues(nThisItem)
                        rs(sField) = sValue
                    Next
                    rs.MoveNext
                Next nRSCnt
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "DoUpdate: Fields and update values don't match"
        End If
        
        m_bUpdateSets = False
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Class_Initialize()
    m_bUpdateSets = False
End Sub
Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function
Private Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
