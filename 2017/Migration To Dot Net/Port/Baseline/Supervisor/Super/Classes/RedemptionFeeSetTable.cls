VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RedemptionFeeSetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : RedemptionFeeSetTable
' Description   : Performs all tableaccess functions on the REDEMPTIONFEESET table
' BMIDS SPECIFIC CLASS
' BMIDS Change history
'---------------------
' Prog      Date        Description
' GD        16/05/2002  Added Class BMIDS00011
' AW        21/05/02    Added DeleteIfNoBands - (BM017)
' AW        29/05/02    BMIDS00016  Ammendment toDeleteIfNoBands
' KRW       30/04/04    Redemption Fee Combo populated with None on Screen Entry
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "REDEMPTIONFEESET"
Private Const m_sVersionField As String = "RDSETVERSIONNUMBER"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetRedemFeeSet() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("RedemptionFeeSet")
    GetRedemFeeSet = CStr(vVal)
End Function
Friend Function GetRedemFeeStartDate() As String
'    Dim vVal As Variant
 '   vVal = m_clsTableAccess.GetVal("RedemptionFeeStartDate")
'    GetRedemFeeStartDate = CStr(vVal)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetRedemFeeSet(sRateSet As String) As Boolean
    SetRedemFeeSet = m_clsTableAccess.SetVal("RedemptionFeeSet", CVar(sRateSet))
    
    g_clsVersion.HandleVersioning Me
End Function

Friend Function SetRedemStartDate(sRateSet As String) As Boolean
  '  SetRedemStartDate = m_clsTableAccess.SetVal("RedemptionFeeStartDate", CVar(sRateSet))
    
  '  g_clsVersion.HandleVersioning Me
End Function

Private Sub Class_Initialize()

    On Error GoTo Failed

    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "RedemptionFeeSet"
    
    If g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False) Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Function GetRedemptionSets() As ADODB.Recordset
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
 
    sSearch = "SELECT " & "* " & "FROM " & m_sTableName

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    
    Set GetRedemptionSets = rs

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function GetRedemptionSetDesc(sKey As String) As ADODB.Recordset
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed

    sSearch = "SELECT " & "* " & "FROM " & m_sTableName & " WHERE RedemptionFeeSet = " & sKey

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    
    Set GetRedemptionSetDesc = rs

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function GetTheRedemptionFeeSet() As Variant

    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("RedemptionFeeset")
    GetTheRedemptionFeeSet = vVal
End Function

Friend Function SetTheRedemptionFeeSet(sVal As Variant) As Boolean

    SetTheRedemptionFeeSet = m_clsTableAccess.SetVal("RedemptionFeeSet", CVar(sVal))

    g_clsVersion.HandleVersioning Me
End Function

Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_clsTableAccess.GetIsDeleted()
End Function


Private Sub TableAccess_SetDeleteFields(colFields As Collection)

    m_clsTableAccess.SetDeleteFields colFields
End Sub



Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_clsTableAccess.GetTableDataFromSearch sSearch
End Sub

Private Function TableAccess_GetLength(sField As String) As Integer
   TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function



Private Sub TableAccess_AddRow()

    m_clsTableAccess.AddRow
End Sub


Private Sub TableAccess_CloseRecordSet()

    m_clsTableAccess.CloseRecordSet
End Sub

Private Sub TableAccess_DeleteAllRows()
    m_clsTableAccess.DeleteAllRows
End Sub



Private Sub TableAccess_DeleteRow(colMatchData As Collection)

    m_clsTableAccess.DeleteRow colMatchData
End Sub


Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean

    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function


Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant

    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function


Private Function TableAccess_GetKeyMatchFields()

    Set TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function



Private Function TableAccess_GetKeyMatchValues() As Variant

    Set TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function



Private Function TableAccess_GetRecordSet() As ADODB.Recordset

    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function



Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset

    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function

Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function

Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function

Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
        m_clsTableAccess.ValidateData sDesc
End Sub

Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function

Private Sub TableAccess_MoveFirst()
    m_clsTableAccess.MoveFirst
End Sub

Private Function TableAccess_MoveNext() As Boolean

    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
End Function



Private Function TableAccess_RecordCount() As Integer

    TableAccess_RecordCount = m_clsTableAccess.RecordCount()
End Function

Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean

    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vVal)
End Function


Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)

    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)

    m_clsTableAccess.SetKeyMatchValues colValues
End Sub


Private Function TableAccess_SetNextRow() As Boolean

    TableAccess_SetNextRow = m_clsTableAccess.SetNextRow()
End Function



Private Sub TableAccess_SetPopulateType(enumType As PopulateType)


End Sub


Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)

    m_clsTableAccess.SetRecordSet rs
End Sub



Private Sub TableAccess_SetTableClass(clsTable As TableAccess)

    m_clsTableAccess = clsTable
End Sub


Private Sub TableAccess_SetUpdated()

    m_clsTableAccess.SetUpdated
End Sub



Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean

    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vVal)
End Function



Private Sub TableAccess_Update()

    m_clsTableAccess.Update
End Sub

Private Sub TableAccess_SetPopulateNull()

    m_clsTableAccess.SetPopulateNull
End Sub

Private Function TableAccess_GetSearch() As String

    TableAccess_GetSearch = m_sSearch
End Function

Private Function TableAccess_GetTable() As String

    TableAccess_GetTable = m_sTableName
End Function

Public Function TableAccess_GetUpdateValues() As Collection

End Function


Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")

    m_clsTableAccess.Create sTable, sFirstMatch
End Sub


Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetTheRedemptionFeeSet()
    colHeader.Add GetDescription()
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
End Sub


Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")

    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub TableAccess_SetDatabase(sDatabaseKey As String)

    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub



Public Function TableAccess_GetDatabase() As String

    TableAccess_GetDatabase = m_clsTableAccess.GetDatabase()
End Function

Private Sub TableAccess_SetOrderCriteria(sCriteria As String)

    m_clsTableAccess.SetOrderCriteria sCriteria
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



Friend Function SetDescription(sDescription As String) As Boolean

    SetDescription = m_clsTableAccess.SetVal("REDEMPTIONFEESETDESCRIPTION", CVar(sDescription))
End Function

Friend Function GetDescription() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("REDEMPTIONFEESETDESCRIPTION")
    GetDescription = CStr(vVal)
End Function

Public Function GetRedemptionFeeSet() As Boolean
'GD BMIDS
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    'sSearch = "SELECT CONVERT(nvarchar,BASERATESET) + ' - ' + BASERATESETDESCRIPTION 'CONCATFIELD', BASERATESET,RATEID FROM BASERATESET order by CONVERT(numeric,BASERATESET)"
    'sSearch = "SELECT CONVERT(nvarchar,REDEMPTIONFEESET) + ' - ' + REDEMPTIONFEESETDESCRIPTION 'CONCATFIELD', REDEMPTIONFEESET FROM REDEMPTIONFEESET order by REDEMPTIONFEESET"
    
    'KRW BMIDS737
    sSearch = "SELECT '<None>' AS CONCATFIELD, NULL AS REDEMPTIONFEESET UNION ALL " & _
    "SELECT CONVERT(nvarchar,REDEMPTIONFEESET) + ' - ' + REDEMPTIONFEESETDESCRIPTION 'CONCATFIELD', REDEMPTIONFEESET FROM REDEMPTIONFEESET order by REDEMPTIONFEESET"
   
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    m_clsTableAccess.SetRecordSet rs

    If Not rs Is Nothing Then
        bRet = True
    Else
        bRet = False
    End If
End Function

Public Function GetRedemptionFeeSetTable() As TableAccess
'GD BMIDS
    On Error GoTo Failed
    Set GetRedemptionFeeSetTable = m_clsTableAccess
    
    Exit Function
Failed:    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DeleteIfNoBands
' Description   : Delete the RedemptionSet if all RedemptionSetBands have also been deleted.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteIfNoBands(ByVal sRedemptionSet As String, ByVal sConnection As String)

    Dim sSQL As String, sCurrConnection As String
    Dim rsRateBands As ADODB.Recordset
    Dim bConnectionChanged As Boolean
    
    bConnectionChanged = False

    If Len(sConnection) > 0 Then
        sCurrConnection = g_clsDataAccess.GetConnectionKey
        g_clsDataAccess.SetActiveConnection (sConnection)
        bConnectionChanged = True
    End If
    
    'Ascertain if there are any more MigRateBand records.
    sSQL = "SELECT REDEMPTIONFEESET FROM REDEMPTIONFEEBAND WHERE REDEMPTIONFEESET = " & sRedemptionSet
    Set rsRateBands = g_clsDataAccess.GetTableData(, sSQL)
    
    'If not then delete the base rate set.
    If (rsRateBands.BOF And rsRateBands.EOF) Then
        sSQL = "DELETE FROM REDEMPTIONFEESET WHERE REDEMPTIONFEESET = " & sRedemptionSet
        g_clsDataAccess.ExecuteCommand sSQL
    End If

    ' Re-establish  old connection
    If bConnectionChanged = True Then
        g_clsDataAccess.SetActiveConnection (sCurrConnection)
    End If
    
End Sub


