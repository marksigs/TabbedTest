VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionControlTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess

Private m_clsTableAccess As TableAccess

Private Const m_sTableName = "MPVERSIONCONTROL"
Private m_sSearch As String
Private m_bExists As Boolean
Private Const m_sVersionField As String = "VERSIONNUMBER"
Private Const STATUS_LIVE As String = "L"
Private Const STATUS_SUPERCEEDED As String = "S"
Private Const STATUS_PENDING As String = "P"
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add m_sVersionField
    m_clsTableAccess.SetKeyMatchFields colFields
    
    CheckVersioning

    If m_bExists Then
        g_sVersionNumber = CLng(GetCurrentVersion)
    End If
End Sub
Public Sub CheckVersioning()
    On Error Resume Next
    m_bExists = g_clsDataAccess.DoesTableExist(m_clsTableAccess.GetTable(), , False)
    
    If Err.Number <> 0 Then
        m_bExists = False
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetVersionNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal(m_sVersionField)
    GetVersionNumber = CStr(vVal)
End Function
Friend Function GetOrganisationID() As String
    GetOrganisationID = m_clsTableAccess.GetVal("ORGANISATIONID")
End Function
Friend Function GetStatus() As String
    GetStatus = m_clsTableAccess.GetVal("VERSIONSTATUS")
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetVersionNumber(sVersionNumber As String) As Boolean
    SetVersionNumber = m_clsTableAccess.SetVal(m_sVersionField, CVar(sVersionNumber))
End Function
Friend Function SetStatus(sStatus As String) As Boolean
    SetStatus = m_clsTableAccess.SetVal("VERSIONSTATUS", CVar(sStatus))
End Function
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

'Private Function TableAccess_DoUpdateSets() As Boolean
'    TableAccess_DoUpdateSets = m_clsTableAccess.DoUpdateSets()
'End Function
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
' DJP This can be added later if required.
Public Function TableAccess_GetUpdateValues() As Collection
End Function
' DJP This can be added later if required.
'Public Function TableAccess_GetUpdateFields() As Collection
'End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
Public Function DoesVersioningExist() As Boolean
    DoesVersioningExist = m_bExists
End Function
Public Function GetCurrentVersion() As String
    Dim sCommand As String
    Dim rs As ADODB.Recordset
    Dim colMatchValues As New Collection
    Dim sVersion As String
    ' Remove the status of existing versions
    sCommand = "SELECT * FROM " + m_sTableName + " WHERE VERSIONSTATUS = " + g_clsSQLAssistSP.FormatString(STATUS_LIVE)
    
    m_clsTableAccess.GetTableDataFromSearch sCommand
    Set rs = m_clsTableAccess.GetRecordSet()
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            sVersion = GetVersionNumber()
        End If
    End If
    GetCurrentVersion = sVersion
End Function
Public Function GetVersionField()
    GetVersionField = m_sVersionField
End Function
Public Sub GetVersions()
    Dim sCommand As String
    Dim rs As ADODB.Recordset
    Dim colMatchValues As New Collection
    Dim sStatus As String
    On Error GoTo Failed

    sCommand = "SELECT " + m_sVersionField + " FROM " + m_sTableName
    m_clsTableAccess.GetTableDataFromSearch sCommand
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetActiveVersion(sVersion As String)
    Dim sCommand As String
    Dim rs As ADODB.Recordset
    Dim colMatchValues As New Collection
    Dim sStatus As String
    On Error GoTo Failed


    ' Remove the status of existing versions
    sCommand = "SELECT * FROM " + m_sTableName + " WHERE VERSIONSTATUS = " + g_clsSQLAssistSP.FormatString(STATUS_LIVE)
    
    m_clsTableAccess.GetTableDataFromSearch sCommand
    Set rs = m_clsTableAccess.GetRecordSet()
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            
            While rs.EOF = False
                sStatus = GetStatus()
                
                If UCase(sStatus) = STATUS_LIVE Then
                    SetStatus STATUS_SUPERCEEDED
                End If

                rs.MoveNext
            Wend
            
            rs.UpdateBatch
        End If
    End If
    m_clsTableAccess.CloseRecordSet
    
    colMatchValues.Add sVersion
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    
    Set rs = m_clsTableAccess.GetTableData()
    
    If m_clsTableAccess.RecordCount() > 0 Then
        rs.MoveFirst
        SetVersionNumber sVersion
    
        SetStatus STATUS_LIVE
        m_clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub CreateVersion(sNewVersion As String)
    Dim sCommand As String
    Dim rs As ADODB.Recordset
    Dim colMatchValues As New Collection
    Dim sStatus As String
    On Error GoTo Failed

    'g_clsDataAccess.BeginTrans

    colMatchValues.Add sNewVersion
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    
    Set rs = m_clsTableAccess.GetTableData()
    
    If m_clsTableAccess.RecordCount() = 0 Then
        m_clsTableAccess.AddRow
        SetVersionNumber sNewVersion
        SetStatus STATUS_PENDING
        
        m_clsTableAccess.Update
    
    End If
    m_clsTableAccess.CloseRecordSet
    g_sVersionNumber = sNewVersion
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub HandleVersioning(clsTableAccess As TableAccess)
    On Error GoTo Failed
    
    If m_bExists Then
        ' First, Lender
        If TypeOf clsTableAccess Is LegalFeeSetTable Then
            clsTableAccess.SetVal "FEESETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MIGRateSetsTable Then
            clsTableAccess.SetVal "MIGRATESETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is LegalFeesTable Then
            clsTableAccess.SetVal "FEESETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MIGRateTable Then
            clsTableAccess.SetVal "MIGRATESETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortLenderParamsTable Then
            clsTableAccess.SetVal "MORTGAGELENDERVERSIONNUMBER", g_sVersionNumber
        
        ElseIf TypeOf clsTableAccess Is MortgageLendersTable Then
            clsTableAccess.SetVal "MORTGAGELENDERVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "MIGRATESETVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "FEESETVERSIONNUMBER", g_sVersionNumber
        
        ' Now, Product data
        ElseIf TypeOf clsTableAccess Is MortgageProductTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "MORTGAGELENDERVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "VFSETVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "AFSETVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "BRSETVERSIONNUMBER", g_sVersionNumber
        
        'BMIDS959  Remove MortgageProductBandsTable
        'ElseIf TypeOf clsTableAccess Is MortgageProductBandsTable Then
        '   clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is ValuationFeeSetTable Then
            clsTableAccess.SetVal "VFSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is ValuationFeeTable Then
            clsTableAccess.SetVal "VFSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is AdminFeeSetTable Then
            clsTableAccess.SetVal "AFSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is AdminFeeTable Then
            clsTableAccess.SetVal "AFSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is BaseRateSetTable Then
            clsTableAccess.SetVal "BRSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is BaseRateTable Then
            clsTableAccess.SetVal "BRSETVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdTypeofAppEligTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdPurpOfLoanTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdParamsTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdIntRatesTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdIncentiveTable Then
            clsTableAccess.SetVal "INCENTIVEVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdIncIncentiveTable Then
            clsTableAccess.SetVal "INCENTIVEVERSIONNUMBER", g_sVersionNumber
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        
        ElseIf TypeOf clsTableAccess Is MortProdLanguageTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdTypeOfBuyerTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdEmpEligTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        ElseIf TypeOf clsTableAccess Is MortProdPropLocTable Then
            clsTableAccess.SetVal "PRODUCTVERSIONNUMBER", g_sVersionNumber
        End If
    
    End If
    
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
