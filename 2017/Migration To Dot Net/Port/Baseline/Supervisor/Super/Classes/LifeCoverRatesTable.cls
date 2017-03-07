VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LifeCoverRatesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LifeCoverRatesTable
' Description   : Contains all access routines and SQL for the LifeCoverRates table. Implements
'                 TableAccess.
' Change history
' Prog      Date        Description
' DJP       11/06/01    SQL Server port.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'                       Added NextCoverRate to create the next Life Cover rate.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "LIFECOVERRATES"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "LifeCoverRatesNumber"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetLifeCoverRateNumber() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("LIFECOVERRATESNUMBER")
    GetLifeCoverRateNumber = vVal
End Function
Friend Function GetStartDate() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetDate("LIFECOVERRATESSTARTDATE")
    
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If

End Function
Friend Function GetCoverType() As Variant
    GetCoverType = m_clsTableAccess.GetVal("CoverType")
End Function
Friend Function GetCoverTypeText() As Variant
    GetCoverTypeText = m_clsTableAccess.GetVal("CoverTypeText")
End Function
Friend Function GetApplicantGender() As Variant
    GetApplicantGender = m_clsTableAccess.GetVal("ApplicantsGender")
End Function
Friend Function GetApplicantGenderText() As Variant
    GetApplicantGenderText = m_clsTableAccess.GetVal("ApplicantsGenderText")
End Function
Friend Function GetApplicantAgeMax() As Variant
    GetApplicantAgeMax = m_clsTableAccess.GetVal("HighApplicantsAge")
End Function
Friend Function GetTermMax() As Variant
    GetTermMax = m_clsTableAccess.GetVal("HighTerm")
End Function
Friend Function GetSmokerRate() As Variant
    GetSmokerRate = m_clsTableAccess.GetVal("AdditionalSmokerRate")
End Function
Friend Function GetAnnualRate() As Variant
    GetAnnualRate = m_clsTableAccess.GetVal("AnnualRate")
End Function
Friend Function GetNotGoodHealthRate() As Variant
    GetNotGoodHealthRate = m_clsTableAccess.GetVal("NotGoodHealthRate")
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetLifeCoverRateNumber(sVal As Variant) As Boolean
    SetLifeCoverRateNumber = m_clsTableAccess.SetVal("LIFECOVERRATESNUMBER", CVar(sVal))
End Function
Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("LIFECOVERRATESSTARTDATE", CVar(sVal))
End Function
Friend Function SetCoverType(sVal As Variant) As Boolean
    SetCoverType = m_clsTableAccess.SetVal("CoverType", CVar(sVal))
End Function
Friend Function SetApplicantGender(sVal As Variant) As Boolean
    SetApplicantGender = m_clsTableAccess.SetVal("ApplicantsGender", CVar(sVal))
End Function
Friend Function SetApplicantAgeMin(sVal As Variant) As Boolean
    SetApplicantAgeMin = m_clsTableAccess.SetVal("LowApplicantsAge", CVar(sVal))
End Function
Friend Function SetApplicantAgeMax(sVal As Variant) As Boolean
    SetApplicantAgeMax = m_clsTableAccess.SetVal("HighApplicantsAge", CVar(sVal))
End Function
Friend Function SetTermMin(sVal As Variant) As Boolean
    SetTermMin = m_clsTableAccess.SetVal("LowTerm", CVar(sVal))
End Function
Friend Function SetTermMax(sVal As Variant) As Boolean
    SetTermMax = m_clsTableAccess.SetVal("HighTerm", CVar(sVal))
End Function
Friend Function SetSmokerRate(sVal As Variant) As Boolean
    SetSmokerRate = m_clsTableAccess.SetVal("AdditionalSmokerRate", CVar(sVal))
End Function
Friend Function SetAnnualRate(sVal As Variant) As Boolean
    SetAnnualRate = m_clsTableAccess.SetVal("AnnualRate", CVar(sVal))
End Function
Friend Function SetNotGoodHealthRate(sVal As Variant) As Boolean
    SetNotGoodHealthRate = m_clsTableAccess.SetVal("NotGoodHealthRate", CVar(sVal))
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

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub
Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function

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
    colListLine.Add GetLifeCoverRateNumber()
    colListLine.Add GetStartDate()
    colListLine.Add GetCoverTypeText()
    colListLine.Add GetApplicantGenderText()
    colListLine.Add GetApplicantAgeMax()
    colListLine.Add GetTermMax()
    
    colListLine.Add GetAnnualRate()
    colListLine.Add GetSmokerRate()
    colListLine.Add GetNotGoodHealthRate()
End Sub
Friend Sub GetLifeCoverRates()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    sSearch = "select B.ValueName CoverTypeText, C.ValueName ApplicantsGenderText, A.LIFECOVERRATESNUMBER, A.LIFECOVERRATESSTARTDATE,  A.COVERTYPE, A.APPLICANTSGENDER, " + _
    "A.HIGHAPPLICANTSAGE, A.HIGHTERM, A.ADDITIONALSMOKERRATE, A.ANNUALRATE, A.NOTGOODHEALTHRATE " + _
    " from " + m_sTableName + " A, ComboValue B, ComboValue C where B.GroupName = 'LifeCoverType' and B.ValueID = A.COVERTYPE and " + _
    " C.GroupName = 'LifeCoverGender' and C.ValueID = A.APPLICANTSGENDER"

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, m_sTableName
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
Public Function NextCoverRate() As String
    On Error GoTo Failed
    
    NextCoverRate = g_clsDataAccess.NextNumber(m_sTableName, "LIFECOVERRATESNUMBER")
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

