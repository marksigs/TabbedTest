VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PayProtRatesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PayProtRates
' Description   : Contains all data access routines for the PaymentProtectionRates table.
'                 This table is a banded table which means for each row there are parent keys
'                 that need to be updated, hence implementing BandedTable
' Change history
' Prog      Date        Description
' DJP       26/06/0     SQL Server port
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Added
'                       NextRate method to return the next Rate number.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
Implements BandedTable

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable

' Constants
Private Const m_sTableName = "PAYMENTPROTECTIONRATES"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "PaymentProtectionRatesNumber"
    colFields.Add "PPStartDate"
    
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "PaymentProtectionRatesNumber"
    colUpdateFields.Add "PPStartDate"

    m_clsBandedTable.SetUpdateFields colUpdateFields

    Set colDupKeys = New Collection
    colDupKeys.Add "PaymentProtectionRatesNumber"
    colDupKeys.Add "PPStartDate"
    colDupKeys.Add "ApplicantsGender"
    colDupKeys.Add "HighApplicantsAge"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetRateNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PaymentProtectionRatesNumber")
    GetRateNumber = CStr(vVal)
End Function
Friend Function GetStartDate() As String
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("PPStartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetChannelID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ChannelID")
    GetChannelID = CStr(vVal)
End Function
Friend Function GetChannelName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ChannelName")
    GetChannelName = CStr(vVal)
End Function
Friend Function GetApplicantGender() As String
    GetApplicantGender = CStr(m_clsTableAccess.GetVal("ApplicantsGender"))
End Function
Friend Function GetHighApplicantAge() As String
    GetHighApplicantAge = CStr(m_clsTableAccess.GetVal("HighApplicantsAge"))
End Function
Friend Function GetASURate() As String
    GetASURate = CStr(m_clsTableAccess.GetVal("ASURate"))
End Function
Friend Function GetURate() As String
    GetURate = CStr(m_clsTableAccess.GetVal("URate"))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetRateNumber(vVal As Variant) As Boolean
    SetRateNumber = m_clsTableAccess.SetVal("PaymentProtectionRatesNumber", vVal)
End Function
Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("PPStartDate", CVar(sVal))
End Function
Friend Function SetChannelID(sVal As Variant) As Boolean
    SetChannelID = m_clsTableAccess.SetVal("ChannelID", CVar(sVal))
End Function
Friend Function SetApplicantGender(sVal As String) As Boolean
    SetApplicantGender = m_clsTableAccess.SetVal("ApplicantsGender", CVar(sVal))
End Function
'Friend Function SetLowApplicantAge(sVal As String) As Boolean
'    SetLowApplicantAge = m_clsTableAccess.SetVal("LowApplicantsAge", CVar(sVal))
'End Function
Friend Function SetHighApplicantAge(sVal As String) As Boolean
    SetHighApplicantAge = m_clsTableAccess.SetVal("HighApplicantsAge", CVar(sVal))
End Function
Friend Function SetASURate(sVal As String) As Boolean
    SetASURate = m_clsTableAccess.SetVal("ASURate", CVar(sVal))
End Function
Friend Function SetURate(sVal As String) As Boolean
    SetURate = m_clsTableAccess.SetVal("URate", CVar(sVal))
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
    colListLine.Add GetRateNumber()
    colListLine.Add GetStartDate()
    colListLine.Add GetChannelName()
    colListLine.Add GetApplicantGender()
    colListLine.Add GetHighApplicantAge()
    colListLine.Add GetASURate()
    colListLine.Add GetURate()
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Banded functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub
Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub
Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub
Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
End Sub
Public Sub GetPaymentProtectionRates()
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    sSearch = "select A.PAYMENTPROTECTIONRATESNUMBER,A.ChannelID, E.ChannelName,  A.PPSTARTDATE, A.ChannelID, COMBO.VALUENAME ""APPLICANTSGENDER"", A.HIGHAPPLICANTSAGE, A.ASURATE, A.ASRATE, A.URate from " + _
    "PaymentProtectionRates A,distributionchannel E, COMBOVALUE COMBO where A.ChannelID = " + _
    "(select MIN(ChannelID) from PaymentProtectionRates B " + _
    "where A.PPStartDate = B.PPStartDate and A.PaymentProtectionRatesNumber = B.PaymentProtectionRatesNumber ) and A.ApplicantsGender = " + _
    "(select MIN(ApplicantsGender) from PaymentProtectionRates C " + _
    "where A.PPStartDate = C.PPStartDate and A.PaymentProtectionRatesNumber = C.PaymentProtectionRatesNumber and A.ChannelID = C.ChannelID) and A.HighApplicantsAge = " + _
    "(select MIN(HighApplicantsAge) from PaymentProtectionRates D " + _
    "where A.PPStartDate = D.PPStartDate and A.PaymentProtectionRatesNumber = D.PaymentProtectionRatesNumber and A.ChannelID = D.ChannelID and A.ApplicantsGender = D.ApplicantsGender ) AND " + _
    "a.ChannelID = E.ChannelID AND " + _
    "UPPER(COMBO.GROUPNAME) = " + g_clsSQLAssistSP.FormatString("LIFECOVERGENDER") + " AND " + _
    " COMBO.VALUEID = A.APPLICANTSGENDER"

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    ValidateRecordset rs, m_clsTableAccess.GetTable()
    m_clsTableAccess.SetRecordSet rs

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetRates() As Boolean
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    sSearch = "select A.PAYMENTPROTECTIONRATESNUMBER from " + _
    "PaymentProtectionRates A,distributionchannel E where A.ChannelID = " + _
    "(select MIN(ChannelID) from PaymentProtectionRates B " + _
    "where A.PPStartDate = B.PPStartDate and A.PaymentProtectionRatesNumber = B.PaymentProtectionRatesNumber ) and A.ApplicantsGender = " + _
    "(select MIN(ApplicantsGender) from PaymentProtectionRates C " + _
    "where A.PPStartDate = C.PPStartDate and A.PaymentProtectionRatesNumber = C.PaymentProtectionRatesNumber and A.ChannelID = C.ChannelID) and A.HighApplicantsAge = " + _
    "(select MIN(HighApplicantsAge) from PaymentProtectionRates D " + _
    "where A.PPStartDate = D.PPStartDate and A.PaymentProtectionRatesNumber = D.PaymentProtectionRatesNumber and A.ChannelID = D.ChannelID and A.ApplicantsGender = D.ApplicantsGender ) AND " + _
    "a.ChannelID = E.ChannelID"
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    m_clsTableAccess.SetRecordSet rs

    If Not rs Is Nothing Then
        bRet = True
    Else
        bRet = False
    End If
End Function
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : NextRate
' Description   : Returns the next PaymentProtectionRatesNumber in the sequence.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NextRate() As String
    On Error GoTo Failed
    
    NextRate = g_clsDataAccess.NextNumber(m_sTableName, "PaymentProtectionRatesNumber")
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
