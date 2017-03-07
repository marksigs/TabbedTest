VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : RateTable
' Description   : Contains all data access functions for the BaseRate Table
' Change history
' Prog      Date        Description
' AA        26/02/01    Added class
' DJP       22/06/01    Use global Supervisor SQLAssist
' SA        17/12/01    SYS3504 New functions
' STB       08/03/02    SYS4251 - Applied Date is not a double.
' STB       20/03/02    SYS4300 - Added the IsRateIDUnique() and GetLatestRate()
'                       methods.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        18/05/2007  VR262 Show Description field on Base Rate in Supervisor main form
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName = "BaseRate"
Private Const m_sRateTypeGroup = "RateType"

Implements TableAccess

Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    On Error GoTo Failed
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "RATEID"
    colFields.Add "BASERATESTARTDATE"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    m_clsTableAccess.SetDuplicateKeys colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

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

Private Function TableAccess_SetDate(sField As String, vval As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vval)
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
Private Function TableAccess_SetVal(sField As Variant, vval As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vval)
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

    colListLine.Add GetRateID
' TW 18/05/2007 VR262
    colListLine.Add GetDescription
' TW 18/05/2007 VR262 End
    colListLine.Add GetStartDate
    colListLine.Add GetInterestRate
    colListLine.Add GetRateTypeText
    colListLine.Add GetRateAppliedDate
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

Friend Function GetRateID() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("RATEID")
    GetRateID = CStr(vval)
End Function
Friend Function GetStartDate() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("BASERATESTARTDATE")
    GetStartDate = CStr(vval)
End Function
Friend Function GetInterestRate() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("BASEINTERESTRATE")
    GetInterestRate = CStr(vval)
End Function
Friend Function GetDescription() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("RATEDESCRIPTION")
    GetDescription = CStr(vval)
End Function
Friend Function GetRateType() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("RATETYPE")
    GetRateType = CStr(vval)
End Function
Friend Function GetRateTypeText() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("RATETYPETEXT")
    GetRateTypeText = CStr(vval)
End Function
Friend Function GetRateAppliedDate() As String
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("BASERATEAPPLIEDDATE")
    GetRateAppliedDate = CStr(vval)
End Function
Friend Function SetRateID(sRateID As String) As Boolean
    SetRateID = m_clsTableAccess.SetVal("RATEID", CVar(sRateID))
End Function
Friend Function SetStartDate(dStart As Date) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("BASERATESTARTDATE", CVar(dStart))
End Function
Friend Function SetInterestRate(sRate As String) As Boolean
    SetInterestRate = m_clsTableAccess.SetVal("BASEINTERESTRATE", CVar(sRate))
End Function
Friend Function SetRateType(vType As Variant) As Boolean
    SetRateType = m_clsTableAccess.SetVal("RATETYPE", vType)
End Function
Friend Function SetDecsription(vval As Variant) As Boolean
    SetDecsription = m_clsTableAccess.SetVal("RATEDESCRIPTION", vval)
End Function
Friend Function SetRateAppliedDate(vval As Variant) As Boolean
    SetRateAppliedDate = m_clsTableAccess.SetVal("BASERATEAPPLIEDDATE", vval)
End Function
' DJP Added global SQL Assist as part of SQL Server port
Public Function GetBaseRate()
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
    sSQL = "SELECT " & m_sTableName & ".*, ComboRateType.ValueName RateTypeText" & _
           " FROM combovalue ComboRateType, " & m_sTableName & _
           " WHERE" & _
           " ComboRateType.GroupName = " & g_clsSQLAssistSP.FormatString(m_sRateTypeGroup) & " AND" & _
           " ComboRateType.ValueID = RateType"
           
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    ValidateRecordset rs, "BaseRate"
    m_clsTableAccess.SetRecordSet rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetListField() As String
    GetListField = "BASERATESET"
End Function
' SYS3504
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetBaseRateData
' Description   : Returns the RateId and BaseInterestRate from the BaseRate table in a recordset.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetBaseRateData() As ADODB.Recordset
    On Error GoTo Failed
    Dim sTable As String
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    sTable = "BASERATE"
    
    sSearch = "SELECT RateId, BaseInterestRate FROM " + sTable + " ORDER BY RateId"

    Set rs = g_clsDataAccess.GetTableData(sTable, sSearch)
    m_clsTableAccess.SetRecordSet rs
    
    Set GetBaseRateData = rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'SYS3504
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetRateIdField
' Description   : Returns the field name used to retrieve the RateId.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRateIdField() As String
    GetRateIdField = "RateId"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsRateIDUnique
' Description   : Returns true if the Rate ID already exists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsRateIDUnique(ByVal sRateID As String) As Boolean

    Dim sSearch As String
    Dim rsRates As ADODB.Recordset

    'Build a query to return something if a matching ID exists.
    sSearch = "SELECT 1 FROM BASERATE WHERE RATEID = " & g_clsSQLAssistSP.FormatString(sRateID)
    
    'Obtain a recordset of results.
    Set rsRates = g_clsDataAccess.GetTableData(, sSearch)

    'Return true if the recordset is empty.
    IsRateIDUnique = (rsRates.BOF And rsRates.EOF)
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLatestRate
' Description   : Populates the underlying recordset with data for the latest
'                 Rate record with a mathcing RateID. If no data is found then
'                 FALSE is returned.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLatestRate(ByVal sRateID As String) As Boolean

    Dim sSearch As String

    'Build a query to return something if a matching ID exists.
    sSearch = "SELECT * FROM BASERATE WHERE RATEID = " & g_clsSQLAssistSP.FormatString(sRateID)
    sSearch = sSearch & " AND BASERATESTARTDATE = (SELECT MAX(BASERATESTARTDATE) FROM BASERATE BR1 WHERE BR1.RATEID = BASERATE.RATEID)"
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    'Return true if the recordset is not empty.
    GetLatestRate = (m_clsTableAccess.RecordCount > 0)

End Function
