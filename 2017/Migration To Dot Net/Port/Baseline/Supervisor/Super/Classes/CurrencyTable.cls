VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : CurrencyTable
' Description   : Contains functions associated with database access. This class
'                 maintains a list of all currently defined databases and manages
'                 switches between them
' Who?  When?       What?
' CL    24/04/01    Initial creation of screen
' DJP   26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sQuestionReference As String
Private Const m_sComboGroup = "ConditionType"
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "CURRENCY"

Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "CURRENCYCODE"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
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
    'TableAccess_GetSearch = m_sSearch
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

    colListLine.Add GetCurrencyCode
    colListLine.Add GetCurrencyName
    colListLine.Add GetConversionRate
    colListLine.Add GetRoundingPrecision
    colListLine.Add GetDirectionText
    colListLine.Add GetBooleanFromNumber(GetBaseCurrencyInd)
            
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
Public Function GetCurrencyCode() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CURRENCYCODE")
    GetCurrencyCode = CStr(vVal)
End Function
Public Function GetBaseCurrencyInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BASECURRENCYIND")
    GetBaseCurrencyInd = CStr(vVal)
End Function
Public Function GetCurrencyName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CURRENCYNAME")
    GetCurrencyName = CStr(vVal)
End Function
Public Function GetConversionRate() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONVERSIONRATE")
    GetConversionRate = CStr(vVal)
End Function
Public Function GetRoundingPrecision() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ROUNDINGPRECISION")
    GetRoundingPrecision = CStr(vVal)
End Function
Public Function GetRoundingDirection() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ROUNDINGDIRECTION")
    GetRoundingDirection = CStr(vVal)
End Function
Friend Function SetCurrencyCode(sName As String) As Boolean
    SetCurrencyCode = m_clsTableAccess.SetVal("CURRENCYCODE", CVar(sName))
End Function
Friend Function SetBaseCurrency(sName As String) As Boolean
    SetBaseCurrency = m_clsTableAccess.SetVal("BASECURRENCYIND", CVar(sName))
End Function
Friend Function SetCurrencyName(sName As String) As Boolean
    SetCurrencyName = m_clsTableAccess.SetVal("CURRENCYNAME", CVar(sName))
End Function
Friend Function SetConversionRate(sName As String) As Boolean
    SetConversionRate = m_clsTableAccess.SetVal("CONVERSIONRATE", CVar(sName))
End Function
Friend Function SetRoundingPrecision(sName As String) As Boolean
    SetRoundingPrecision = m_clsTableAccess.SetVal("ROUNDINGPRECISION", CVar(sName))
End Function
Friend Function SetRoundingDirection(sName As String) As Boolean
    SetRoundingDirection = m_clsTableAccess.SetVal("ROUNDINGDIRECTION", CVar(sName))
End Function
Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub
Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
Public Sub GetCurrencies()
On Error GoTo Failed
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim sComboPrecision As String
    Dim sComboDirection As String
    
    sComboPrecision = "RoundingPrecision"
    sComboDirection = "RoundingDirection"
    
    On Error GoTo Failed
        sSQL = "SELECT " & m_sTableName & ".*,  ComboDirection.valueName DirectionText" & _
           " FROM comboValue comboDirection ," & m_sTableName & _
           " WHERE" & _
           " comboDirection.groupname = " & g_clsSQLAssistSP.FormatString(sComboDirection) & _
           " AND ComboDirection.valueID = RoundingDirection"
        
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    ValidateRecordset rs, "Currency"
    m_clsTableAccess.SetRecordSet rs
            
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Sub
Private Function GetDirectionText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("DIRECTIONTEXT")
    GetDirectionText = CStr(vVal)
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckBaseCurrency
' Description   : Returns a boolean indicating whether the base currency has been set
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AA 04/05/01
Public Function CheckBaseCurrency(sCurrencyCode As String) As Boolean
    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim bRet As Boolean
    
    bRet = True
    Set rs = New ADODB.Recordset
    
    sSQL = "Select CURRENCYNAME from " & m_sTableName & " Where BASECURRENCYIND = 1 And CURRENCYCODE <> " & g_clsSQLAssistSP.FormatString(sCurrencyCode)

    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    If rs.RecordCount > 0 Then
        g_clsErrorHandling.DisplayError "The base currency is currently set to " & rs.fields("CURRENCYNAME").Value
        bRet = False
    End If
    
    CheckBaseCurrency = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
