VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllowableIncomeFactorTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : AllowableIncomeFactorTable
' Description   : Table class for the above table.
'
' Change history
' Prog  Date        Description
' STB   13-May-2002 SYS4417 Added AllowableIncomeFactor.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements TableAccess

'Constants.
Private Const m_sTableName = "ALLOWABLEINCOMEFACTORS"

'Underlying table object.
Private m_clsTableAccess As TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog  Date        Description
' SA    04/11/2002  BMIDS00646 Employment Status now part of primary key
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
        
    Dim colFields As Collection
    Dim colDupKeys As Collection
    
    On Error GoTo Failed
    
    Set colFields = New Collection
    Set colDupKeys = New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    'Primary keys.
    colFields.Add "ORGANISATIONID"
    colFields.Add "INCOMEGROUP"
    colFields.Add "TYPE"
    '++ BMIDS00646 SA 04/11/02 EmploymentStatus is part of primary key
    colFields.Add "EMPLOYMENTSTATUS"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    
    'Duplicate keys
    colDupKeys.Add "ORGANISATIONID"
    colDupKeys.Add "EMPLOYMENTSTATUS"
    colDupKeys.Add "INCOMEGROUP"
    colDupKeys.Add "TYPE"
    colDupKeys.Add "FACTOR"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    
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
    'Stub.
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
    'Stub.
End Function

Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function

Public Function TableAccess_GetUpdateValues() As Collection
    'Stub.
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)

    colListLine.Add GetLenderName()
    colListLine.Add GetEmploymentStatusText()
    colListLine.Add GetIncomeGroupText()
    colListLine.Add GetIncomeTypeText()
    colListLine.Add GetFactor()
    colListLine.Add GetLenderName() & ", " & GetEmploymentStatusText() & ", " & GetIncomeGroupText() & ", " & GetIncomeTypeText(), OBJECT_DESCRIPTION_KEY
    
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

Friend Function GetLenderName() As Variant
    GetLenderName = m_clsTableAccess.GetVal("LENDERNAME")
End Function

Friend Function GetOrganisationID() As Variant
    GetOrganisationID = m_clsTableAccess.GetVal("ORGANISATIONID")
End Function

Friend Function GetEmploymentStatus() As Variant
    GetEmploymentStatus = m_clsTableAccess.GetVal("EMPLOYMENTSTATUS")
End Function

Friend Function GetEmploymentStatusText() As Variant
    GetEmploymentStatusText = m_clsTableAccess.GetVal("EMPLOYMENTSTATUSTEXT")
End Function

Friend Function GetIncomeGroup() As Variant
    GetIncomeGroup = m_clsTableAccess.GetVal("INCOMEGROUP")
End Function

Friend Function GetIncomeGroupText() As Variant
    GetIncomeGroupText = m_clsTableAccess.GetVal("INCOMEGROUPTEXT")
End Function

Friend Function GetIncomeType() As Variant
    GetIncomeType = m_clsTableAccess.GetVal("TYPE")
End Function

Friend Function GetIncomeTypeText() As Variant
    GetIncomeTypeText = m_clsTableAccess.GetVal("TYPETEXT")
End Function

Friend Function GetFactor() As Variant
    GetFactor = m_clsTableAccess.GetVal("FACTOR")
End Function

Friend Function SetOrganisationID(ByVal vVal As Variant) As Boolean
    SetOrganisationID = m_clsTableAccess.SetVal("ORGANISATIONID", vVal)
End Function

Friend Function SetEmploymentStatus(ByVal vVal As Variant) As Boolean
    SetEmploymentStatus = m_clsTableAccess.SetVal("EMPLOYMENTSTATUS", vVal)
End Function

Friend Function SetIncomeGroup(ByVal vVal As Variant) As Boolean
    SetIncomeGroup = m_clsTableAccess.SetVal("INCOMEGROUP", vVal)
End Function

Friend Function SetIncomeType(ByVal vVal As Variant) As Boolean
    SetIncomeType = m_clsTableAccess.SetVal("TYPE", vVal)
End Function

Friend Function SetFactor(ByVal vVal As Variant) As Boolean
    SetFactor = m_clsTableAccess.SetVal("FACTOR", vVal)
End Function

Private Function CopyAndRenameFieldNames(ByRef colKeyMatchFields As Collection, ByVal sTableAlias As String) As Collection

    Dim iIndex As Integer
    Dim colNewKeyFields As Collection
    
    'Create a collection to hold modified key names in.
    Set colNewKeyFields = New Collection
    
    For iIndex = 1 To colKeyMatchFields.Count
        colNewKeyFields.Add sTableAlias & "." & colKeyMatchFields.Item(iIndex)
    Next iIndex
    
    Set CopyAndRenameFieldNames = colNewKeyFields

End Function

Friend Function GetDescription() As String

    Dim sSQL As String
    Dim sIncomeType As String
    Dim sIncomeGroup As String
    Dim sOrganisationID As String
    Dim sEmploymentStatus As String
    Dim rsResults As ADODB.Recordset
    
    'Obtain the criteria field values.
    sIncomeType = IIf(GetIncomeType() = "", "NULL", GetIncomeType())
    sIncomeGroup = IIf(GetIncomeGroup() = "", "NULL", GetIncomeGroup())
    sEmploymentStatus = IIf(GetEmploymentStatus() = "", "NULL", GetEmploymentStatus())
    sOrganisationID = "0x" & g_clsSQLAssistSP.GuidToString(GetOrganisationID())
    
    'Construct the SQL to obtain the displayable combovalues.
    sSQL = "SELECT  EMPLOYMENTSTATUSTEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'EmploymentStatus' AND VALUEID = " & sEmploymentStatus & "), " & _
           "        INCOMEGROUPTEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'IncomeGroupType' AND VALUEID = " & sIncomeGroup & "), " & _
           "        LENDERNAME = (SELECT LENDERNAME FROM MORTGAGELENDER WHERE ORGANISATIONID = " & sOrganisationID & ") "
               
    Select Case sIncomeGroup
        Case "1"
            sSQL = sSQL & ", TYPETEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'ConfirmedIncomeType' AND VALUEID = " & sIncomeType & ")"
            
        Case "2"
            sSQL = sSQL & ", TYPETEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'IncomeType' AND VALUEID = " & sIncomeType & ")"
            
        Case "3"
            sSQL = sSQL & ", TYPETEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'SelfEmpConfirmedIncomeType' AND VALUEID = " & sIncomeType & ")"
            
        Case "4"
            sSQL = sSQL & ", TYPETEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'SelfEmpDeclaredIncomeType' AND VALUEID = " & sIncomeType & ")"
            
        Case "5"
            sSQL = sSQL & ", TYPETEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'UnEarnedIncomeDescription' AND VALUEID = " & sIncomeType & ")"
    End Select
               
    'Execute the query.
    Set rsResults = g_clsDataAccess.ExecuteCommand(sSQL)

    'Lender,Emp,Group,Type
    GetDescription = (rsResults!LENDERNAME & ", " & rsResults!EMPLOYMENTSTATUSTEXT & ", " & rsResults!INCOMEGROUPTEXT & ", " & rsResults!TYPETEXT)

End Function

Public Function GetIncomeFactors() As ADODB.Recordset

    Dim sSearch As String
    Dim rsFactors As ADODB.Recordset
    Dim colKeyMatchFields As Collection
            
    'This SQL pulls back descriptive names for the income values.
    sSearch = "SELECT  ORGANISATIONID = AIF.ORGANISATIONID, " & _
              "        INCOMEGROUP = AIF.INCOMEGROUP, AIF.EMPLOYMENTSTATUS, " & _
              "        TYPE = AIF.TYPE, " & _
              "        EMPLOYMENTSTATUSTEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'EmploymentStatus' AND VALUEID = AIF.EMPLOYMENTSTATUS), " & _
              "        INCOMEGROUPTEXT = (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'IncomeGroupType' AND VALUEID = AIF.INCOMEGROUP), " & _
              "        TYPETEXT = " & _
              "        Case AIF.IncomeGroup " & _
              "             WHEN 1 THEN (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'ConfirmedIncomeType' AND VALUEID = AIF.TYPE) " & _
              "             WHEN 2 THEN (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'IncomeType' AND VALUEID = AIF.TYPE) " & _
              "             WHEN 3 THEN (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'SelfEmpConfirmedIncomeType' AND VALUEID = AIF.TYPE) " & _
              "             WHEN 4 THEN (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'SelfEmpDeclaredIncomeType' AND VALUEID = AIF.TYPE) " & _
              "             WHEN 5 THEN (SELECT VALUENAME FROM COMBOVALUE WHERE GROUPNAME = 'UnEarnedIncomeDescription' AND VALUEID = AIF.TYPE) " & _
              "        END, " & _
              "        FACTOR = AIF.FACTOR, " & _
              "        ML.LENDERNAME " & _
              "FROM    ALLOWABLEINCOMEFACTORS AIF INNER JOIN " & _
              "        MORTGAGELENDER ML ON ML.ORGANISATIONID = AIF.ORGANISATIONID"
              
    'If the keys haven't been specified yet, bring-back all records.
    If Not m_clsTableAccess.GetKeyMatchValues Is Nothing Then
        'The field names must be renamed to include the AIF. prefix.
        Set colKeyMatchFields = CopyAndRenameFieldNames(m_clsTableAccess.GetKeyMatchFields, "AIF")
        
        'Append the generated WHERE clause.
        sSearch = sSearch & " WHERE " & g_clsDataAccess.GenerateSearchCriteria(colKeyMatchFields, m_clsTableAccess.GetKeyMatchValues)
    End If
    
    'Get the query results into a recordset.
    Set rsFactors = g_clsDataAccess.GetTableData(, sSearch)

    'Ensure a recordset was returned.
    ValidateRecordset rsFactors, "IncomeFactors"

    'Associcate this recordset with the underlying table object.
    m_clsTableAccess.SetRecordSet rsFactors

    Set GetIncomeFactors = rsFactors

End Function

