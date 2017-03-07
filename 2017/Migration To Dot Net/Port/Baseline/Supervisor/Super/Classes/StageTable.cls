VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StageTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   StageTable
' Description   :   Contains all access functions for the Stage table.
' Change history
' Prog      Date        Description
' DJP       10/11/00    Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "STAGE"

' Public enum's
Public Enum StageType
    enumExceptionStage = 1
    enumDefaultStage
    enumStageAll
    enumStageAllIncDelete
End Enum

Private m_sNextNumberField As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' DJP       27/02/03    BM0354 - return EXCEPTIONSTAGEINDICATOR in GetStagesForActivityso it can be checked against
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Next number field
    m_sNextNumberField = "STAGEID"
    
    ' Key fields
    colFields.Add m_sNextNumberField
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "STAGENAME"
End Sub

Friend Function GetStageID() As Variant
    GetStageID = CStr(m_clsTableAccess.GetVal("STAGEID"))
End Function

Friend Function GetActivityID() As Variant
    GetActivityID = CStr(m_clsTableAccess.GetVal("ACTIVITYID"))
End Function

Friend Function GetStageSequenceNumber() As Variant
    GetStageSequenceNumber = CStr(m_clsTableAccess.GetVal("STAGESEQUENCENO"))
End Function

Friend Function GetStageName() As Variant
    GetStageName = CStr(m_clsTableAccess.GetVal("STAGENAME"))
End Function

Friend Function GetRuleRef() As Variant
   GetRuleRef = CStr(m_clsTableAccess.GetVal("STAGERULEREFERENCE"))
End Function

Friend Function GetAuthorityLevelText() As Variant
   GetAuthorityLevelText = CStr(m_clsTableAccess.GetVal("USERAUTHORITYLEVELEXT"))
End Function

Friend Function GetAuthorityLevel() As Variant
   GetAuthorityLevel = CStr(m_clsTableAccess.GetVal("USERAUTHORITYLEVEL"))
End Function

Friend Function GetException() As Variant
   GetException = CStr(m_clsTableAccess.GetVal("EXCEPTIONSTAGEINDICATOR"))
End Function

Friend Function SetStageID(sStageID As Variant) As Boolean
    SetStageID = m_clsTableAccess.SetVal("STAGEID", CVar(sStageID))
End Function

Friend Function SetStageName(sStageName As Variant) As Boolean
    SetStageName = m_clsTableAccess.SetVal("STAGENAME", CVar(sStageName))
End Function

Friend Function SetRuleRef(sRuleRef As Variant) As Boolean
    SetRuleRef = m_clsTableAccess.SetVal("STAGERULEREFERENCE", CVar(sRuleRef))
End Function

Friend Function SetStageSequenceNumber(sStageSequenceNumber As Variant) As Boolean
    SetStageSequenceNumber = m_clsTableAccess.SetVal("STAGESEQUENCENO", CVar(sStageSequenceNumber))
End Function

Friend Function SetAuthorityLevel(sAuthorityLevel As Variant) As Boolean
    SetAuthorityLevel = m_clsTableAccess.SetVal("USERAUTHORITYLEVEL", CVar(sAuthorityLevel))
End Function

Friend Function SetActivityID(sActivityID As Variant) As Boolean
    SetActivityID = m_clsTableAccess.SetVal("ACTIVITYID", CVar(sActivityID))
End Function

Friend Function SetException(sException As Variant) As Boolean
    SetException = m_clsTableAccess.SetVal("EXCEPTIONSTAGEINDICATOR", CVar(sException))
End Function

Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function

Friend Function GetNextNumberField() As String
    GetNextNumberField = m_sNextNumberField
End Function

Public Sub DeleteStage()
    On Error GoTo Failed
    
    GetStages m_clsTableAccess.GetKeyMatchValues()
    
    If m_clsTableAccess.RecordCount() > 0 Then
        SetDeleteFlag True
        m_clsTableAccess.Update
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetStagesWithCombos()
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsStages As ADODB.Recordset
    
    sSearch = "SELECT " & _
        m_sTableName & ".*, C1.VALUENAME ""USERAUTHORITYLEVELEXT"" " & _
    "FROM " & _
        m_sTableName & ", COMBOVALUE C1 " & _
    "WHERE " & _
        "DELETEFLAG <> 1 AND USERAUTHORITYLEVEL = C1.VALUEID  AND C1.GROUPNAME = " & g_clsSQLAssistSP.FormatString("UserRole")
    
    Set rsStages = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rsStages

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetStages(Optional colStageList As Collection = Nothing, Optional bDelete As Boolean = True, Optional bCheckActivityAttached As Boolean = False, Optional enumStageType As StageType = enumStageAll)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsStages As ADODB.Recordset
    Dim sStageID As String
    Dim nCount As Long
    Dim nThisItem As Long
    Dim nDeleteFlag As Integer
    Dim nExceptionStage As Integer
    
    If bDelete Then
        nDeleteFlag = DELETE_FLAG_SET
    Else
        nDeleteFlag = DELETE_FLAG_NOT_SET
    End If
    
    sSearch = "SELECT * FROM " + m_sTableName + " WHERE DELETEFLAG <> " & nDeleteFlag
    
    If Not colStageList Is Nothing Then
        If colStageList.Count > 0 Then
            sSearch = sSearch + " AND ("
            
            nCount = colStageList.Count
            For nThisItem = 1 To nCount
                sStageID = colStageList(nThisItem)
                sStageID = g_clsSQLAssistSP.FormatString(sStageID)
                sSearch = sSearch + " STAGEID = " & sStageID
                            
                If nThisItem < nCount Then
                    sSearch = sSearch + " OR "
                End If
            
            Next
            sSearch = sSearch + ")"
        End If
    End If
    
    If bCheckActivityAttached Then
        sSearch = sSearch & " AND ACTIVITYID IS  NULL"
    End If
    
    If enumStageType <> enumStageAll Then
        If enumStageType = enumExceptionStage Then
            nExceptionStage = 1
        Else
            nExceptionStage = 0
        End If
        
        sSearch = sSearch & " AND EXCEPTIONSTAGEINDICATOR = " & nExceptionStage
    End If
    
    Set rsStages = g_clsDataAccess.GetTableData(m_sTableName, sSearch, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rsStages
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetStagesForActivity(ByVal sActivityID As String, Optional enumStageType As StageType = enumStageAll, Optional bSelectAll As Boolean = False)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsActivities As ADODB.Recordset
    Dim nExceptionStage As Integer
    
    If Len(sActivityID) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, " GetStagesForActivity: ActivityID is empty"
    End If
    
    sActivityID = g_clsSQLAssistSP.FormatString(sActivityID)
        
    ' DJP BM0354 - return EXCEPTIONSTAGEINDICATOR so it can be checked against
    If bSelectAll Then
        sSearch = "SELECT * "
    Else
        sSearch = "SELECT " & _
                    "STAGENAME, STAGEID,ACTIVITYID,STAGESEQUENCENO, EXCEPTIONSTAGEINDICATOR "
    End If
    sSearch = sSearch & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE " & _
                    "ACTIVITYID = " & sActivityID
            If enumStageType <> enumStageAllIncDelete Then
                sSearch = sSearch & " AND " & _
                "DELETEFLAG <> " & DELETE_FLAG_SET & " "
            End If
                    
            If enumStageType <> enumStageAll And enumStageType <> enumStageAllIncDelete Then
                If enumStageType = enumExceptionStage Then
                    nExceptionStage = 1
                Else
                    nExceptionStage = 0
                End If
                
                sSearch = sSearch & " AND EXCEPTIONSTAGEINDICATOR = " & nExceptionStage
            End If
                    
            sSearch = sSearch & " ORDER BY STAGESEQUENCENO"
    
    Set rsActivities = g_clsDataAccess.GetTableData(m_sTableName, sSearch, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rsActivities
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
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

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetStageID()
    colListLine.Add GetStageName()
    colListLine.Add GetAuthorityLevelText()
    colListLine.Add GetBooleanFromNumber(GetException())
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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

