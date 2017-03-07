VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupervisorUpdateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : SupervisorUpdateTable
' Description   : Contains all code to manage access to the SupervisorUpdate table.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' STB       04/01/02    SYS3581 - Promotion fixes.
' DB        07/01/03    SYS5457 - Modify the promotion of Intermediaries
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' BS        01/05/03    BM0498 Add sort to results in FindTableUpdates
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        23/03/2007  EP2_1942 - Promotion of Income Factors does not work
' SR        21/11/2007  CORE00000313(VR 37)/MAR1968 - Allow UserID to be empty (when using integrated security)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements TableAccess

' Private data
Private m_sSearch As String
Private m_bOperationExists As Boolean
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "SUPERVISORUPDATE"
Private Const m_sObjectFieldName As String = "OBJECTNAME"
Private Const m_sOperationField As String = "SUPERVISOROPERATION"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    m_bOperationExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sOperationField)
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "SUPERVISOROBJECTINSTANCEID"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetInstanceID() As Variant
    Dim vval As Variant
    vval = m_clsTableAccess.GetVal("SUPERVISOROBJECTINSTANCEID")
    GetInstanceID = CStr(vval)
End Function
Friend Function GetOperation() As SupervisorPromoteType
    If m_bOperationExists Then
        GetOperation = m_clsTableAccess.GetVal(m_sOperationField)
    End If
End Function
Friend Function GetConnectionID() As Variant
    Dim vval As Variant
    
    vval = m_clsTableAccess.GetVal("SUPERVISORCONNECTIONID")

    GetConnectionID = vval
End Function
Friend Function GetCreationDate() As Variant
    Dim vval As Variant
    
    vval = m_clsTableAccess.GetDate("CREATIONDATE")

    GetCreationDate = vval
End Function
Friend Function GetDirection() As Variant
    Dim vval As Variant
    
    vval = m_clsTableAccess.GetVal("SUPERVISORDIRECTION")

    GetDirection = vval
End Function
Friend Function GetSupervisorUserID() As Variant
    Dim vval As Variant
    
    vval = m_clsTableAccess.GetVal("USERID")

    GetSupervisorUserID = vval
End Function
Friend Function GetObjectName() As String
    Dim vval As Variant
    
    vval = m_clsTableAccess.GetVal(m_sObjectFieldName)

    GetObjectName = vval
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetOperation(enumOperation As SupervisorPromoteType) As Boolean
    
    If m_bOperationExists Then
        SetOperation = m_clsTableAccess.SetVal(m_sOperationField, CVar(enumOperation))
    End If
End Function
Friend Function SetInstanceID(sVal As Variant) As Boolean
    SetInstanceID = m_clsTableAccess.SetVal("SUPERVISOROBJECTINSTANCEID", CVar(sVal))
End Function
Friend Function SetConnectionID(vval As Variant) As Boolean
    ' DJP SQL Server port
    Dim sVal As String
    
    If TypeName(vval) = "Byte()" Then
        sVal = g_clsSQLAssistSP.GuidToString(CStr(vval))
    Else
        sVal = CStr(vval)
    End If
    
    SetConnectionID = m_clsTableAccess.SetVal("SUPERVISORCONNECTIONID", sVal)
End Function
Friend Function SetCreationDate(sVal As Variant) As Boolean
    SetCreationDate = m_clsTableAccess.SetDate("CREATIONDATE", CVar(sVal))
End Function
Friend Function SetDirection(sVal As Variant) As Boolean
    SetDirection = m_clsTableAccess.SetVal("SUPERVISORDIRECTION", CVar(sVal))
End Function
Friend Function SetSupervisorUserID(sVal As Variant) As Boolean
    SetSupervisorUserID = m_clsTableAccess.SetVal("USERID", CVar(sVal))
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

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    m_clsTableAccess = clsTable
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
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetObjectNameField() As String
    GetObjectNameField = m_sObjectFieldName
End Function
Public Sub FindTableUpdates(ByVal sSrcServiceName As String, ByVal sSrcUserID As String, _
                            ByVal sDestServiceName As String, ByVal sDestUserID As String)
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSrcDirection As String
    Dim sDestDirection As String
    
    sSrcServiceName = g_clsSQLAssistSP.FormatString(sSrcServiceName)
    If Len(sSrcUserID) > 0 Then  'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
        sSrcUserID = g_clsSQLAssistSP.FormatString(sSrcUserID)
    End If
    sSrcDirection = g_clsSQLAssistSP.FormatString(DIRECTION_SOURCE)

    
    sDestServiceName = g_clsSQLAssistSP.FormatString(sDestServiceName)
    If Len(sDestUserID) > 0 Then  'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
        sDestUserID = g_clsSQLAssistSP.FormatString(sDestUserID)
    End If
    sDestDirection = g_clsSQLAssistSP.FormatString(DIRECTION_DEST)
        
sSearch = "SELECT " + _
    "Distinct F.SUPERVISOROBJECTNAME ""OBJECTNAME"" " + _
"From " + _
    "SUPERVISORUPDATE E, " + _
    "SUPERVISOROBJECTINSTANCEDET F, " + _
    "SUPERVISORCONNECTION Connection1, " + _
    "(SELECT " + _
        "SUPERVISOROBJECTNAME, MAX(CREATIONDATE) MAXDATE, " + _
        "b.SUPERVISOROBJECTVALUE " + _
     "FROM SUPERVISORUPDATE A, SUPERVISOROBJECTINSTANCEDET b " + _
    "Where " + _
        "A.SUPERVISOROBJECTINSTANCEID = B.SUPERVISOROBJECTINSTANCEID AND " + _
        "A.SUPERVISORDIRECTION = " + sSrcDirection + " " + _
    "GROUP BY B.SUPERVISOROBJECTNAME,B.SUPERVISOROBJECTVALUE " + _
    ") G " + _
"Where " + _
    "E.CREATIONDATE = G.MAXDATE AND " + _
    "E.SUPERVISOROBJECTINSTANCEID = F.SUPERVISOROBJECTINSTANCEID AND " + _
    "F.SUPERVISOROBJECTNAME = G.SUPERVISOROBJECTNAME AND " + _
    "F.SUPERVISOROBJECTVALUE = G.SUPERVISOROBJECTVALUE AND " + _
    "Connection1.SUPERVISORCONNECTIONID = " + _
    "(select SupervisorConnectionID from SupervisorConnection Connection2 "
    
    'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty - Start
    'Allow User IDs to be null
    If Len(sSrcUserID) > 0 Then
        sSearch = sSearch + "Where " + _
            "Connection2.SupervisorServiceName = " + sSrcServiceName + " AND " + _
            "Connection2.SupervisorSchemaName = " + sSrcUserID + " ) "
    Else
        sSearch = sSearch + "Where " + _
            "Connection2.SupervisorServiceName = " + sSrcServiceName + " AND " + _
            "Connection2.SupervisorSchemaName is null) "
    End If
        
    sSearch = sSearch + "AND E.CreationDate  NOT IN " + _
        "(select SupervisorUpdate2.CreationDate from SupervisorUpdate SupervisorUpdate2, SupervisorObjectInstanceDet Instance2 " + _
        "Where " + _
        "SupervisorUpdate2.SupervisorObjectInstanceID = Instance2.SupervisorObjectInstanceID AND " + _
        "SupervisorUpdate2.SupervisorConnectionID in " + _
        "(select SupervisorConnectionID from SupervisorConnection Connection3 "
                
    If Len(sDestUserID) > 0 Then
        sSearch = sSearch + "Where " + _
            "Connection3.SupervisorServiceName = " + sDestServiceName + " AND " + _
            "Connection3.SupervisorSchemaName = " + sDestUserID + " ) "
    Else
        sSearch = sSearch + "Where " + _
            "Connection3.SupervisorServiceName = " + sDestServiceName + " AND " + _
            "Connection3.SupervisorSchemaName is null) "
    End If
     'CORE00000313 (VR 37)/MAR1968 - End
        
    sSearch = sSearch + "AND SupervisorUpdate2.SupervisorDirection = " + sDestDirection + " )"

    'BS BM0498 01/05/03
    'Add Order By ComboGroupName
sSearch = sSearch + " ORDER BY 1"

    m_clsTableAccess.SetRecordSet g_clsDataAccess.ExecuteCommand(sSearch)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'DB SYS5457 07/01/03 - added wildcard parameter
Public Sub FindUpdateItemDetails(ByVal sObjectName As String, ByVal sSrcServiceName As String, ByVal sSrcUserID As String, _
                            ByVal sDestServiceName As String, ByVal sDestUserID As String, Optional bWildCard As Boolean = False)
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSrcDirection  As String
    Dim sDestDirection As String
    
    sSrcServiceName = g_clsSQLAssistSP.FormatString(sSrcServiceName)
    If Len(sSrcUserID) > 0 Then 'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
        sSrcUserID = g_clsSQLAssistSP.FormatString(sSrcUserID)
    End If
    sSrcDirection = g_clsSQLAssistSP.FormatString(DIRECTION_SOURCE)
        
    sDestServiceName = g_clsSQLAssistSP.FormatString(sDestServiceName)
    If Len(sDestUserID) > 0 Then 'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
        sDestUserID = g_clsSQLAssistSP.FormatString(sDestUserID)
    End If
    sDestDirection = g_clsSQLAssistSP.FormatString(DIRECTION_DEST)
    
    'DB SYS5457 07/01/03 If this is a wildcard search on the object name, add the wildcard
    If bWildCard Then
        Dim bWildCardFound As Boolean
        sObjectName = sObjectName & "%"
    End If
    'DB End
    
    sObjectName = g_clsSQLAssistSP.FormatString(sObjectName)
    
    'DB SYS5457 07/01/03 when looking for the object name, use LIKE operator instead
' TW 23/03/2007 EP2_1942
'    sSearch = "select su1.UserID, su1.SupervisorObjectInstanceID, su1.CreationDate, su1.SupervisorDirection, su1.SupervisorOperation, " + _
' TW 23/03/2007 EP2_1942 End
    sSearch = "select DISTINCT su1.UserID, su1.SupervisorObjectInstanceID, su1.CreationDate, su1.SupervisorDirection, su1.SupervisorOperation, " + _
    "so2.SupervisorObjectName , so2.SUPERVISOROBJECTKEY, so2.SUPERVISOROBJECTVALUE, so2.SupervisorObjectDescription, so2.SUPERVISOROBJECTVALUETYPE " + _
    "From " + _
      "supervisorupdate su1, " + _
      "supervisorconnection connection1, " + _
      "supervisorobjectinstancedet so2, " + _
      "(select so.SUPERVISOROBJECTKEY, so.SUPERVISOROBJECTVALUE, max(su.creationdate) MAXDATE,so.SupervisorObjectInstanceID "
     
sSearch = sSearch + "From " + _
        "supervisorupdate su, " + _
        "SupervisorObjectInstanceDet so " + _
      "Where su.SupervisorObjectInstanceID = so.SupervisorObjectInstanceID " + _
       "group by so.SUPERVISOROBJECTKEY, so.SUPERVISOROBJECTVALUE,so.SupervisorObjectInstanceID) MaxDateView " + _
      "Where su1.SupervisorObjectInstanceID = so2.SupervisorObjectInstanceID " + _
         "and su1.CreationDate = MaxDateView.maxdate and " + _
         "so2.SUPERVISOROBJECTKEY = MaxDateView.SUPERVISOROBJECTKEY And so2.SUPERVISOROBJECTVALUE = MaxDateView.SUPERVISOROBJECTVALUE " + _
         "AND su1.SupervisorDirection = " + sSrcDirection + " AND so2.SupervisorObjectName " & g_clsSQLAssistSP.GetLikeOperator + " " + sObjectName + " " + _
         "AND Connection1.SUPERVISORCONNECTIONID = "
         
    'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
    If Len(sSrcUserID) > 0 Then
        sSearch = sSearch + "(select SupervisorConnectionID From SupervisorConnection Connection2 " + _
        "Where Connection2.SupervisorServiceName = " + sSrcServiceName + " " + _
        "AND Connection2.SupervisorSchemaName = " + sSrcUserID + " ) "
    Else
        sSearch = sSearch + "(select SupervisorConnectionID From SupervisorConnection Connection2 " + _
        "Where Connection2.SupervisorServiceName = " + sSrcServiceName + " " + _
        "AND Connection2.SupervisorSchemaName is null) "
    End If
    
    sSearch = sSearch + "AND su1.CREATIONDATE  NOT IN " + _
     "(select " + _
        "SupervisorUpdate2.CreationDate " + _
      "From " + _
        "SupervisorUpdate SupervisorUpdate2, " + _
        "SupervisorObjectInstanceDet Instance2 " + _
      "Where"
      sSearch = sSearch + " SupervisorUpdate2.SupervisorObjectInstanceID = Instance2.SupervisorObjectInstanceID " + _
      "AND SupervisorUpdate2.SupervisorConnectionID in "
      
    'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
    If Len(sDestUserID) > 0 Then
        sSearch = sSearch + "(select SupervisorConnectionID From SupervisorConnection Connection3 " + _
        "Where Connection3.SupervisorServiceName = " + sDestServiceName + " " + _
        "AND Connection3.SupervisorSchemaName = " + sDestUserID + " ) "
    Else
        sSearch = sSearch + "(select SupervisorConnectionID From SupervisorConnection Connection3 " + _
        "Where Connection3.SupervisorServiceName = " + sDestServiceName + " " + _
        "AND Connection3.SupervisorSchemaName is null) "
    End If
      
    sSearch = sSearch + "AND SupervisorUpdate2.SupervisorDirection = " + sDestDirection + " ) " + _
           "order by su1.creationdate DESC , " + _
           "so2.SUPERVISOROBJECTKEY, " + _
           "so2.SUPERVISOROBJECTVALUE "
    
    m_clsTableAccess.SetRecordSet g_clsDataAccess.ExecuteCommand(sSearch)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub
Public Sub ResetAllUpdates(Optional sUserID As String)
    On Error GoTo Failed
    
    Dim sCommand As String
    
    sCommand = "DELETE FROM SUPERVISOROBJECTINSTANCE"
    
    If Len(sUserID) > 0 Then
        sCommand = sCommand + " WHERE SUPERVISOROBJECTINSTANCEID in (SELECT SUPERVISOROBJECTINSTANCEID FROM " + _
        "SUPERVISORUPDATE WHERE USERID = " + g_clsSQLAssistSP.FormatString(sUserID) + ")"
    End If
    
    g_clsDataAccess.ExecuteCommand sCommand
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
