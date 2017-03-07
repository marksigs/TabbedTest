VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupervisorObjectTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : SupervisorObjectTable
' Description   : Contains all code to manage access to the SupervisorObject table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port.
' STB       11/01/02    SYS3581 Records keyed on GUIDs are handled.
' SDS       07/03/02    SYS3859 Base rate sets with same set number but different start dates
'                       had an error while promoting
' SR        21/11/2007  CORE00000313(VR 37)/MAR1968 - Allow UserID to be empty (when using integrated security)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "SUPERVISOROBJECTINSTANCE"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
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
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("SUPERVISOROBJECTINSTANCEID")
    ' DJP SQL Server port removed converting of GUID to string
    GetInstanceID = vVal
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetInstanceID() As Variant
    Dim clsGUID As New GuidAssist
    Dim vVal As Variant
    Dim bRet As Boolean
    
    vVal = CVar(clsGUID.CreateGUID())
    
    bRet = m_clsTableAccess.SetVal("SUPERVISOROBJECTINSTANCEID", vVal)

    If bRet = True Then
        ' DJP SQL Server port, convert to byte format
        SetInstanceID = vVal 'g_clsSQLAssistSP.GuidStringToByteArray(vVal)
    Else
        SetInstanceID = ""
    End If
    
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
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function DoesConnectionExist(sServiceName As String, sUserID As String) As Boolean
    On Error GoTo Failed
    
    FindConnection sServiceName, sUserID

    If m_clsTableAccess.RecordCount > 0 Then
        DoesConnectionExist = True
    Else
        DoesConnectionExist = False
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
Public Sub FindConnection(sServiceName As String, sUserID As String)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
     'CORE00000313 (VR 37)/MAR1968 - Allow UserID to be empty
    If Len(sServiceName) > 0 Then
        If Len(sUserID) > 0 Then
            sSearch = "select * from " + m_sTableName + " where  = SUPERVISORSCHEMANAME = '" + FormatString(sUserID) + "' AND " + _
            " SUPERVISORSERVICENAME = '" + FormatString(sServiceName) + "'"
        Else
            sSearch = "select * from " + m_sTableName + " where  = SUPERVISORSCHEMANAME is null AND " + _
            " SUPERVISORSERVICENAME = '" + FormatString(sServiceName) + "'"
        End If
        
        Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "FindConnection: Service name must be set"
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
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function
Public Sub FindInstance(colKeys As Collection, colValues As Collection, ByVal sName As String)
    On Error GoTo Failed
    Dim sFunctionName As String
    Dim sSearch As String
    Dim sKey As String
    Dim sField As String
    Dim sValue As String
    Dim nKeyCount As Long
    Dim nThisKey As Long
    Dim sFrom As String
    Dim sWhere As String
    Dim rs As ADODB.Recordset
    sFunctionName = "FindInstance"
    
    If Len(sName) > 0 Then
        sName = g_clsSQLAssistSP.FormatString(sName)
        If Not colKeys Is Nothing And Not colValues Is Nothing Then
            If colKeys.Count = colValues.Count And colKeys.Count > 0 Then
                
                nKeyCount = colKeys.Count
                sFrom = "SELECT KEY1.SUPERVISOROBJECTINSTANCEID FROM " + _
                " SUPERVISOROBJECTINSTANCEDET KEY1"
                
                sField = colKeys(1)
                
                'If this value is a byte array, then format it as a string.
                If TypeName(colValues(1)) = "Byte()" Then
                    sValue = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(colValues(1)))
                Else
                    sValue = colValues(1)
                End If
                
                sWhere = " WHERE KEY1.SUPERVISOROBJECTNAME = " + sName + " AND " + _
                "KEY1.SUPERVISOROBJECTKEY = " + g_clsSQLAssistSP.FormatString(sField) + " AND " + _
                "KEY1.SUPERVISOROBJECTVALUE = " + g_clsSQLAssistSP.FormatString(sValue)
                
                For nThisKey = 2 To nKeyCount
                    sKey = "KEY" & nThisKey
                    sFrom = sFrom + ", SUPERVISOROBJECTINSTANCEDET " & sKey
                    
                    sField = colKeys(nThisKey)
                
                    'If this value is a byte array, then format it as a string.
                    If TypeName(colValues(nThisKey)) = "Byte()" Then
                        sValue = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(colValues(nThisKey)))
                    Else
                        sValue = colValues(nThisKey)
                    End If
                    
                    sWhere = sWhere + " AND " + sKey & ".SUPERVISOROBJECTNAME = " + sName + _
                    " AND " + sKey + ".SUPERVISOROBJECTKEY = " + g_clsSQLAssistSP.FormatString(sField) + " AND " + _
                    sKey + ".SUPERVISOROBJECTVALUE = " + g_clsSQLAssistSP.FormatString(sValue) + " AND " + _
                    "KEY1.SUPERVISOROBJECTINSTANCEID = " + sKey + ".SUPERVISOROBJECTINSTANCEID"
                    
                Next
                sSearch = sFrom + sWhere
                
                m_clsTableAccess.GetTableDataFromSearch sSearch
                                
                
            Else
                g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key and value counts don't match"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Keys and values are empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Object Name is empty"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
