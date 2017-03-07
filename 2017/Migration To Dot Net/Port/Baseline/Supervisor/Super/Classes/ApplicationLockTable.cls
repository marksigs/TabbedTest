VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationLockTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ApplicationLockTable
' Description   : Contains all code to access the ApplicationLock table, including and SQL
'                 statements and data access routines
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Constants
Private Const m_sTableName = "APPLICATIONLOCK"

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' BMIDS Change history
' Prog      Date        Description
' SA        19/11/02    BMIDS00992 - Customer Versioning consideration added to DeleteApplicationAndCustomers
' DJP       05/03/03    BM0282 Lock searching filtering added.
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "ApplicationNumber"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Friend Function GetApplicationNumber() As Variant
    GetApplicationNumber = m_clsTableAccess.GetVal("ApplicationNumber")
End Function

Friend Function GetUnitID() As Variant
    GetUnitID = m_clsTableAccess.GetVal("UnitID")
End Function

Friend Function GetUserID() As Variant
    GetUserID = m_clsTableAccess.GetVal("UserID")
End Function

Friend Function GetCustomerNumber() As Variant
    GetCustomerNumber = m_clsTableAccess.GetVal("CustomerNumber")
End Function

Friend Function GetLockDate() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("LockDate")
    
    If Not IsNull(vVal) Then
        GetLockDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetTypeOfLock() As Variant
    GetTypeOfLock = m_clsTableAccess.GetVal("TypeOfLock")
End Function

Friend Function GetMachineID() As Variant
    GetMachineID = m_clsTableAccess.GetVal("MachineID")
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
    colListLine.Add GetApplicationNumber()
    colListLine.Add GetLockDate()
    'colListLine.Add GetCustomerNumber()
    colListLine.Add GetUserID()
    colListLine.Add GetUnitID()
    colListLine.Add GetMachineID()
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Public Sub GetApplicationLocks(Optional nType As Variant, Optional sAppNo As String, Optional sUserID As String, Optional sUnitID As String)
    On Error GoTo Failed
    Dim sType As String
    Dim bFound As Boolean
    Dim sSearch As String
    Dim sSubSearch As String
    Dim bLockSearch As Boolean
    
    sSearch = "  SELECT " & _
                    "AL.LOCKDATE, " & _
                    "AL.UNITID, " & _
                    "AL.USERID, " & _
                    "AL.MACHINEID, " & _
                    "AL.APPLICATIONNUMBER " & _
                "FROM " & _
                    m_sTableName + " AL "

    
    bLockSearch = False
    If Not IsMissing(nType) Then

        Select Case nType
        Case OnLine
            sType = "On"
        Case Offline
            sType = "Off"
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "GetApplicationLocks: Invalid Application Type - " & nType
        End Select

        sSubSearch = "AL.TYPEOFLOCK = " & g_clsSQLAssistSP.FormatString(sType) & " OR " & _
                            "AL.TYPEOFLOCK IS NULL "
        bLockSearch = True
    End If
    
    If Len(sAppNo) > 0 Then
        If bLockSearch Then
            sSubSearch = sSubSearch & " AND "
        Else
            bLockSearch = True
        End If
        sSubSearch = sSubSearch & " AL.APPLICATIONNUMBER " & g_clsSQLAssistSP.GetLikeOperator & " " & sAppNo
    End If

    If Len(sUserID) > 0 Then
        If bLockSearch Then
            sSubSearch = sSubSearch & " AND "
        Else
            bLockSearch = True
        End If
        
        sSubSearch = sSubSearch & " AL.USERID " & g_clsSQLAssistSP.GetLikeOperator & " " & sUserID
    End If

    If Len(sUnitID) > 0 Then
        If bLockSearch Then
            sSubSearch = sSubSearch & " AND "
        Else
            bLockSearch = True
        End If
        
        sSubSearch = sSubSearch & " AL.UNITID " + g_clsSQLAssistSP.GetLikeOperator + " " & sUnitID
    End If

    If Len(sSubSearch) > 0 Then
        sSearch = sSearch & " WHERE " & sSubSearch
    End If

    m_clsTableAccess.GetTableDataFromSearch sSearch
    m_clsTableAccess.ValidateData

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

Public Sub DeleteApplicationAndCustomers(sApplicationNumber As String)
    On Error GoTo Failed
    Dim sCommand As String
    Dim rs As ADODB.Recordset
    Dim sCustomerNumber As String
    sCommand = "SELECT CUSTOMERNUMBER FROM CUSTOMERLOCKAPPLICATIONLOCK WHERE APPLICATIONNUMBER = " + g_clsSQLAssistSP.FormatString(sApplicationNumber)
    
    Set rs = g_clsDataAccess.ExecuteCommand(sCommand)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
                sCustomerNumber = sCustomerNumber + g_clsSQLAssistSP.FormatString(rs("CUSTOMERNUMBER"))
            
                rs.MoveNext
                
                If Not rs.EOF Then
                    sCustomerNumber = sCustomerNumber + ", "
                End If
            Wend
        End If
        rs.Close
    End If
    
    ' CustomerLockApplicationLock
    sCommand = "DELETE FROM CUSTOMERLOCKAPPLICATIONLOCK WHERE APPLICATIONNUMBER = " + g_clsSQLAssistSP.FormatString(sApplicationNumber)
    g_clsDataAccess.ExecuteCommand sCommand
    
    If Len(sCustomerNumber) > 0 Then
        ' Delete Customer locks
        'BMIDS00992 - but only if the same customer is not locked elsewhere...{
        On Error GoTo CustomerLockErr
        'BMIDS00992 }
        sCommand = "DELETE FROM CUSTOMERLOCK WHERE CUSTOMERNUMBER in (" + sCustomerNumber + ")"
        g_clsDataAccess.ExecuteCommand sCommand
        On Error GoTo Failed        'BMIDS00992
    End If
    
    ' Application locks
    sCommand = "DELETE FROM APPLICATIONLOCK WHERE APPLICATIONNUMBER = " + g_clsSQLAssistSP.FormatString(sApplicationNumber)
    g_clsDataAccess.ExecuteCommand sCommand
    
    Exit Sub
'BMIDS00992 Ignore foreign key constraint errors when deleting customer lock. {
CustomerLockErr:
    If Err.Number = -2147217873 Then
        Resume Next
    End If
'BMIDS00992 }
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

