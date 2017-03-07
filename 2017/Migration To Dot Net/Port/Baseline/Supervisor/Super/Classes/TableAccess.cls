VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TableAccess
' Description   : Class that contains prototypes for table access. Classes that use
'                 this must provide their own versions of the (public) prototypes listed
'                 here. There are routines for reading/writing updating/deleting.
' Change history
' Prog      Date        Description
' AA        24/02/01    Added ApplyFilter, CancelFilter
' DJP       26/06/01    SQL Server port
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Change
'                       AddRow to populate an empty recordset if it hasn't been created yet.
' STB       06/12/01    SYS1942 - AddRow, GetRecordSet and RecordCount now handles a closed
'                       recordset.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' note '''
' Need to change this so DoUpdates works for one record, or many. Example
' is GlobalBandedParameter and FixedParameters. Both should be able to use
' the same routine to do their updates.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Change history
' Prog      Date        Description
' PSC       07/03/2006  MAR1298 Allow for populating and ordering on primary keys
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' Prog      Date        Description
' PB        25/04/2006  EP448   Third party bank details no longer mandatory
' TW        21/03/2007  EP2_1224 - Run-time error if combo valuename entered with apostrophe
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const KEY_VIOLATION = 3631 + vbObjectError

' Private data
Private m_rs As ADODB.Recordset
Private m_bSuccess As Boolean
Private m_sSearch As String
Private m_sTable As String
Private m_bUpdated As Boolean
Private m_nFirstNewRow As Long
Private m_nLastNewRow As Long
Private m_colValues As Collection
Private m_PopulateType As PopulateType
Private m_colFields As Collection
Private m_colDeleteFields As Collection
Private m_sFirstMatch As String
Private m_bIsDeleted As Boolean
Private m_sDatabaseKey As String
Private m_sOrderCriteria As String
Private m_colDuplicateKeys As Collection
Private m_sCollectionField As String

Private m_colPrimaryKeys As Collection ' PSC 07/03/2006 MAR1298

' Enums
Private Enum DataOperation
    DATA_SELECT = 1
    DATA_DELETE = 2
End Enum


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetDatabase
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub SetDatabase(sDatabaseKey As String)
    m_sDatabaseKey = sDatabaseKey
End Sub


Public Function GetDatabase() As String
    GetDatabase = m_sDatabaseKey
End Function


Public Sub SetRecordSet(rs As ADODB.Recordset)
    Set m_rs = rs
End Sub


Public Function GetRecordSet() As ADODB.Recordset
    
'TODO: This may be required (when adding frmInterestRates for example).
'    If Not m_rs Is Nothing Then
'        'If the recordset is closed, open it. This can occur through the use of
'        'the 'Another' button where transactions are terminated but table objects
'        'are reused.
'        If m_rs.State = adStateClosed Then
'            GetTableData POPULATE_EMPTY
'        End If
'    End If
    
    'Return the recordset to the caller.
    Set GetRecordSet = m_rs
    
End Function


Public Sub Update()
    
    On Error GoTo Failed
    
    BeginWaitCursor
    
    If Not m_rs Is Nothing Then
        If m_rs.State = adStateOpen Then
            m_rs.UpdateBatch
            m_bIsDeleted = False
            SetUpdated

        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Recordset already closed"
        End If
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, "Update - "
    End If
    
    EndWaitCursor
    Exit Sub
Failed:
    EndWaitCursor
    g_clsErrorHandling.RaiseError Err.Number, GetErrorString(Err.Number, Err.DESCRIPTION)
    'DisplayErrors Err.Number, Err.DESCRIPTION
End Sub


Private Function GetErrorString(nNumber As Long, sDescription As String) As String
    On Error GoTo Failed
    Dim sError As String
    
    Select Case nNumber
'        Case KEY_VIOLATION
            'sError = "Duplicate record(s) found - please enter a unique combination"
        Case Else
            sError = sDescription
    End Select
    
    GetErrorString = sError
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sSetMatch As String
    Dim nFields As Integer
    Dim nMatch As Integer
    On Error GoTo Failed
    
    bRet = False
    
    nFields = colFields.Count
    nMatch = colMatchData.Count
    
    If nFields = nMatch Then
        ' Now build the query.
        
        sSetMatch = GenerateSearch(colFields, colMatchData)
        
        ' Do the search
        Set rs = g_clsDataAccess.GetTableData(, sSetMatch)
    Else
        MsgBox "MatchRecords: Fields collection count doesn't equal data collection count"
    End If
    
    Set MatchRecords = rs
    Exit Function
Failed:
    MsgBox "MatchRecords: Error is, " + Err.DESCRIPTION
    Set MatchRecords = Nothing
End Function


Public Function DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim colFields As Collection
    Dim rs As ADODB.Recordset
    
    If colMatchFields Is Nothing Then
        Set colFields = GetKeyMatchFields()
    Else
        Set colFields = colMatchFields
    End If
            
    Set rs = MatchRecords(colFields, colMatchData)
    bRet = False
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
        End If
        rs.Close
    End If

    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub Class_Initialize()
    m_bIsDeleted = False
End Sub


Public Sub CloseRecordSet()
    If Not m_rs Is Nothing Then
        If m_rs.State = adStateOpen Then
            m_rs.Close
            Set m_rs = Nothing
        End If
    End If
End Sub


Private Sub CreateNullValues()
    Dim nCount As Integer
    Dim nThisItem As Integer
    
    nCount = m_colFields.Count
    Set m_colValues = New Collection
    
    For nThisItem = 1 To nCount
        m_colValues.Add "null"
    Next
End Sub


' PSC 07/03/2006 MAR1298
Private Function PopulateKeys(Optional ByVal bOrderByPrimaryKeys As Boolean = False) As Boolean
    On Error GoTo Failed
    Dim nFields As Integer
    Dim nValues As Integer
    
    PopulateKeys = True
    If (Not m_colFields Is Nothing) And (Not m_colValues Is Nothing) Then
        nFields = m_colFields.Count
        nValues = m_colValues.Count

        If nFields = nValues And nValues > 0 Then
            ' PSC 07/03/2006 MAR1298
            m_sSearch = GenerateSearch(m_colFields, m_colValues, , True, bOrderByPrimaryKeys)
        Else
            MsgBox "GetTableData: Field count does not equal field value count"
            PopulateKeys = False
        End If
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Dim sTableName As String
    Dim bRet As Boolean
    
    BeginWaitCursor
    
    bRet = True
    m_bUpdated = False
    sTableName = GetTable()
    m_sSearch = ""
    
    Select Case enumPopulateType
    Case POPULATE_EMPTY
        CreateNullValues
        bRet = PopulateKeys()
    
    Case POPULATE_FIRST_BAND
        If Len(m_sFirstMatch) > 0 Then
            m_sSearch = m_sFirstMatch
        Else
            MsgBox "GetTableData: Can't populate first band, search is empty"
            bRet = False
        End If
        
    Case POPULATE_KEYS
        bRet = PopulateKeys()
    
    Case POPULATE_TABLE
        m_sSearch = sTableName
    
    Case POPULATE_ALL
        m_sSearch = "SELECT * FROM " + sTableName
    
        If Len(m_sOrderCriteria) > 0 Then
            m_sSearch = m_sSearch + " ORDER BY " + m_sOrderCriteria
        End If
    ' PSC 07/03/2006 MAR1298 - Start
    Case POPULATE_KEYS_ORDERBY_PRIMARY
        bRet = PopulateKeys(True)
    ' PSC 07/03/2006 MAR1298 - End
    End Select
    
    CloseRecordSet
    
    If bRet = True Then
        'MsgBox m_sSearch
        Set m_rs = g_clsDataAccess.GetTableData(sTable:=sTableName, sSearch:=m_sSearch, sDatabaseName:=m_sDatabaseKey, enumClassOption:=enumClassOption)
    Else
        Set m_rs = Nothing
    End If
    
    Set GetTableData = m_rs
    EndWaitCursor
End Function


Public Sub GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    On Error GoTo Failed
    
    CloseRecordSet
    
    Set m_rs = g_clsDataAccess.GetTableData(m_sTable, sSearch, m_sDatabaseKey, bUseExistingSearch)
    'PB 19/10/2006 E2_13 Begin
    'limit results returned if lngResultCap parameter is supplied
    If Not m_rs Is Nothing Then
        If g_lngResultCap > 0 Then
            If m_rs.RecordCount >= g_lngResultCap Then
                MsgBox "Too many rows were returned." & vbCrLf & "Please apply additional filtering.", vbInformation, "Too many records returned"
                m_rs.Close
            End If
        End If
        If m_rs.State = adStateOpen Then
        'PB 19/10/2006 E2_13 End
            ValidateRecordset m_rs, m_sTable
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' PSC 07/03/2006 MAR1298
Private Function GenerateSearch(colFields As Collection, colValues As Collection, Optional enumOperation As DataOperation = DATA_SELECT, Optional bAllowOrdering As Boolean = False, Optional bOrderByPrimaryKeys As Boolean = False) As String
    Dim sOperation As String
    Dim sCriteria As String
    Dim sSearch As String
    Select Case enumOperation
        Case DATA_SELECT
            sOperation = "SELECT * "
        Case DATA_DELETE
            sOperation = "DELETE "
    End Select
    
    sSearch = sOperation & " from " + m_sTable + " where "
    
    ' DJP SQL Server port - moved to DataAccess
    'sCriteria = g_clsFormProcessing.GenerateSearchCriteria(colFields, colValues, m_sDatabaseKey)
    sCriteria = g_clsDataAccess.GenerateSearchCriteria(colFields, colValues, m_sDatabaseKey)
    sSearch = sSearch + sCriteria
    
    If bAllowOrdering Then
        ' PSC 07/03/2006 MAR1298 - Start
        If bOrderByPrimaryKeys = False Then
            If Len(m_sOrderCriteria) > 0 Then
                sSearch = sSearch + " ORDER BY " + m_sOrderCriteria
            End If
        Else
            If m_colPrimaryKeys Is Nothing Then
                Set m_colPrimaryKeys = g_clsDataAccess.GetTablePrimaryKeys(m_sTable)
            End If
                
            Dim strOrderBy As String
            Dim strKey As Variant
                        
            For Each strKey In m_colPrimaryKeys
                If Len(strOrderBy) > 0 Then
                   strOrderBy = strOrderBy & ","
                End If
                strOrderBy = strOrderBy & strKey
            Next
            
            sSearch = sSearch & " ORDER BY " & strOrderBy
        End If
        ' PSC 07/03/2006 MAR1298 - End
    End If
    
    GenerateSearch = sSearch
End Function


Public Sub AddRow()
    
    Dim sTableName As String
    
    On Error GoTo Failed
    
    'Get the underlying table name.
    sTableName = GetTable()
    
    'If the recordset doesn't exist, initalise one.
    If m_rs Is Nothing Then
        GetTableData POPULATE_EMPTY
    'If the recordset exists, if its closed, the initialise it.
    ElseIf m_rs.State = adStateClosed Then
        GetTableData POPULATE_EMPTY
    End If

    'Ensure we have a recordset.
    ValidateRecordset m_rs, sTableName
    
    'Now add a new record to the recordset.
    m_rs.AddNew

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub DeleteAllRows()
    On Error GoTo Failed
    Dim nRecordCount As Integer
    Dim nThisRecord As Integer
    ValidateData
    
    nRecordCount = m_rs.RecordCount
    
    If nRecordCount > 0 Then
        m_rs.MoveFirst

        For nThisRecord = 1 To nRecordCount
            m_rs.Delete
            m_rs.MoveNext
        Next
    End If

    m_bIsDeleted = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub DeleteRow(colMatchData As Collection)
    Dim colKeys As Collection
    On Error GoTo Failed
    
    If Not m_colDeleteFields Is Nothing Then
        Set m_rs = MatchRecords(m_colDeleteFields, colMatchData)
    Else
        Set m_rs = MatchRecords(m_colFields, colMatchData)
    End If
    
    ValidateData
    
    If m_rs.BOF = False And m_rs.EOF = False Then
        m_rs.MoveFirst

        Do While Not m_rs.EOF
            m_rs.Delete
            m_rs.MoveNext
        Loop
    Else
        g_clsErrorHandling.RaiseError errRecordNotFound, m_sTable
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.d
End Sub


Public Sub ValidateData(Optional sDesc As String = "")
    On Error GoTo Failed
    ValidateRecordset m_rs, sDesc
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetVal(sField As Variant) As Variant
    On Error GoTo Failed
    Dim vTmp As Variant
    
    ValidateData
    If m_rs.RecordCount > 0 Then
        vTmp = m_rs(sField)
        ' DJP SQL Server port
        If Not (m_rs(sField).Type = adBinary Or m_rs(sField).Type = adVarBinary) Then
            If Not IsNull(vTmp) Then
' TW 21/03/2007 EP2_1224
'                GetVal = CStr(vTmp)
                If m_rs(sField).Type = adVarChar Or m_rs(sField).Type = adChar Or m_rs(sField).Type = adWChar Then
                    GetVal = FormatString(CStr(vTmp))
                Else
                    GetVal = CStr(vTmp)
                End If
' TW 21/03/2007 EP2_1224 End
            Else
                GetVal = ""
            End If
        Else
            GetVal = vTmp
        End If
    Else
        'BMIDS00565 08/10/02 CL
        'EP448 25/04/2006 PB
        'If sField <> "GroupName" Then
        If sField <> "GroupName" And sField <> "BankSortCode" Then
            'EP448 END
            g_clsErrorHandling.RaiseError errRecordSetEmpty
        End If
        'BMIDS00565 08/10/02 CL END
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "Error reading from field <" + sField + "> in table <" + m_sTable + "> - " + Err.DESCRIPTION
End Function

Public Function SetVal(sField As Variant, vVal As Variant) As Boolean
    On Error GoTo Failed
    Dim vTmp As Variant
    vTmp = vVal
    
    ' DJP MCAP
    ' Does the field exist?
    If g_clsDataAccess.DoesFieldExist(m_sTable, CStr(sField)) Then
        ValidateData
        
        If Not IsNull(vTmp) Then
            If TypeName(vTmp) = "Byte()" Then
                vTmp = g_clsSQLAssistSP.GuidToString(CStr(vVal))
            End If
            
            If Len(vTmp) > 0 Then
                If m_rs(sField).Type = adVarChar Or m_rs(sField).Type = adChar Or m_rs(sField).Type = adWChar Then
' TW 21/03/2007 EP2_1224
'                    vTmp = CStr(vTmp)
                    vTmp = Replace(CStr(vTmp), "''", "'")
' TW 21/03/2007 EP2_1224 End
                End If
                m_rs(sField) = vTmp
            Else
                m_rs(sField) = Null
            End If
        Else
            m_rs(sField) = Null
        End If
    End If
    
    SetVal = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "Error updating field <" + sField + "> in table <" + m_sTable + "> - " + Err.DESCRIPTION
    SetVal = False
End Function


Public Function SetNextRow() As Boolean
    On Error GoTo Failed
    m_rs.Move 1
    SetNextRow = True
    Exit Function
Failed:
    MsgBox "SetNextRow: Error is, " + Err.DESCRIPTION
    SetNextRow = False
End Function


Public Function RecordCount() As Integer
    
    Dim nRecordCount As Integer
    
    nRecordCount = 0
    
    If Not m_rs Is Nothing Then
        If m_rs.State = adStateClosed Then
            nRecordCount = 0
        Else
            nRecordCount = m_rs.RecordCount()
        End If
    End If
    
    RecordCount = nRecordCount
    
End Function


Public Function MoveNext() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = False
    If m_rs.EOF = False Then
        m_rs.MoveNext
        bRet = True
    End If
    
    If m_rs.EOF = True Then
        bRet = False
    End If
    
    MoveNext = bRet
    Exit Function
Failed:
    MsgBox "MoveNext: Error is, " + Err.DESCRIPTION
    MoveNext = False
End Function


Public Sub MoveFirst()
    On Error GoTo Failed
    
    If m_rs.RecordCount > 0 Then
        m_rs.MoveFirst
    Else
        g_clsErrorHandling.RaiseError errRecordNotFound, m_sTable
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetUpdated()
    m_bUpdated = True
End Sub


Public Function GetUpdated() As Boolean
    GetUpdated = m_bUpdated
End Function


Public Function GetIsDeleted() As Boolean
    GetIsDeleted = m_bIsDeleted
End Function


Public Function GetTable() As String
    GetTable = m_sTable
End Function

Public Function GetSearch() As String
    GetSearch = m_sSearch
End Function


Public Sub SetPopulateNull()
    
    Dim sTableName As String
    
    sTableName = GetTable()
    
    m_sSearch = "Select * from " + sTableName + " where Name is null"
End Sub


Public Function GetUpdateValues() As Collection
End Function


Public Sub GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
End Sub


Public Sub Create(sTable As String, Optional sFirstMatch As String = "")
    m_sTable = sTable
    m_sFirstMatch = sFirstMatch
End Sub


Public Sub SetDeleteFields(colFields As Collection)
    Set m_colDeleteFields = colFields
End Sub


Public Sub SetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    
    Set m_colValues = colValues
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetKeyMatchFields(colFields As Collection)
    Set m_colFields = colFields
End Sub


Public Function GetKeyMatchValues()
    Set GetKeyMatchValues = m_colValues
End Function


Public Function GetKeyMatchFields()
    Set GetKeyMatchFields = m_colFields
End Function


Public Sub SetPopulateType(enumType As PopulateType)
End Sub


Public Function SetDate(sField As String, vVal As Variant) As Boolean
    On Error GoTo Failed
    Dim vDate As Variant
    
    If Not IsNull(vVal) Then
        vDate = g_clsValidation.GetFullDate(CStr(vVal))
    Else
        vDate = vVal
    End If
    
    SetDate = SetVal(sField, vDate)
    
    Exit Function
Failed:
    SetDate = False
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function GetLength(sField As String) As Integer
    On Error GoTo Failed
    GetLength = m_rs(sField).DefinedSize
    
    Exit Function
Failed:
    MsgBox "GetLength - error is " + Err.DESCRIPTION
    GetLength = 0
End Function


Public Sub DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    Dim sCommand As String
    
    If Len(sDatabase) = 0 Then
        sDatabase = m_sDatabaseKey
    End If
    
    If g_clsDataAccess.DoesTableExist(m_sTable) Then
        sCommand = GenerateSearch(m_colFields, m_colValues, DATA_DELETE)
        g_clsDataAccess.ExecuteCommand sCommand, sDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Function GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    On Error GoTo Failed
    Dim sVal As String
    
    sVal = GetData(sField, ctrl, ctrlLabel, bCheck)
    
    If Len(sVal) > 0 Then
        If Not bGetTime Then
            sVal = Format(sVal, "Short Date")
        End If
        
        GetDate = g_clsValidation.GetFullDate(sVal)
    Else
        GetDate = sVal
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Friend Function GetData(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True) As String
    On Error GoTo Failed
    Dim sText As String
    Dim bExists As Boolean
    
    bExists = True
    
    If bCheck Then
        bExists = g_clsDataAccess.DoesFieldExist(m_sTable, sField)
    End If
    
    If bExists Then
        sText = CStr(GetVal(sField))
        
        If Not ctrl Is Nothing Then
            g_clsFormProcessing.SetText ctrl, True, sText
            g_clsFormProcessing.SetText ctrlLabel, True
        End If
    Else
        If Not ctrl Is Nothing Then
            g_clsFormProcessing.SetText ctrl, False
            g_clsFormProcessing.SetText ctrlLabel, False
        End If
    End If
    
    GetData = sText
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function


Friend Sub SetData(sField As String, vText As Variant)
    On Error GoTo Failed
    
    ' Does the field exist?
    If g_clsDataAccess.DoesFieldExist(m_sTable, sField) Then
        SetVal sField, vText
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetOrderCriteria(sCriteria As String)
    m_sOrderCriteria = sCriteria
End Sub


Public Function GetDuplicateKeys() As Collection
    Set GetDuplicateKeys = m_colDuplicateKeys
End Function


Public Sub SetDuplicateKeys(colDupKeys As Collection)
    On Error GoTo Failed
    
    If Not colDupKeys Is Nothing Then
        Set m_colDuplicateKeys = colDupKeys
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetDuplicateKeys: No keys to set"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetCollectionField(sField As String)
    m_sCollectionField = sField
End Sub


Public Function GetCollectionField() As String
    On Error GoTo Failed
    
    If Len(m_sCollectionField) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Collection field is empty"
    End If
    
    GetCollectionField = m_sCollectionField
    
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub Class_Terminate()
'    If Not m_rs Is Nothing Then
'        If m_rs.State = adStateOpen Then
'            m_rs.Close
'        End If
'    End If
End Sub


Public Sub ApplyFilter(Optional sSearch As String)
    
    On Error GoTo Failed
    
    If Len(sSearch) = 0 Then
        ' DJP SQL Server port - moved to DataAccess
        sSearch = g_clsDataAccess.GenerateSearchCriteria(m_colFields, m_colValues)
    End If
    
    m_rs.Filter = sSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub


Public Sub CancelFilter()
    
    On Error GoTo Failed
    
    m_rs.Filter = adFilterNone
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub
