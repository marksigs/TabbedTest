VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnitTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   UnitTable
' Description   :   Contains all access functions for the Unit table.
' Change history
' Prog      Date        Description
' DJP       10/11/00    Created
' DJP       10/06/01    Updated for SQL Server port
' STB       05/04/02    SYS3327 - Added method GetUnitsForUserID() and
'                       GetChannelIDFromUnitID() for batch processing.
' SDS       06/02/02    SYS3951 - Modified the SQL query in method GetActiveUnits()
' CL        24/05/02    SYS4736 - Modification to SQL on GetUnitsForUserID
' CL        24/05/02    SYS4736 - 2nd revision to SQL on GetUnitsForUserID
' GHun      17/10/2005  MAR57 Added AllowOmigaLogOn and AllowOmigaExitFromWrap
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "UNIT"
Private m_sSearch As String

Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "UNITID"
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetUnitID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("UnitID")
    GetUnitID = CStr(vVal)
End Function
Friend Function GetActiveFrom() As String
    GetActiveFrom = m_clsTableAccess.GetDate("UnitActiveFrom")
End Function
' DJP SQL Server port, return variant
Friend Function GetAddressGUID() As Variant
    ' DJP SQL Server port
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AddressGUID")
    GetAddressGUID = vVal
End Function

Friend Function GetActiveTo() As String
    GetActiveTo = m_clsTableAccess.GetDate("UnitActiveTo")
End Function
Friend Function GetDXID() As String
    GetDXID = CStr(m_clsTableAccess.GetVal("DXID"))
End Function
Friend Function GetDXLocation() As String
    GetDXLocation = CStr(m_clsTableAccess.GetVal("DXLocation"))
End Function
Friend Function GetOrganisationIndicator() As String
    GetOrganisationIndicator = CStr(m_clsTableAccess.GetVal("OrganisationIndicator"))
End Function
Friend Function GetProcessingUnitIndicator() As String
    GetProcessingUnitIndicator = CStr(m_clsTableAccess.GetVal("ProcessingUnitIndicator"))
End Function
Friend Function GetUnitName() As String
    GetUnitName = CStr(m_clsTableAccess.GetVal("UnitName"))
End Function
Friend Function GetDepartmentID() As String
    GetDepartmentID = CStr(m_clsTableAccess.GetVal("DepartmentID"))
End Function
'MAR57 GHun
Friend Function GetAllowOmigaLogon() As String
    GetAllowOmigaLogon = CStr(m_clsTableAccess.GetVal("AllowOmigaLogon"))
End Function
Friend Function GetAllowOmigaExitFromWrap() As String
    GetAllowOmigaExitFromWrap = CStr(m_clsTableAccess.GetVal("AllowOmigaExitFromWrap"))
End Function
'MAR57 End

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetUnitID(sUnitID As String) As Boolean
    SetUnitID = m_clsTableAccess.SetVal("UnitID", CVar(sUnitID))
End Function
Friend Function SetActiveFrom(sActiveFrom As Variant) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("UnitActiveFrom", CVar(sActiveFrom))
End Function
Friend Function SetActiveTo(sActiveTo As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("UnitActiveTo", CVar(sActiveTo))
End Function
Friend Function SetDXID(sDXID As String) As Boolean
    SetDXID = m_clsTableAccess.SetVal("DXID", CVar(sDXID))
End Function
Friend Function SetDXLocation(sDXLocation As String) As Boolean
    SetDXLocation = m_clsTableAccess.SetVal("DXLocation", CVar(sDXLocation))
End Function
Friend Function SetOrganisationIndicator(sOrganisationIndicator As String) As Boolean
    SetOrganisationIndicator = m_clsTableAccess.SetVal("OrganisationIndicator", CVar(sOrganisationIndicator))
End Function
Friend Function SetProcessingUnitIndicator(sProcessingUnitIndicator As String) As Boolean
    SetProcessingUnitIndicator = m_clsTableAccess.SetVal("ProcessingUnitIndicator", CVar(sProcessingUnitIndicator))
End Function
Friend Function SetUnitName(sName As String) As Boolean
    SetUnitName = m_clsTableAccess.SetVal("UnitName", CVar(sName))
End Function
Friend Function SetDepartmentID(sID As String) As Boolean
    SetDepartmentID = m_clsTableAccess.SetVal("DepartmentID", CVar(sID))
End Function

Friend Function SetAddressGUID(sAddressGUID As String) As Boolean
    SetAddressGUID = m_clsTableAccess.SetVal("AddressGUID", CVar(sAddressGUID))
End Function

'MAR57 GHun
Friend Function SetAllowOmigaLogon(ByVal sAllowOmigaLogon As String) As String
    SetAllowOmigaLogon = m_clsTableAccess.SetVal("AllowOmigaLogon", CVar(sAllowOmigaLogon))
End Function
Friend Function SetAllowOmigaExitFromWrap(ByVal sAllowOmigaExitFromWrap As String) As String
    SetAllowOmigaExitFromWrap = m_clsTableAccess.SetVal("AllowOmigaExitFromWrap", CVar(sAllowOmigaExitFromWrap))
End Function
'MAR57 End

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
    colListLine.Add GetUnitID()
    colListLine.Add GetUnitName()
    colListLine.Add GetDepartmentID()
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
Friend Sub GetUnitsByRegion(sRegion As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    sSearch = "select * from " + m_sTableName + " where UnitID in( select UnitID from UnitRegion where RegionID = " + sRegion + ")    "
    'sSearch = "select * from " + m_sTableName + " where UnitID = '" + sRegion + "'"
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetUnitsByUnit(ByVal sUnit As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSearch As String

    Dim bFound As Boolean
    sUnit = g_clsSQLAssistSP.FormatWildcardedString(sUnit, bFound)
    sSearch = "select * from " + m_sTableName + " where UnitID like " + sUnit
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Sub GetUnitsByDepartment(sDepartment As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
    sSearch = "select * from " + m_sTableName + " where DepartmentID = '" + FormatString(sDepartment) + "'"
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetActiveUnits
' Description   :   Gets a list of active units
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetActiveUnits()
    On Error GoTo Failed
        
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    Dim sSystemDate As String
    
    sSystemDate = g_clsSQLAssistSP.TruncateDateColumn(g_clsSQLAssistSP.GetSystemDate())
    
    ' DJP SQL Server port
   sSearch = "SELECT " & _
                    "DISTINCT " & m_sTableName & ".UNITID " & _
                "FROM " & _
                    m_sTableName & ", USERROLE " & _
                "WHERE " & _
                    g_clsSQLAssistSP.TruncateDateColumn("UNITACTIVEFROM") & " <= " & sSystemDate & " AND (( " & _
                    g_clsSQLAssistSP.TruncateDateColumn("UNITACTIVETO") & " >= " & sSystemDate & " ) or UNITACTIVETO is null) AND " & _
                    "USERROLE.UNITID = " & m_sTableName & ".UNITID"

    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetUnitsAsCollection
' Description   : Locates all units and adds them to the collection
'                 passed in. If bActiveUnitsOnly is set to true, only active
'                 units are located.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetUnitsAsCollection(colValues As Collection, Optional bActiveUnitsOnly As Boolean = False)
    On Error GoTo Failed
    Dim sSysDate As String
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim sUnit As String
    Dim sUnitName As String
    
    sSearch = "SELECT " & _
                    "UNITID " & _
              "FROM " & _
                    m_sTableName
    
    If bActiveUnitsOnly = True Then
        ' DJP SQL Server port
        sSysDate = g_clsSQLAssistSP.GetSystemDate
        
        ' DJP SQL Server port
        sSearch = sSearch + " WHERE " & _
                                g_clsSQLAssistSP.TruncateDateColumn("UNITACTIVEFROM") & " <= " + g_clsSQLAssistSP.TruncateDateColumn(sSysDate)
    End If
    
    sSearch = sSearch + " ORDER BY UNITID"
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, "Units"
    m_clsTableAccess.SetRecordSet rs

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        While Not rs.EOF
            sUnit = GetUnitID()
            colValues.Add sUnit
            
            rs.MoveNext
        Wend
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No Units exist"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetComboField()
    GetComboField = "UNITID"
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetUnitsForUserID
' Description   : Populates the underlying table object with only the Units for
'                 the supplied UserID (where the users role is SYS ADMIN).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetUnitsForUserID(ByVal sUserID As String)

    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim sSystemDate As String
    
    sSystemDate = g_clsSQLAssistSP.GetSystemDate()
           
    sSearch = "SELECT  ur.UNITID, un.UNITNAME " & _
        "FROM  USERROLE ur, UNIT un, COMBOVALIDATION " & _
        "WHERE  ur.UNITID = un.UNITID " & _
        "AND ur.ROLE = COMBOVALIDATION.VALUEID " & _
        "AND COMBOVALIDATION.VALIDATIONTYPE = 'S' " & _
        "AND ur.USERID = " & g_clsSQLAssistSP.FormatString(sUserID) & _
        "AND COMBOVALIDATION.GROUPNAME = 'UserRole'" & _
        "AND " & g_clsSQLAssistSP.TruncateDateColumn("ur.USERROLEACTIVEFROM") & "<=" & sSystemDate & _
        "AND (" & g_clsSQLAssistSP.TruncateDateColumn("ur.USERROLEACTIVETO") & " IS NULL " & _
        "OR " & g_clsSQLAssistSP.TruncateDateColumn("ur.USERROLEACTIVETO") & " > " & sSystemDate & ")"

    'Populate the recordset from the query.
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    'Ensure we have a valid recordset.
    ValidateRecordset rs, m_sTableName
    
    'Assign the recordset to the base table object.
    m_clsTableAccess.SetRecordSet rs
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetChannelIDFromUnitID
' Description   : A lookup routine which will obtain the parent channel id
'                 from the specified UnitID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetChannelIDFromUnitID(ByVal sUnitID As String) As String

    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    'Get the ChannelID via the Department of the Unit.
    sSearch = "SELECT DP.CHANNELID " & _
              "FROM DEPARTMENT DP, UNIT UN " & _
              "WHERE DP.DEPARTMENTID = UN.DEPARTMENTID AND UN.UNITID = " & g_clsSQLAssistSP.FormatString(sUnitID)

    'Populate the recordset from the query.
    Set rs = g_clsDataAccess.GetTableData(, sSearch)

    'Ensure we have a valid recordset.
    ValidateRecordset rs, m_sTableName
    
    'Return the ChannelID.
    GetChannelIDFromUnitID = rs.fields("CHANNELID").Value

End Function
