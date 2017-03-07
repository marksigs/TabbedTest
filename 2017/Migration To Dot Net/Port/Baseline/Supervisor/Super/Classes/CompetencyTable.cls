VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompetencyTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : CompetencyTable
' Description   : Contains all code specific to the competency table. Implements TableAccess
' Change history
' Prog      Date        Description
' DJP       09/06/01    SQL Server port
' STB       03/12/01    SYS2468 - Competency bands should be without a 'Active To' date are
'                       also returned from get GetCompetenciesAsCollection().
' STB       22/04/02    MSMS0035 UserIncreaseLimitPercentage field added.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
' GHun      16/11/2005  MAR312 Changed GetCompetenciesAsCollection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "COMPETENCY"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "CompetencyType"
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsTableAccess.SetDuplicateKeys colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CompetencyType")
    GetType = CStr(vVal)
End Function
Friend Function GetTypeText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CompetencyTypeText")
    GetTypeText = CStr(vVal)
End Function
Friend Function GetActiveFrom() As String
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("CompetencyActiveFrom")
        
    If Not IsNull(vVal) Then
        GetActiveFrom = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetActiveTo() As String
    Dim vVal As Variant
        
    vVal = m_clsTableAccess.GetVal("CompetencyActiveTo")
        
    If Not IsNull(vVal) Then
        GetActiveTo = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetDescription() As String
    GetDescription = CStr(m_clsTableAccess.GetVal("CompetencyDescription"))
End Function
Friend Function GetRiskAssessmentMandateLevel() As String
    GetRiskAssessmentMandateLevel = CStr(m_clsTableAccess.GetVal("RiskAssessmentMandate"))
End Function
Friend Function GetFundsReleaseMandate() As String
    GetFundsReleaseMandate = CStr(m_clsTableAccess.GetVal("FundsReleaseMandate"))
End Function
Friend Function GetLoanAmountMandate() As String
    GetLoanAmountMandate = CStr(m_clsTableAccess.GetVal("LoanAmountMandate"))
End Function
Friend Function GetLTVMandate() As String
    GetLTVMandate = CStr(m_clsTableAccess.GetVal("LTVMandate"))
End Function
Friend Function GetUserMandateDecreaseLimitPercent() As String
    GetUserMandateDecreaseLimitPercent = CStr(m_clsTableAccess.GetVal("USERDECREASELIMITPERCENT"))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Function SetRiskAssmentMandateLevel(sType As String) As Boolean
    SetRiskAssmentMandateLevel = m_clsTableAccess.SetVal("RiskAssessmentMandate", CVar(sType))
End Function
Friend Function SetType(sType As String) As Boolean
    SetType = m_clsTableAccess.SetVal("CompetencyType", CVar(sType))
End Function
Friend Function SetActiveFrom(sActiveFrom As Variant) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("CompetencyActiveFrom", CVar(sActiveFrom))
End Function
Friend Function SetActiveTo(sActiveTo As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("CompetencyActiveTo", CVar(sActiveTo))
End Function
Friend Function SetFundsReleaseMandate(sFundsReleaseMandate As String) As Boolean
    SetFundsReleaseMandate = m_clsTableAccess.SetVal("FundsReleaseMandate", CVar(sFundsReleaseMandate))
End Function
Friend Function SetDescription(sDescription As String) As Boolean
    SetDescription = m_clsTableAccess.SetVal("CompetencyDescription", CVar(sDescription))
End Function
Friend Function SetLoanAmountMandate(sLoanAmountMandate As String) As Boolean
    SetLoanAmountMandate = m_clsTableAccess.SetVal("LoanAmountMandate", CVar(sLoanAmountMandate))
End Function
Friend Function SetLTVMandate(sLTVMandate As String) As Boolean
    SetLTVMandate = m_clsTableAccess.SetVal("LTVMandate", CVar(sLTVMandate))
End Function
Friend Function SetUserMandateDecreaseLimitPercent(sType As String) As Boolean
    SetUserMandateDecreaseLimitPercent = m_clsTableAccess.SetVal("USERDECREASELIMITPERCENT", CVar(sType))
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
    colListLine.Add GetTypeText()
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
    colListLine.Add Me.GetFundsReleaseMandate()
    colListLine.Add Me.GetLoanAmountMandate()
    colListLine.Add Me.GetLTVMandate()
    colListLine.Add Me.GetRiskAssessmentMandateLevel()
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
Friend Sub GetCompetencies()
    On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset

    sSearch = "select B.Valuename CompetencyTypeText, A. CompetencyType, A.CompetencyDescription, A.CompetencyActiveFrom, A.CompetencyActiveTo, A.FundsReleaseMandate, A.LoanAmountMandate, A.RiskAssessmentMandate,A.LTVMandate from " + m_sTableName + " A, ComboValue B " + _
              "where A.CompetencyType = B.Valueid and B.GroupName = 'CompetencyType'"

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    ValidateRecordset rs, m_sTableName
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR312 GHun Added extra optional parameter
Friend Sub GetCompetenciesAsCollection(colValues As Collection, colIDS As Collection, Optional ByVal blnHasWorkgroupLimit As Boolean = False)
    On Error GoTo Failed
    Dim sType As String
    Dim sSearch As String
    Dim sTypeText As String
    Dim sSystemDate As String
    Dim rs As ADODB.Recordset
    
    ' DJP SQL Server port
    sSystemDate = g_clsSQLAssistSP.GetSystemDate()

'    sSearch = "SELECT " & _
'                    "COMBO.VALUENAME COMPETENCYTYPETEXT, " & _
'                    "COMP.COMPETENCYTYPE " & _
'                "FROM " & _
'                    m_sTableName & " COMP, " & _
'                    "COMBOVALUE COMBO " & _
'                "WHERE " & _
'                    "COMP.COMPETENCYTYPE = COMBO.VALUEID AND " & _
'                    "COMBO.GROUPNAME = 'CompetencyType' AND " & _
'                    "(COMP.COMPETENCYACTIVETO >= " & sSystemDate & " OR " & _
'                    "COMP.COMPETENCYACTIVETO IS NULL)"

    'MAR312 GHun
    sSearch = "SELECT CV.ValueName COMPETENCYTYPETEXT," & _
                "C.COMPETENCYTYPE " & _
                "FROM " & m_sTableName & " C " & _
                "INNER JOIN ComboValue CV " & _
                    "ON C.CompetencyType = CV.ValueId " & _
                    "AND CV.GroupName = 'CompetencyType' " & _
                "LEFT OUTER JOIN ComboValidation CVAL " & _
                    "ON CVAL.ValueId = CV.ValueID " & _
                    "AND CVAL.GroupName = CV.GroupName " & _
                "WHERE " & _
                    "(C.CompetencyActiveTo >= " & sSystemDate & " OR " & _
                    "C.CompetencyActiveTo IS NULL)"
    
    If blnHasWorkgroupLimit Then
        sSearch = sSearch & " AND CVAL.ValidationType = '0'"
    End If
    'MAR312 End

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, "Competencies"
    m_clsTableAccess.SetRecordSet rs

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        While Not rs.EOF
            sTypeText = GetTypeText()
            colValues.Add sTypeText
            
            sType = GetType()
            colIDS.Add sType
            rs.MoveNext
        Wend
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No Competencies exist"
    End If
    rs.Close
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

Friend Function SetUserIncreaseLimitPercent(ByVal sUserIncreaseLimitPercent As String) As Boolean
    SetUserIncreaseLimitPercent = m_clsTableAccess.SetVal("USERINCREASELIMITPERCENT", CVar(sUserIncreaseLimitPercent))
End Function
Friend Function GetUserIncreaseLimitPercent() As String
    GetUserIncreaseLimitPercent = CStr(m_clsTableAccess.GetVal("USERINCREASELIMITPERCENT"))
End Function
