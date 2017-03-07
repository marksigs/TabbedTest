VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseRateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BaseRateTable
' Description   : Contains all access routines and SQL for the BaseRate table. Implements
'                 TableAccess and BandedTable as it is a banded table.
' Change history
' Prog      Date        Description
' DJP       11/06/01    SQL Server port.
' DJP       04/08/01    DJP don't log fact that version fields don't exist
' SA        18/12/01    SYS3504 GetRateDifference & SetRateDifference added
' SDS       06/02/02    SYS2608 (13) In order to check whether mortgage products exist
'                       for the selected base rate , modification was made in method TableAccess_DeleteRecords()
' SDS       07/03/02    SYS2608
' STB       07/03/02    SYS4224 Only display base rate sets with a start date in the future.
' STB       08/03/02    SYS4238 Added GetBandsForSet() method.
' STB       05/04/02    SYS4247 If the base rate band is in use, raise an error rather than displaying
'                       a message, otherwise other objects may get deleted...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
Implements BandedTable

' Constants
Private Const m_sTableName = "BASERATEBAND"
Private Const m_sVersionField As String = "BRSETVERSIONNUMBER"

' Private data
Private m_bVersionExists As Boolean
Private m_sSearch As String
Private m_sFirstMatch As String
Private m_clsBandedTable As BandedTable
Private m_clsTableAccess As TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetBandsForSet
' Description   : Populates the underlying table object with bands for the
'                 given BaseRateSet number. If none is specified then an empty
'                 recordset is created. This routine ensures the RATE column is
'                 not used.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' AW        29/05/02    BMID00016   Changed DeleteRecord to use generic method DoesProductExistForThisItem()
' GD        02/07/2002  BMIDS00081 - application of core AQR SYS4935



Public Sub GetBandsForSet(Optional ByVal vSetNumber As Variant = -99, Optional ByVal vStartDate As Variant = 0)
    '' GD        02/07/2002  BMIDS00081 - application of core AQR SYS4935
    Dim sSQL As String
        
    '-99 should be an impossible rate number. This will give us an empty recordset
    'with the desired columns.
    'sSQL = "SELECT BASERATESET, BASERATEBANDSTARTDATE, MAXIMUMLOANAMOUNT, MAXIMUMLTV, RATEDIFFERENCE FROM BASERATEBAND WHERE BASERATESET = " & CDec(vSetNumber)
    'SYS4935 SA Add Start date to SQL
    sSQL = "SELECT BASERATESET, BASERATEBANDSTARTDATE, MAXIMUMLOANAMOUNT, MAXIMUMLTV, RATEDIFFERENCE FROM BASERATEBAND WHERE BASERATESET = " & CDec(vSetNumber) & " AND BASERATEBANDSTARTDATE = '" & CDate(vStartDate) & "'"
    m_clsTableAccess.GetTableDataFromSearch sSQL

End Sub


Friend Function GetFeeSet() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BaseRateSet")
    GetFeeSet = CStr(vVal)
End Function

Friend Function GetStartDate() As String
    Dim vVal As Variant

    vVal = m_clsTableAccess.GetVal("BaseRateBandStartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetMaxLoanAmount() As String
    GetMaxLoanAmount = CStr(m_clsTableAccess.GetVal("MaximumLoanAmount"))
End Function

Friend Function GetRateDifference() As String
    GetRateDifference = CStr(m_clsTableAccess.GetVal("RateDifference"))
End Function

Friend Function GetMaxLTV() As String
    GetMaxLTV = CStr(m_clsTableAccess.GetVal("MaximumLTV"))
End Function

Friend Function GetRate() As String
    GetRate = CStr(m_clsTableAccess.GetVal("Rate"))
End Function

Friend Function GetDescription() As String
    GetDescription = CStr(m_clsTableAccess.GetVal("BASERATESETDESCRIPTION"))
End Function

Friend Function SetFeeSet(sFeeSet As String) As Boolean
    SetFeeSet = m_clsTableAccess.SetVal("BaseRateSet", CVar(sFeeSet))
End Function

Friend Function SetDescription(sDescription As String) As Boolean
    SetDescription = m_clsTableAccess.SetVal("BASERATESETDESCRIPTION", CVar(sDescription))
End Function

Friend Function SetStartDate(sStartDate As String) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("BaseRateBandStartDate", CVar(sStartDate))
End Function

Friend Function SetMaxLoanAmount(sMaxLoanAmount As String) As Boolean
    SetMaxLoanAmount = m_clsTableAccess.SetVal("MaximumLoanAmount", CVar(sMaxLoanAmount))
End Function

Friend Function SetRateDifference(sRateDifference As String) As Boolean
    SetRateDifference = m_clsTableAccess.SetVal("RateDifference", CVar(sRateDifference))
End Function

Friend Function SetMaxLTV(sMaxLTV As String) As Boolean
    SetMaxLTV = m_clsTableAccess.SetVal("MaximumLTV", CVar(sMaxLTV))
End Function

Friend Function SetRate(sRate As String) As Boolean
    SetRate = m_clsTableAccess.SetVal("Rate", CVar(sRate))
End Function

Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub

Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub

Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub

Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub

Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    Dim bVersionExists As Boolean
    
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    
    m_sFirstMatch = "select * from BaseRateBand A where A.MAXIMUMLOANAMOUNT = " + _
    "(select min(MAXIMUMLOANAMOUNT) from  BaseRateBand B " + _
        "where A.BASERATESET = B.BASERATESET and A.BaseRateBandStartDate = B.BaseRateBandStartDate) and A.MAXIMUMLTV = " + _
    "(select min(MAXIMUMLTV) from BaseRateBand C " + _
        "where a.BaseRateBandStartDate = c.BaseRateBandStartDate and a.BASERATESET = C.BASERATESET and " + _
        "A.MAXIMUMLOANAMOUNT = C.MAXIMUMLOANAMOUNT)"
    
    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
        
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch

    colFields.Add "BaseRateSet"
    colFields.Add "BaseRateBandStartDate"
    
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "BaseRateSet"
    colUpdateFields.Add "BaseRateBandStartDate"

    m_clsBandedTable.SetUpdateFields colUpdateFields
    
    Set colDupKeys = New Collection
    colDupKeys.Add "MAXIMUMLOANAMOUNT"
    colDupKeys.Add "MAXIMUMLTV"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub SetFindNewestBaseRateSet()
    Dim sSysDate As String
    
    sSysDate = g_clsSQLAssistSP.GetSystemDate
    m_sFirstMatch = "SELECT " & _
                        "BR1.* " & _
                    "FROM " & _
                        "BASERATEBAND BR1 " & _
                    "WHERE " & _
                        g_clsSQLAssistSP.TruncateDateColumn("BASERATEBANDSTARTDATE") & " IN " & _
                        "(SELECT " & _
                            "MAX(" & g_clsSQLAssistSP.TruncateDateColumn("BASERATEBANDSTARTDATE") & ") " & _
                        "FROM " & _
                            "BASERATEBAND BR2 " & _
                        "WHERE " & _
                            "BR1.BASERATESET = BR2.BASERATESET AND " & _
                            g_clsSQLAssistSP.TruncateDateColumn(sSysDate) & " > " & g_clsSQLAssistSP.TruncateDateColumn("BASERATEBANDSTARTDATE") & ") AND " & _
                            "BR1.MAXIMUMLOANAMOUNT = "
    m_sFirstMatch = m_sFirstMatch & "(SELECT " & _
                                            "MIN(MAXIMUMLOANAMOUNT) " & _
                                        "FROM " & _
                                            "BASERATEBAND BR3 " & _
                                        "WHERE " & _
                                            "BR1.BASERATESET = BR3.BASERATESET AND " & _
                                            g_clsSQLAssistSP.TruncateDateColumn("BR1.BASERATEBANDSTARTDATE") & " = " & g_clsSQLAssistSP.TruncateDateColumn("BR3.BASERATEBANDSTARTDATE") & ") AND BR1.MAXIMUMLTV = " & _
                                            "(SELECT " & _
                                                "MIN(MAXIMUMLTV) " & _
                                            "FROM " & _
                                                "BASERATEBAND BR4 " & _
                                            "WHERE " & _
                                                g_clsSQLAssistSP.TruncateDateColumn("BR1.BASERATEBANDSTARTDATE") & " = " & g_clsSQLAssistSP.TruncateDateColumn("BR4.BaseRateBandStartDate") + " and BR1.BASERATESET = BR4.BASERATESET AND " + _
                                                "BR1.MAXIMUMLOANAMOUNT = BR4.MAXIMUMLOANAMOUNT)"
    
    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If

    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub

Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetFeeSet()
    colHeader.Add GetStartDate()
End Sub

Public Sub SetPopulateNull()
    m_sSearch = "Select * from " + m_sTableName + " where Name is null"
End Sub

Public Sub SetPopulateType(enumType As PopulateType)
    m_clsTableAccess.SetPopulateType enumType
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

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub

Private Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
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

Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
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

Private Function TableAccess_GetUpdateValues() As Collection
    Set TableAccess_GetUpdateValues = m_clsTableAccess.GetUpdateValues
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
    m_clsTableAccess.SetPopulateType enumType
End Sub

Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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

Public Function TableAccess_GetUpdateFields() As Collection
    Dim colKeys As New Collection
    
    colKeys.Add "BaseRateSet"
    colKeys.Add "BaseRateBandStartDate"

    Set TableAccess_GetUpdateFields = colKeys
End Function

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetMaxLoanAmount()
    colListLine.Add GetMaxLTV()
    colListLine.Add GetRateDifference()

End Sub

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    'Modified by Sunil 07-03-2002
    'Modified by Sunil 06-02-2002
    'In order to check whether mortgage products exist for the selected
    'base rate
    On Error GoTo Failed
    Dim sMessage As String
    Dim inum As Integer
    Dim colValues  As New Collection
    Set colValues = m_clsTableAccess.GetKeyMatchValues
    
    Dim clsMortgageProductTable As MortgageProductTable
    Set clsMortgageProductTable = New MortgageProductTable
    
    Set colValues = clsMortgageProductTable.DoesProductExistForBaseRate(colValues.Item(1))
    
    If (Not colValues Is Nothing) And (colValues.Count > 0) Then
        sMessage = ""
        For inum = 1 To colValues.Count
            sMessage = sMessage & "(" & inum & ") " & colValues.Item(inum) & vbCrLf
        Next
        
        MsgBox "Record can not be deleted. The following Mortgage Products exist for the selected Base rate set : " & vbCrLf & sMessage, vbCritical, "Cannot Delete"
        g_clsErrorHandling.RaiseError , "Record can not be deleted. The following Mortgage Products exist for the selected Base rate set : " & vbCrLf & sMessage, ErrorSystem
    Else
        m_clsTableAccess.DeleteRecords sDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteBaseRates()
    On Error GoTo Failed
    ' First, delete the admin fee values
    m_clsTableAccess.DeleteRecords
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

