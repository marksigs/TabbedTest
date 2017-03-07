VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTFeeBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
' Class         : TTFeeBand
' Description   : Contains all access routines and SQL for the TTFeeBand table, which is a
'                 banded table. This class implements both TableAccess for generic routines,
'                 and BandedTable for Banded specific functionality.
' Change history
' Initials      Prog        Date        Description
'
'   MC          BMIDS763    01/06/2004  ADDED
'
'****************************************************************************************************

Option Explicit

'*=[MC]INTERFACE IMPLEMENTATION
Implements TableAccess
Implements BandedTable

'*=[MC]CONSTANTS & MEMBER VARIABLES
Private Const m_sTableName          As String = "TTFEEBAND"
Private Const m_sVersionField       As String = "AFSETVERSIONNUMBER"

Private m_sSearch                   As String
Private m_sFirstMatch               As String
Private m_bVersionExists            As Boolean
Private m_sProdSwitFeeStartDateFld  As String

Private m_clsBandedTable            As BandedTable
Private m_clsTableAccess            As TableAccess

'****************************************************************************************************
' Get Access Functions
'****************************************************************************************************

Friend Function GetFeeSet() As String
    Dim vVal As Variant
    '*=[MC]GET TTFeeSet ID
    vVal = m_clsTableAccess.GetVal("TTFeeSet")
    '*=RETURN SETID
    GetFeeSet = CStr(vVal)
End Function

Friend Function GetStartDate() As String
    
    Dim vVal As Variant
    '*=[MC]GET PRODUCT SWITCH FEE START DATE
    vVal = m_clsTableAccess.GetVal(m_sProdSwitFeeStartDateFld)
    '*=If Data found, return date
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetTypeOfApplication() As String
    GetTypeOfApplication = CStr(m_clsTableAccess.GetVal("TypeOfApplication"))
End Function

Friend Function GetTypeOfApplicationText() As String
    GetTypeOfApplicationText = CStr(m_clsTableAccess.GetVal("TypeOfApplicationText"))
End Function

Friend Function GetLocationText() As String
    GetLocationText = CStr(m_clsTableAccess.GetVal("LocationText"))
End Function

Friend Function GetLocation() As String
    GetLocation = CStr(m_clsTableAccess.GetVal("Location"))
End Function

Friend Function GetAmount() As String
    GetAmount = CStr(m_clsTableAccess.GetVal("Amount"))
End Function

Friend Function SetFeeSet(sFeeSet As String) As Boolean
    SetFeeSet = m_clsTableAccess.SetVal("TTFeeSet", CVar(sFeeSet))
    
    g_clsVersion.HandleVersioning Me
End Function

Friend Function SetStartDate(sStartDate As String) As Boolean
    SetStartDate = m_clsTableAccess.SetVal(m_sProdSwitFeeStartDateFld, CVar(sStartDate))
End Function

Friend Function SetTypeOfApplication(sTypeOfApplication As String) As Boolean
    SetTypeOfApplication = m_clsTableAccess.SetVal("TypeOfApplication", CVar(sTypeOfApplication))
End Function

Friend Function SetTypeOfApplicationText(sTypeOfApplicationText As String) As Boolean
    SetTypeOfApplicationText = m_clsTableAccess.SetVal("TypeOfApplicationText", CVar(sTypeOfApplicationText))
End Function

Friend Function SetLocationText(sLocationText As String) As Boolean
    SetLocationText = m_clsTableAccess.SetVal("LocationText", CVar(sLocationText))
End Function

Friend Function SetLocation(sLocation As String) As Boolean
    SetLocation = m_clsTableAccess.SetVal("Location", CVar(sLocation))
End Function

Friend Function SetAmount(sAmount As String) As Boolean
    SetAmount = m_clsTableAccess.SetVal("Amount", CVar(sAmount))
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

Public Function GetStartDateField() As String
    GetStartDateField = m_sProdSwitFeeStartDateFld
End Function

Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields       As Collection
    Dim colUpdateFields As Collection
    Dim colDupKeys      As Collection
    Dim bVersionExists  As Boolean
    
    Const sProdSwitchFeeField As String = "STARTDATE"
    Const sProductSwitchFeeField As String = "STARTDATE"
    
    Set colFields = New Collection
    Set colUpdateFields = New Collection
    
    '*=Check for the version field
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    '*=check for the StartDate field
    If g_clsDataAccess.DoesFieldExist(m_sTableName, sProductSwitchFeeField) Then
        m_sProdSwitFeeStartDateFld = sProductSwitchFeeField
    ElseIf g_clsDataAccess.DoesFieldExist(m_sTableName, sProdSwitchFeeField, True) Then
        m_sProdSwitFeeStartDateFld = sProdSwitchFeeField
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to find StartDate field for " + m_clsTableAccess.GetTable
    End If
    '*=Construct SQL with Lock Hint (NoLock)
    'SR 16/07/2004 : BMIDS797 - Modify the SQL to be run
    m_sFirstMatch = "SELECT * FROM TTFeeBand A WITH (NOLOCK) WHERE A.location = " + _
                    "(SELECT MIN(location) FROM  TTFeeBand B  WITH(NOLOCK)  " + _
                    "WHERE A.TTFeeSet = B.TTFeeSet AND A." + m_sProdSwitFeeStartDateFld + " = B." + m_sProdSwitFeeStartDateFld + ") AND A.Typeofapplication = " + _
                    "(SELECT MIN(typeofapplication) FROM TTFeeBand C  WITH (NOLOCK) " + _
                    "WHERE A.TTFeeSet = C.TTFeeSet AND A." + m_sProdSwitFeeStartDateFld + " = C." + m_sProdSwitFeeStartDateFld + _
                     " AND A.Location = C.Location" + ")"  'SR 16/07/2004 : BMIDS797
    '*=If Version field exist, Add
    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If
    '*=Create Instance of TableAccess & Banded Table objects
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    '*=[MC]Create PRODUCTSWITCHFEEBAND DB Table object
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
    '*=Add DB Table PRODUCTSWITCHFEEBAND columns to the collection
    colFields.Add "TTFeeSet"
    colFields.Add m_sProdSwitFeeStartDateFld
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    '*=[MC]Set columns collection to the DB
    m_clsTableAccess.SetKeyMatchFields colFields
    '*=Set to the Banded Table as well
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    '*=[MC]Create update key columns
    colUpdateFields.Add "TTFeeSet"
    colUpdateFields.Add m_sProdSwitFeeStartDateFld
    '*=[MC]Add Key columns to the update collection
    m_clsBandedTable.SetUpdateFields colUpdateFields

    Set colDupKeys = New Collection
    colDupKeys.Add "LOCATION"
    colDupKeys.Add "TYPEOFAPPLICATION"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetHeaderData(colHeader As Collection)
    '*=Add Header columns
    Call colHeader.Add(GetFeeSet())
    Call colHeader.Add(GetStartDate())
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
    
    colKeys.Add "TTFeeSet"
    colKeys.Add m_sProdSwitFeeStartDateFld

    Set TableAccess_GetUpdateFields = colKeys
End Function

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetTypeOfApplicationText()
    colListLine.Add GetLocationText()
    colListLine.Add GetAmount()

End Sub

Public Sub SetFindNewestAdminFeeSet()
    Dim sSysDate As String
    
    sSysDate = g_clsSQLAssistSP.GetSystemDate

    m_sFirstMatch = "SELECT " & _
                        "AF1.* " & _
                    "FROM " & _
                        "TTFeeBand AF1 with (nolock)" & _
                    "WHERE " & _
                        m_sProdSwitFeeStartDateFld & " IN " + _
                        "(SELECT " & _
                            "MAX(" + m_sProdSwitFeeStartDateFld + ") " & _
                        "FROM " & _
                            "TTFeeBand AF2  with (nolock)" & _
                        "WHERE " & _
                            "AF1.TTFeeSet = AF2.TTFeeSet AND " & _
                            g_clsSQLAssistSP.TruncateDateColumn(sSysDate) & " >= AF2. " & m_sProdSwitFeeStartDateFld & " AND " & _
                            "AF1.LOCATION = "
    
    m_sFirstMatch = m_sFirstMatch & "(SELECT " & _
                            "MIN(LOCATION) " & _
                        "FROM " & _
                            "TTFeeBand AF3 with (nolock)" & _
                        "WHERE " & _
                            "AF1.TTFeeSet = AF3.TTFeeSet AND " & _
                            "AF1." & m_sProdSwitFeeStartDateFld & " = AF3." & m_sProdSwitFeeStartDateFld & ") AND " & _
                            "AF1.TYPEOFAPPLICATION = " & _
                            "(SELECT " & _
                                "MIN(TYPEOFAPPLICATION) " & _
                            "FROM " & _
                                "TTFeeBand AF4 with (nolock)" & _
                            "WHERE " & _
                                "AF1.TTFeeSet = AF4.TTFeeSet AND " & _
                                "AF1." & m_sProdSwitFeeStartDateFld & " = AF4." & m_sProdSwitFeeStartDateFld & "))"

    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If
    
    '*=Create Table
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    
    Dim sMessage                As String
    Dim inum                    As Integer
    Dim colValues               As Collection
    Dim clsMortgageProductTable As MortgageProductTable
    
    On Error GoTo ErrorHandler
    
    '*=Fields Collection
    Set colValues = New Collection
    Set colValues = m_clsTableAccess.GetKeyMatchValues
    
    '*=Initialize Mortgage Products Table
    Set clsMortgageProductTable = New MortgageProductTable
    Set colValues = clsMortgageProductTable.DoesProductExistForThisItem("TTFeeSet", colValues.Item(1), sDatabase)
    '*=Check columns and take appropriate action
    If (Not colValues Is Nothing) And (colValues.Count > 0) Then
        sMessage = ""
        For inum = 1 To colValues.Count
            sMessage = sMessage & "(" & inum & ") " & colValues.Item(inum) & vbCrLf
        Next
        
        MsgBox "Record can not be deleted. The following Mortgage Products exist for the selected Admin Fee Set : " & vbCrLf & sMessage, vbCritical, "Cannot Delete"
        g_clsErrorHandling.RaiseError , "Record can not be deleted. The following Mortgage Products exist for the selected Admin Fee Set : " & vbCrLf & sMessage, ErrorSystem
    Else
        '*=Delete Records from the given database
        m_clsTableAccess.DeleteRecords sDatabase
    End If
    
ExitHandler:

    '*=[MC]Cleanup
    Exit Sub
    
ErrorHandler:
    '*=[MC]Handle Errors/ Log to file
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteProductSwitchFeeFees()
    On Error GoTo ErrorHandler
    
    '*=Delete product switch fee records
    m_clsTableAccess.DeleteRecords

ExitHandler:
    '*=Cleanup Code here
    
    Exit Sub
ErrorHandler:
    '*=[MC]Handle Errors
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'***********************************************************************************
'
'   INTERFACE IMPLEMENTATION METHODS
'
'***********************************************************************************

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





