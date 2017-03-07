VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BandedGlobalParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess
Implements BandedTable

Private Const m_sTableName = "GlobalBandedParameter"
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private m_sSearch As String
'Private m_sName As String
'Private m_sStartDate As String
'Private m_sDescription As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Sub SetUpdateValues(sName As String, sStartDate As String, sDescription)
'    m_sName = sName
'    m_sStartDate = sStartDate
'    m_sDescription = sDescription
'End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("Name")
    GetName = CStr(vVal)
End Function
Friend Function GetStartDate() As String
    Dim vVal As Variant

    vVal = m_clsTableAccess.GetVal("StartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetDescription() As String
    GetDescription = CStr(m_clsTableAccess.GetVal("Description"))
End Function

Friend Function GetAmount() As String
    GetAmount = CStr(m_clsTableAccess.GetVal("Amount"))
End Function

Friend Function GetPercentage() As String
    GetPercentage = CStr(m_clsTableAccess.GetVal("Percentage"))
End Function

Friend Function GetMaximum() As String
    GetMaximum = CStr(m_clsTableAccess.GetVal("Maximum"))
End Function

Friend Function GetBoolean() As String
    GetBoolean = CStr(m_clsTableAccess.GetVal("Boolean"))
End Function

Friend Function GetString() As String
    GetString = CStr(m_clsTableAccess.GetVal("String"))
End Function

Friend Function GetToMaxValue() As String
    'GetToMaxValue = CStr(g_clsTableAccess.GetVal("ToMaxValue"))
    GetToMaxValue = CStr(m_clsTableAccess.GetVal("HighBand"))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetName(sName As String) As Boolean
    SetName = m_clsTableAccess.SetVal("Name", CVar(sName))
End Function

Friend Function SetStartDate(sStartDate As String) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("StartDate", CVar(sStartDate))
End Function

Friend Function SetDescription(sDescription As String) As Boolean
    SetDescription = m_clsTableAccess.SetVal("Description", CVar(sDescription))
End Function

Friend Function SetAmount(sAmount As String) As Boolean
    SetAmount = m_clsTableAccess.SetVal("Amount", CVar(sAmount))
End Function

Friend Function SetPercentage(sPercentage As String) As Boolean
    SetPercentage = m_clsTableAccess.SetVal("Percentage", CVar(sPercentage))
End Function

Friend Function SetMaximum(sMaximum As String) As Boolean
    SetMaximum = m_clsTableAccess.SetVal("Maximum", CVar(sMaximum))
End Function
Friend Function SetBoolean(sBoolean As String) As Boolean
    SetBoolean = m_clsTableAccess.SetVal("Boolean", CVar(sBoolean))
End Function
Friend Function SetString(sString As String) As Boolean
    SetString = m_clsTableAccess.SetVal("String", CVar(sString))
End Function
' This has to change when the database field changes.
Friend Function SetToMaxValue(sMax As String) As Boolean
    'SetToMaxValue = g_clsTableAccess.SetVal("ToMaxValue", CVar(sMax))
    SetToMaxValue = m_clsTableAccess.SetVal("HighBand", CVar(sMax))
    SetToMaxValue = m_clsTableAccess.SetVal("LowBand", CVar(sMax))
End Function

Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub
Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub

Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
        
    m_clsTableAccess.Create m_sTableName

    colFields.Add "Name"
    colFields.Add "StartDate"
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "Name"
    colUpdateFields.Add "StartDate"
    colUpdateFields.Add "Description"
    m_clsBandedTable.SetUpdateFields colUpdateFields
End Sub
Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetName()
    colHeader.Add GetDescription()
    colHeader.Add GetStartDate()
End Sub
Public Sub SetPopulateNull()
    m_sSearch = "Select * from " + m_sTableName + " where Name is null"
End Sub
Public Function SetPopulateBand(sName As String, sStartDate As String) As Boolean
    Dim bRet As Boolean
    
    bRet = False
    If Len(sName) > 0 And IsDateValid(sStartDate) = True Then
        m_sSearch = "select * from " + m_sTableName + " Where Name = '" + sName + "' AND to_char(StartDate,'dd/mm/yyyy') = '" + sStartDate + "' order by HighBand"
        bRet = True
    Else
        MsgBox "SetPopulateBand: Name and StartDate must be valid"
        m_sSearch = ""
    End If
    SetPopulateBand = bRet
End Function
Public Sub SetPopulateType(enumType As PopulateType)
    m_clsTableAccess.SetPopulateType enumType
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_AddRow() As Boolean
    TableAccess_AddRow = m_clsTableAccess.AddRow()
End Function
Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub

Private Sub TableAccess_Create(sTable As String)
    m_clsTableAccess.Create sTable
End Sub

Private Function TableAccess_DeleteRow(colMatchData As Collection) As Boolean
    TableAccess_DeleteRow = m_clsTableAccess.DeleteRow(colMatchData)
End Function
Private Function TableAccess_DoesRecordExist(colMatchData As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData)
End Function
Private Function BandedTable_DoUpdateSets() As Boolean
    BandedTable_DoUpdateSets = m_clsBandedTable.DoUpdateSets()
End Function

Private Function TableAccess_GetKeyMatchFields()
    TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function

Private Function TableAccess_GetKeyMatchValues() As Variant
    TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function

Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function

Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS) As ADODB.Recordset
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

Private Function TableAccess_IsDataValid() As Boolean
    TableAccess_IsDataValid = m_clsTableAccess.IsDataValid()
End Function

Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function
Private Function TableAccess_MoveFirst() As Boolean
    TableAccess_MoveFirst = m_clsTableAccess.MoveFirst()
End Function
Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
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

Private Sub TableAccess_SetUpdated()
    m_clsTableAccess.SetUpdated
End Sub
Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vVal)
End Function
Private Function TableAccess_Update() As Boolean
    TableAccess_Update = m_clsTableAccess.Update()
End Function
Private Sub TableAccess_SetPopulateNull()
    m_clsTableAccess.SetPopulateNull
End Sub
Private Function TableAccess_GetSearch() As String
    TableAccess_GetSearch = m_sSearch
End Function
Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function
'Public Function TableAccess_GetUpdateValues() As Collection
'    Dim colValues As New Collection
'
'    colValues.Add m_sName
'    colValues.Add m_sStartDate
'    colValues.Add m_sDescription
'    Set TableAccess_GetUpdateValues = colValues
'End Function
Public Function TableAccess_GetUpdateFields() As Collection
    Dim colKeys As New Collection
    
    colKeys.Add "Name"
    colKeys.Add "StartDate"
    colKeys.Add "Description"
    Set TableAccess_GetUpdateFields = colKeys
End Function
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetToMaxValue()
    colListLine.Add GetAmount()
    colListLine.Add GetPercentage()
    colListLine.Add GetMaximum()
    colListLine.Add GetBoolean()
    colListLine.Add GetString()
End Sub




