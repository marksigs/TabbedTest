VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_clsGlobalParameter As New FixedParametersTable
Friend Function FindAmount(sName As String) As Double
    Dim sSearch As String
    Dim sError As String
    Dim sTableName As String
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    FindAmount = 0
    
    Set clsTableAccess = m_clsGlobalParameter
    sTableName = clsTableAccess.GetTable()
    
    sSearch = "select AMOUNT from " + sTableName + " where Name = '" + FormatString(sName) + "' and GlobalParameterStartDate = (select MAX(GlobalParameterStartDate) from " + sTableName + " where Name = '" + FormatString(sName) + "')"

    Set rs = g_clsDataAccess.GetTableData(sTableName, sSearch)
    
    If Not rs Is Nothing Then
        clsTableAccess.SetRecordSet rs
        
        If rs.RecordCount > 0 Then
            FindAmount = CDbl(m_clsGlobalParameter.GetAmount())
        Else
            Err.Raise ErrGlobalParamNotFound, , "Amount Parameter " + sName + " cannot be found"
        End If
    End If
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function


Friend Function FindString(sName As String) As String
    Dim sSearch As String
    Dim sError As String
    Dim sTableName As String
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    FindString = ""
    
    Set clsTableAccess = m_clsGlobalParameter
    sTableName = clsTableAccess.GetTable()
    
    sSearch = "select STRING from " + sTableName + " where Name = '" + FormatString(sName) + "' and GlobalParameterStartDate = (select MAX(GlobalParameterStartDate) from " + sTableName + " where Name = '" + FormatString(sName) + "')"

    Set rs = g_clsDataAccess.GetTableData(sTableName, sSearch)
    
    If Not rs Is Nothing Then
        clsTableAccess.SetRecordSet rs
        
        If rs.RecordCount > 0 Then
            FindString = m_clsGlobalParameter.GetString()
        Else
            Err.Raise ErrGlobalParamNotFound, , "String Parameter " + sName + " cannot be found"
        End If
    End If
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function

Friend Function FindBoolean(sName As String) As Boolean
    Dim sSearch As String
    Dim sError As String
    Dim sTableName As String
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    FindBoolean = False
    
    Set clsTableAccess = m_clsGlobalParameter
    sTableName = clsTableAccess.GetTable()
    
    sSearch = "select BOOLEAN from " + sTableName + " where Name = '" + FormatString(sName) + "' and GlobalParameterStartDate = (select MAX(GlobalParameterStartDate) from " + sTableName + " where Name = '" + FormatString(sName) + "')"

    Set rs = g_clsDataAccess.GetTableData(sTableName, sSearch)
    
    If Not rs Is Nothing Then
        clsTableAccess.SetRecordSet rs
        
        If rs.RecordCount > 0 Then
            FindBoolean = CBool(m_clsGlobalParameter.GetBoolean())
        Else
            Err.Raise ErrGlobalParamNotFound, , "Boolean Parameter " + sName + " cannot be found"
        End If
    End If
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function

