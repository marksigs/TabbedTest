VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ComboUtils
' Description   : Contains helper functions for access to Omiga 4 combos.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_clsTableAccess As ComboValueTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History:
'
'Prog   Date        AQR     Description
'MV     19/12/2002  BM0085  Added GetValueIDFromValueName
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsTableAccess = New ComboValueTable
End Sub

Friend Sub FindComboGroup(sComboGroup As String, colValues As Collection, colIDS As Collection, Optional sValidationType As String, Optional bExcludeValidationType As Boolean = False, Optional sOrderBy As String, Optional bCheckForValidationType As Boolean = False)
    Dim sSearch As String
    Dim group As ComboGroup
    Dim clsComboTable As ComboValueTable
    Dim sTable As String
    Dim rs As ADODB.Recordset
    Dim sValue As String
    sTable = m_clsTableAccess.GetTableName()
    On Error GoTo Failed
    
    Set colValues = New Collection
    Set colIDS = New Collection
    
    sSearch = "SELECT " & _
                    "ValueName Value, " & _
                    "ValueID ID " & _
                "FROM " & _
                    sTable & _
                " WHERE GroupName = " & g_clsSQLAssistSP.FormatString(sComboGroup)
           
    If Len(sValidationType) > 0 Then
        If bExcludeValidationType Then
            sSearch = sSearch & _
                " AND VALUEID IN ( SELECT VALUEID FROM COMBOVALIDATION " & _
                "WHERE GROUPNAME = " & g_clsSQLAssistSP.FormatString(sComboGroup) & _
                " AND VALIDATIONTYPE = " & g_clsSQLAssistSP.FormatString(sValidationType) & ")"
        Else
            sSearch = sSearch & _
                " AND VALUEID NOT IN ( SELECT VALUEID FROM COMBOVALIDATION " & _
                "WHERE GROUPNAME = " & g_clsSQLAssistSP.FormatString(sComboGroup) & _
                " AND VALIDATIONTYPE = " & g_clsSQLAssistSP.FormatString(sValidationType) & ")"
        End If
        
    End If
        
    If Len(sOrderBy) > 0 Then
        sSearch = sSearch & _
            " ORDER BY " & sOrderBy
    End If
    
    Set rs = g_clsDataAccess.GetTableData(sTable, sSearch)
    
    ValidateRecordset rs, "ComboValue"
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            colValues.Add rs("Value").Value
            colIDS.Add rs("ID").Value
            rs.MoveNext
        Loop
    ElseIf bCheckForValidationType = True Then
        g_clsErrorHandling.RaiseError errGeneralError, "FindComboGroup: For this function to operate a validation type of " + sValidationType + " must be present in the combogroup " + sComboGroup
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "FindComboGroup: Unable to locate group - " + sComboGroup
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetValueNameFromValueID(sValueID As String, sGroup As String) As String
    
    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vValueID As Variant
    Dim sTableName As String
    
    sTableName = m_clsTableAccess.GetTableName()
    
    sSQL = "SELECT " & _
                "VALUENAME " & _
            "FROM " & _
                sTableName & " " & _
            "WHERE " & _
                "GROUPNAME = " & g_clsSQLAssistSP.FormatString(sGroup) & " " & _
                "AND ValueID = " & sValueID
           
    Set rs = g_clsDataAccess.GetTableData(sTableName, sSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            vValueID = rs.fields("ValueName")
        End If
    End If
    
    GetValueNameFromValueID = CStr(vValueID)
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Public Function GetValueIDFromValueName(sGroup As String, sValueName As String) As Integer
    
    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vValueID As Variant
    Dim sTableName As String
    
    sTableName = m_clsTableAccess.GetTableName()
    
    sSQL = "SELECT " & _
                "VALUEID " & _
            "FROM " & _
                sTableName & " " & _
            "WHERE " & _
                "GROUPNAME = " & g_clsSQLAssistSP.FormatString(sGroup) & " " & _
                "AND ValueName = '" & sValueName & "' "
           
    Set rs = g_clsDataAccess.GetTableData(sTableName, sSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            vValueID = rs.fields("ValueId")
        End If
    End If
    
    GetValueIDFromValueName = CStr(vValueID)
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function




