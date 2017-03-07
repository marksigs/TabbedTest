VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupervisorUpdateDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : SupervisorUpdateDetails
' Description   : Data wrapper class for the details Supervisor has to maintain when making a
'                 promotion update. Data is held privately with public access routines
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_colKeyValues As Collection
' DJP SQL Server port
Private m_vObjectID As Variant
Private m_sObjectName As String
Private m_sCreationDate As String
Private m_enumOperation As SupervisorPromoteType
' DJP SQL Server port
Public Sub SetObjectID(vObjectID As Variant)
    m_vObjectID = g_clsSQLAssistSP.GuidStringToByteArray(vObjectID)
End Sub
Public Sub SetKeyMatchValues(colValues As Collection)
    Set m_colKeyValues = colValues
End Sub
Public Sub SetCreationDate(sCreationDate As String)
    m_sCreationDate = sCreationDate
End Sub
Public Sub SetOperation(enumOperation As SupervisorPromoteType)
    m_enumOperation = enumOperation
End Sub
Public Sub SetObjectName(sObjectName As String)
    m_sObjectName = sObjectName
End Sub
Public Function GetObjectName() As String
    GetObjectName = m_sObjectName
End Function
Public Function GetOperation() As SupervisorPromoteType
    GetOperation = m_enumOperation
End Function
' DJP SQL Server port, return variant not string
Public Function GetObjectID() As Variant
    On Error GoTo Failed
    
    If Not IsNull(m_vObjectID) Then
        If Len(m_vObjectID) > 0 Then
            GetObjectID = m_vObjectID
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "GetObjectID: Key values are empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetObjectID: Key values are empty"
    End If
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetKeyMatchValues()
    On Error GoTo Failed
    Dim sFunctionName As String
    
    sFunctionName = "GetKeyMatchValues"
    
    If Not m_colKeyValues Is Nothing Then
        If m_colKeyValues.Count > 0 Then
            Set GetKeyMatchValues = m_colKeyValues
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key values are empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key values are empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetCreationDate() As String
    On Error GoTo Failed
    Dim sFunctionName As String
    
    sFunctionName = "GetCreationDate"
    
    If Len(m_sCreationDate) > 0 Then
        GetCreationDate = m_sCreationDate
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Creation Date is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub Class_Initialize()
    m_enumOperation = PromoteEdit
End Sub
