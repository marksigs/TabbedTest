VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeleteObjectCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : DeleteObjectCS
' Description   : Contains client specific code for handling deleting items
' Change history
' Prog      Date        Description
' DJP       19/02/02    SYS4148 Created
' DJP       24/02/02    SYS2936 Add code for Provinces/Tax Authorities.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CanObjectBeDeleted
' Description   : Given the object name passed in, determine if it can be deleted. This is a client
'                 specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CanObjectBeDeleted(sObjectName As String) As Boolean
    On Error GoTo Failed
    Dim bCanBeDeleted As Boolean
    bCanBeDeleted = False
    
    Select Case sObjectName

    End Select
    
    CanObjectBeDeleted = bCanBeDeleted
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DeleteObject
' Description   : Given the object name passed in, delete the record from the database. This is a client
'                 specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteObject(sObjectName As String, _
                             clsTableAccess As TableAccess, _
                             colMatchValues As Collection, _
                             Optional sTargetDatabase As String, _
                             Optional bRecordChange As Boolean = True, _
                             Optional sObjectDescription As String) As Boolean
    On Error GoTo Failed
    Dim bObjectDeleted As Boolean
    
    bObjectDeleted = True
    
    Select Case sObjectName
    
    Case Else
        bObjectDeleted = False
    
    End Select
    
    DeleteObject = bObjectDeleted
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


