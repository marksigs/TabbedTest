VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IADOAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum DBENGINETYPE
    Undefined
    SQLServer
    Oracle
End Enum

Enum enumLockType
    lckShare
    lckUpdate
    lckExclusive
End Enum

Public Function CheckRecordExists(ByVal strTableName As String, ByVal strKeys As String) As Boolean

End Function

Public Function ExecuteSQLCommand( _
    ByVal vstrSQL As String, _
    Optional ByVal vblnRetry As Boolean = True) As Long
    
End Function

Public Function GetConnStr(Optional strLocation As String = "") As String

End Function

Public Function GetDBEngine(Optional strLocation As String = "") As DBENGINETYPE

End Function

Public Function GetFieldTypes(vstrTableName As String) As ADODB.Recordset

End Function

Public Function GetNumberOfRecords( _
    ByVal vstrTableName As String, ByVal vstrCriteria As String) As Integer
    
End Function

Public Function GetValueFromTable(ByVal strTableName As String, _
                                  ByVal strCondition As String, _
                                  ByVal strColumn As String, _
                                  Optional ByRef varReturnValue As Variant, _
                                  Optional ByRef blnRecordFound As Boolean = True) As Variant
                                  
End Function

Public Function IsThisUniqueKeyConstraint(intErrNumber As Integer) As Boolean

End Function

Public Function IsUniqueKeyConstraint() As Boolean

End Function

Public Sub LockTable(strTableName As String, LockType As enumLockType)

End Sub

