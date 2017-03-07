VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Error.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     10/01/01    Created
'------------------------------------------------------------------------------------------

Option Explicit

Public m_SnapInData As SnapInData

Const SZ_ERRTITLE = "MessageQueueController Runtime Error"

' =============================================================================
' Method:       DisplayError
' Type:         Subroutine
' Description:  A method to format and display a runtime error
' Parameters:   szLocation      A string identifying the source location
'                               (i.e. method name) where the error occurred
' Output:       None
' Notes:        The error will be displayed in a messagebox formatted as the
'               following sample:
'
'     Method:        SomeMethodName
'     Source:        MMCListSubItems
'     Error:         2527h  (9511)
'     Description:   There is already an item in the collection that has the specified key
'
' =============================================================================
'
Public Sub DisplayError(szLocation As String)
    
    m_SnapInData.SnapIn.ConsoleMsgBox "Method:" & vbTab & vbTab & szLocation & vbCrLf _
        & "Source:" & vbTab & vbTab & Err.Source & vbCrLf _
        & "Error:" & vbTab & vbTab & Hex(Err.Number) & "h   (" & CStr(Err.Number) & ")" & vbCrLf _
        & "Description:" & vbTab & Err.Description, _
        vbCritical, SZ_ERRTITLE
End Sub



