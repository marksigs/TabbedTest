VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PasswordData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PasswordData
' Description   : Saves and retrieves password text for added security in frmDatabaseOptions
'                 screen.
'
' History
' Prog  Date        Description
' JR    30/05/2002  SYS4816 Created - Ensures password text is saved/retrieved and not the asterix's
' DB    15/11/2002  BMIDS00851 Ensures user id text is saved/retrieved and not the asterix's
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sPassword As String
Private m_sUserID As String 'DB BMIDS00851

Public Sub SetPassword(sPassword As String)
    m_sPassword = sPassword
End Sub
Public Function GetPassword() As String
    GetPassword = m_sPassword
End Function

'
' Procedure ReplaceCollectionForPassword
' Author:   JREID
' Date:     31/05/2002
' Purpose:  Works with frmDatabaseOptions screen by passing in a collection and replacing
'           with a new collection but showing password as Text and not as asterix's.
' Result:
' Input parameters: Collection to be replaced (which displays password as asterix)
' Output parameters:New collection with password displayed as text'
'------------------------------------------------------------------------
Public Function ReplaceCollectionForPassword(OldCol As Collection) As Collection
    
    On Error GoTo Failed
    
    Dim iCounter As Integer
    Dim NewCol As Collection
    Set NewCol = New Collection
    
    If Not OldCol Is Nothing Then
        For iCounter = 1 To OldCol.Count
            If iCounter = COL_PASSWORD Then
                NewCol.Add m_sPassword
            ElseIf iCounter = COL_USER_ID Then 'DB BMIDS00851
                NewCol.Add m_sUserID 'DB BMIDS00851
            Else
                NewCol.Add OldCol(iCounter)
            End If
        Next
    End If
    
    Set ReplaceCollectionForPassword = NewCol
    
    Exit Function
        
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'DB BMIDS00851
Public Sub SetUserID(sUserID As String)
    m_sUserID = sUserID
End Sub
'DB BMIDS00851
Public Function GetUserID() As String
    GetUserID = m_sUserID
End Function

