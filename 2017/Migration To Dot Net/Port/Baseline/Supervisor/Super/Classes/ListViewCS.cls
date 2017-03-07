VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListViewCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ListViewCS
' Description   : Contains client specific code for handling ListView functionality.
' Change history
' Prog      Date        Description
' DJP       19/02/02    SYS4142 Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Given the key passed in, Populate the listview lvListView with the associated
'                 data from the database. This is a client specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PopulateListView(vKey As Variant, lvListView As MSGListView) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    PopulateListView = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ListEdit
' Description   : Given the key passed in, set the form used to edit the data represented by that
'                 key. This is a client specific version of the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ListEdit(vKey As Variant, frmEdit As Form)
    On Error GoTo Failed
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleListViewPopup
' Description   : When the user right clicks the list view, this method will popup a context
'                 menu for client specific data. This is a client specific version of the Core method.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleListViewPopup(vKey As Variant)
    On Error GoTo Failed
    
    Select Case vKey

    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
