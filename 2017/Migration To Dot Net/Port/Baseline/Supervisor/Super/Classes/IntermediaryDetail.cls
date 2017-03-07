VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntermediaryDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : IntermediaryDetail
' Description   :
'
' Change history
' Prog      Date        Description
' STB       15/11/2001  SYS2550: SQL Server support.
' STB       26/11/2001  SYS2912: SQL Server locking issue.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private m_clsTreeItem As TreeItem


Public Sub HandleIntermediaries(nIndex As Integer)
    
    Dim nIntermediaryType As Integer
    Dim sGroupName As String
    Dim clsComboValidation As ComboValidationTable
    Dim frmEdit As Form
    Dim clsTreeItem As TreeItem
    Dim sKey As String
    Dim uReturn As MSGReturnCode
    
    On Error GoTo Failed
    
    Set clsTreeItem = g_clsMainSupport.GetIntermediaryTreeItem
    Set clsComboValidation = New ComboValidationTable
    
    sGroupName = "IntermediaryType"
    
    Select Case nIndex
    
        Case LV_INTERMEDIARY_ADMIN
            nIntermediaryType = clsComboValidation.GetSingleValueFromValidation(sGroupName, INTERMEDIARIES_ADMIN)
        Case LV_INTERMEDIARY_INDIVIDUAL
            nIntermediaryType = clsComboValidation.GetSingleValueFromValidation(sGroupName, INTERMEDIARIES_INDIVIDUAL)
        Case LV_INTERMEDIARY_COMPANY
            nIntermediaryType = clsComboValidation.GetSingleValueFromValidation(sGroupName, INTERMEDIARIES_COMPANY)
        Case LV_INTERMEDIARY_LEADAGENT
            nIntermediaryType = clsComboValidation.GetSingleValueFromValidation(sGroupName, INTERMEDIARIES_LEADAGENT)
    End Select
    
    
    Set frmEdit = frmEditIntermediaries
    frmEdit.SetIntermediaryType nIntermediaryType
    sKey = clsTreeItem.GetTag
    frmEdit.SetParentKey FormatKey(sKey)
    frmEdit.SetIsEdit False
    'frmEdit.cboType.Enabled = False
    
    'Begin a transaction.
    g_clsDataAccess.BeginTrans
    
    'Show the main intermediary form.
    frmEdit.Show vbModal
        
    'Get the return code from the form.
    uReturn = frmEdit.GetReturnCode
    
    'frmEdit.Hide
    g_clsMainSupport.HandleReturn uReturn
        
    'Unload the form now.
    Unload frmEdit

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
    

Public Sub SetTreeItem(clsTreeItem As TreeItem)
    Set m_clsTreeItem = clsTreeItem
End Sub
Public Sub GetIntermediaries(lvListView As MSGListView, clsTreeItem As TreeItem, tvNode As MSComctlLib.node)
    On Error GoTo Failed
    
    Dim clsIntermediary As Intermediary

    Set clsIntermediary = New Intermediary

    clsIntermediary.GetIntermediary clsTreeItem, lvListView, tvNode
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateSingleIntermediary(tv As MSComctlLib.TreeView, node As MSComctlLib.node, clsTreeItem As TreeItem)
    On Error GoTo Failed
    Dim clsIntermediary As Intermediary
    
    Set clsIntermediary = New Intermediary
    
    clsIntermediary.PopulateSingleIntermediary tv, node, clsTreeItem
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
Public Sub PopulateIntermediaries(tv As MSComctlLib.TreeView, node As MSComctlLib.node, sType As String)
    On Error GoTo Failed
    Dim clsComboValidation As ComboValidationTable
    Dim sComboGroup As String
    Dim nIntType As Integer
    Dim clsIntermediaryTable As IntermediaryTable
    Dim clsIntermediary As Intermediary
    Dim clsIntDetails As IntermediaryDetail
    
    Set clsIntDetails = New IntermediaryDetail
    Set clsIntermediary = New Intermediary
    Set clsIntermediaryTable = New IntermediaryTable
    Set clsComboValidation = New ComboValidationTable
        
    If node.Text = INTERMEDIARIES Then
         clsIntDetails.HandleIntermediaryRefresh False
    End If
    
    
    tv.Nodes(INDIVIDUAL).Bold = True
    tv.Nodes(LEADAGENT).Bold = True
    
    'Have intermediaries been populated?
    On Error Resume Next
    If tv.Nodes(LEADAGENT).children <= 1 Or tv.Nodes(INDIVIDUAL).children <= 1 Or Err.Number > 0 Then
        On Error GoTo Failed
        'Get the valueID for LeadAgent's
        sComboGroup = "IntermediaryType"
        nIntType = clsComboValidation.GetSingleValueFromValidation(sComboGroup, INTERMEDIARIES_LEADAGENT)

        clsIntermediaryTable.GetIntermediaries CStr(nIntType)

        'Populate all LeadAgents
        clsIntermediary.SetParentNodeKey LEADAGENT
        clsIntermediary.PopulateIntermediary tv, IntTypeLeadAgent, clsIntermediaryTable, nIntType, node

        'Populate all Individuals
        sComboGroup = "IntermediaryType"
        nIntType = clsComboValidation.GetSingleValueFromValidation(sComboGroup, INTERMEDIARIES_INDIVIDUAL)

        clsIntermediaryTable.GetIntermediaries CStr(nIntType)
        clsIntermediary.SetParentNodeKey INDIVIDUAL
        clsIntermediary.PopulateIntermediary tv, IntTypeIndividual, clsIntermediaryTable, nIntType, node
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetTreeItem() As TreeItem
    Set GetTreeItem = m_clsTreeItem
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleInteremediarySearch
' Description   : Display the search Screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleIntermediarySearch()
    On Error GoTo Failed

    Dim frmSearch As frmFindIntermediary
    
    Set frmSearch = frmFindIntermediary
    
    frmSearch.Show vbModal

    Unload frmSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleInteremediaryRefresh
' Description   : Refreshes the intermediary Node
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleIntermediaryRefresh(Optional bPopulate As Boolean = True)
    On Error GoTo Failed

    Dim IntNode As MSComctlLib.node
    Dim tmpNode As MSComctlLib.node
    
    frmMain.ClearCurrentTreeitem

    Set IntNode = frmMain.tvwDB.Nodes(INTERMEDIARIES)
    
    If IntNode Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Intermediary Node not found"
    Else
        'remove all intermediary nodes
        On Error Resume Next
        Set tmpNode = frmMain.tvwDB.Nodes(LEADAGENT)
        If Not tmpNode Is Nothing Then
            frmMain.tvwDB.Nodes.Remove (LEADAGENT)
            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, LEADAGENT, LEADAGENT, "closed"
        End If

        Set tmpNode = frmMain.tvwDB.Nodes(INDIVIDUAL)

        If Not tmpNode Is Nothing Then
            frmMain.tvwDB.Nodes.Remove (INDIVIDUAL)
            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, INDIVIDUAL, INDIVIDUAL, "closed"
        End If

        On Error GoTo Failed
        If bPopulate Then
            PopulateIntermediaries frmMain.tvwDB, IntNode, INTERMEDIARIES
        End If
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FormatKey
' Description   : Filters the type out of the key if present
'                 when the user adds an intermediary from a type node (ie admin centre) the key contains Admin Centre
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatKey(sKey As String) As String
    On Error GoTo Failed

    Dim vType As Variant
    Dim col As Collection
    Dim nFound As Long
    
    Set col = New Collection
    col.Add LEADAGENT
    col.Add INDIVIDUAL
    col.Add INTERMEDIARY_COMPANY
    col.Add ADMIN_CENTRE
    
    For Each vType In col
        nFound = InStr(1, sKey, vType)
        If nFound > 0 Then
            sKey = Left(sKey, nFound - 1)
        End If
    Next

    FormatKey = sKey
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
