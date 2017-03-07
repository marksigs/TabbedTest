VERSION 5.00
Object = "{1EB14087-1826-44AD-A7E5-2292F4882971}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmAdditionalFeatures 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Features"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   Icon            =   "frmAdditionalFeatures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFormAction 
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   6060
      TabIndex        =   3
      Top             =   3480
      Width           =   1485
   End
   Begin VB.CommandButton cmdFormAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   4500
      TabIndex        =   2
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Frame fmeAF 
      Height          =   3435
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   7545
      Begin MSGOCX.MSGHorizontalSwapList MSGSwapAddlItems 
         Height          =   3225
         Left            =   165
         TabIndex        =   1
         Top             =   135
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   5689
      End
   End
End
Attribute VB_Name = "frmAdditionalFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' Form          : frmAdditionalFeatures.frm
' Description   : This form show available,selected lists for Additonal features
'                 for the Mortgage products
'
' HISTORY
' Prog      Date        Description
' MC        25/05/2004  Created
'
'*******************************************************************************

Option Explicit

Private m_clsAddtFeaTable           As MortProdAddtFeaTable
Private m_clsComboTable             As ComboValueTable
Private m_colKeys                   As Collection

Private m_bIsEdit                   As Boolean
Private m_uReturnCode               As MSGReturnCode

' Constants
Private Const m_sComboAdditionalFeatures As String = "AdditionalFeatures"

'*********************************************************************************
' Function      : SetTableClass
' Description   : Associate the underlying table from the parent form with this
'                 form.
'*********************************************************************************

Public Sub SetTableClass(ByRef clsAdditionalFeaturesTable As MortProdAddtFeaTable)
    Set m_clsAddtFeaTable = clsAdditionalFeaturesTable
End Sub


Private Sub cmdFormAction_Click(Index As Integer)
    On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            '*[MC]Update Form Data
            Call SaveScreenData
            '*=[MC]Set return code success
            SetReturnCode MSGSuccess
            '*=Hide form for later use
            Unload Me
        Case 1
            '*=[MC]Hide form on cancel invoke
            Unload Me
    End Select
    
ExitHandler:
    Exit Sub
    
ErrorHandler:
    '*=[MC]Handle Errors
End Sub

'*********************************************************************************
' Function      : Form_Load
' Description   : This Event sub routinue executes when form loads
'*********************************************************************************
Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Set m_clsComboTable = New ComboValueTable
    
    '*=Set default return code to failure
    SetReturnCode MSGFailure
    
    '*=SetupAdditionalFeatures
    Call SetupAdditionalFeatures
    '*=[MC]Populate AdditionalFeatures swaplists
    Call PopulateScreenFields

ExitHandler:

    Exit Sub

ErrorHandler:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
    
End Sub

'*********************************************************************************
' Function      : PopulateScreenFields
' Description   : Update the GUI controls values from the underlying data.
'*********************************************************************************
Private Sub PopulateScreenFields()
    
    
    If Not TableAccess(m_clsAddtFeaTable).RecordCount > 0 Then
        m_bIsEdit = False
        PopulateAvailableItems
    Else 'populate swap list
        PopulateSwapList
    End If

End Sub


Public Sub SetupAdditionalFeatures()
    
    Dim colHeaders  As Collection
    Dim lvHeaders   As listViewAccess
    
    On Error GoTo ErrorHandler
    
    Set colHeaders = New Collection
    '*=Set ListView Header width
    lvHeaders.nWidth = 95
    '*=[MC] Add column headers to available items list
    lvHeaders.sName = "Additional Features"
    colHeaders.Add lvHeaders
    '*=Set Column headers
    frmAdditionalFeatures.MSGSwapAddlItems.SetFirstColumnHeaders colHeaders
    
    Set colHeaders = New Collection
    '*=[MC] Add column headers to selected item list
    lvHeaders.sName = "Additional Features"
    colHeaders.Add lvHeaders
    frmAdditionalFeatures.MSGSwapAddlItems.SetSecondColumnHeaders colHeaders
    
ExitHandler:
    '*=Cleanup Code
    Exit Sub

ErrorHandler:
    '*=Handle Errors
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub







'*********************************************************************************
' Function      : SetIsEdit
' Description   : Sets whether or not this form is in edit or add mode.
'                 Defaults to edit
'*********************************************************************************
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub


'*********************************************************************************
' Function      : SetReturnCode
' Description   : Sets the return code for the form for any calling method to
'                 check. Defaults to MSG_SUCCESS.
'*********************************************************************************
Private Sub SetReturnCode(ByVal uReturnCode As MSGReturnCode)
    m_uReturnCode = uReturnCode
End Sub


'*********************************************************************************
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'*********************************************************************************
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub



'*********************************************************************************
' Function      : SaveScreenData
' Description   : Update the underlying table object from the screen controls.
'*********************************************************************************
Public Sub SaveScreenData()
    
    On Error GoTo Failed
    
    Dim nSelectedCount As Integer
    Dim nThisItem As Integer
    Dim sValueId As String
    Dim sAmount As String
 
    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess

    nSelectedCount = Me.MSGSwapAddlItems.GetSecondCount()
    Set clsTableAccess = m_clsAddtFeaTable
    
    If clsTableAccess.RecordCount > 0 Then
        clsTableAccess.DeleteAllRows
    Else
        clsTableAccess.GetTableData POPULATE_EMPTY
    End If
    
    For nThisItem = 1 To nSelectedCount
        Set colValues = Me.MSGSwapAddlItems.GetLineSecond(nThisItem, clsSwapExtra)
    
        sValueId = clsSwapExtra.GetValueID()
    
            If Len(sValueId) > 0 Then
                clsTableAccess.AddRow

                Call m_clsAddtFeaTable.SetFeatureID(sValueId)
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
            End If
        Next

        
    Exit Sub
Failed:
        
    g_clsErrorHandling.RaiseError Err.Number, Err.Number
End Sub

Private Sub PopulateSwapList()
    PopulateSelectedItems
    PopulateAvailableItems
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list of items available to move from the first swaplist listview
'                 to the second.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection
    Dim sValue As String
    Dim sValueId As String
    
    Me.MSGSwapAddlItems.ClearFirst
    Call PopulateComboTable(m_sComboAdditionalFeatures)
    
    Set clsTableAccess = m_clsComboTable
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.ValidateData
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sValue = m_clsComboTable.GetValueName()
            sValueId = m_clsComboTable.GetValueID()
            clsSwapExtra.SetValueID sValueId
            
            ' Does this value exist in the Selected Items Listbox?
            If DoesSwapValueExist(sValue, MSGSwapAddlItems) = False Then
                colLine.Add sValue
                frmAdditionalFeatures.MSGSwapAddlItems.AddLineFirst colLine, clsSwapExtra
            End If
            clsTableAccess.MoveNext
        Loop
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim sValue As String
    Dim sValueId As String
    Dim colLine As Collection
    Dim clsComboValues As New ComboValueTable
    Dim clsTableAccess As TableAccess
    Dim rsAdditionalFeatures As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colValues As New Collection
    Dim enumPopulate As PopulateType
    Set clsTableAccess = m_clsAddtFeaTable

    enumPopulate = POPULATE_EMPTY
    
    Me.MSGSwapAddlItems.ClearSecond
    Set clsTableAccess = m_clsAddtFeaTable
       
    Set rsAdditionalFeatures = clsTableAccess.GetRecordSet
    
    If clsTableAccess.RecordCount > 0 Then
        ' Need to read each index, then find the corresponding string from the
        ' combo table.
                    
        If rsAdditionalFeatures.RecordCount > 0 Then
            rsAdditionalFeatures.MoveFirst
            
            Do While Not rsAdditionalFeatures.EOF
                sValueId = m_clsAddtFeaTable.GetFeatureID
                
                If Len(sValueId) > 0 Then
                    sValue = clsComboValues.FindComboValue(m_sComboAdditionalFeatures, sValueId)
                
                    If Len(sValue) > 0 Then
                        Set colLine = New Collection
                        Set clsSwapExtra = New SwapExtra
                        colLine.Add sValue
                        
                        clsSwapExtra.SetValueID sValueId
                        Me.MSGSwapAddlItems.AddLineSecond colLine, clsSwapExtra
                    End If
                Else
                
                End If
                rsAdditionalFeatures.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateComboTable
' Description   : Populates the combo table class with the Combo Group passed in.
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateComboTable(sGroup As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim colKeys As New Collection
    Dim clsTableAccess As TableAccess
    
    Set m_clsComboTable = New ComboValueTable
    Set clsTableAccess = m_clsComboTable
    
    If Len(sGroup) > 0 Then
        colKeys.Add sGroup
        clsTableAccess.SetKeyMatchValues colKeys
        Set rs = clsTableAccess.GetTableData()
    
        clsTableAccess.ValidateData
        If rs.RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - No entries for " + sGroup + " in Combo table"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - Group is empty"
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesSwapValueExist
' Description   : Checks if the value passed in sValue exists in the second listview of the
'                 swaplist msgSwap.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesSwapValueExist(sValue As String, msgSwap As MSGHorizontalSwapList) As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sSecondValue As String
    Dim colValue As Collection
    
    nCount = msgSwap.GetSecondCount()
    
    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        
        Set colValue = msgSwap.GetLineSecond(nThisItem)
        
        If colValue.Count = 1 Then
            sSecondValue = colValue(1)
        
            If sSecondValue = sValue Then
                bFound = True
            End If
        End If
        
        nThisItem = nThisItem + 1
    Wend
    
    DoesSwapValueExist = bFound
End Function

