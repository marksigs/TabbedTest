VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMaintainProductExclusivity 
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   Icon            =   "frmMaintainProductExclusivity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Find"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   10095
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   8040
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin MSGOCX.MSGEditBox txtFSARefNo 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGEditBox txtUnitID 
         Height          =   285
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGEditBox txtCompanyName 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGEditBox txtTown 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGEditBox txtPostCode 
         Height          =   285
         Left            =   5640
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FSA Ref No"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company Name "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Town"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Post Code"
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit ID"
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   7680
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8340
      TabIndex        =   7
      Top             =   7680
      Width           =   1275
   End
   Begin VB.TextBox txtSearchType 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Items"
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   10095
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin MSGOCX.MSGListView lvAvailableItems 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Linked Items"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   10095
      Begin VB.Frame Frame4 
         Caption         =   "Proc Fee Loading"
         Height          =   1455
         Left            =   8040
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   1335
         End
         Begin MSGOCX.MSGEditBox txtProcFeeLoading 
            Height          =   285
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            TextType        =   2
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         End
      End
      Begin MSGOCX.MSGListView lvLinkedItems 
         Height          =   2175
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         AllowColumnReorder=   0   'False
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   1
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Legend"
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   24
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   23
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblLegend 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Legend"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   21
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmMaintainProductExclusivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmMaintainProductExclusivity
' Description   : To Add and Edit Product Exclusivity Arrangements
'
' Change history
' Prog      Date        Description
' TW        18/11/2006  EP2_132 ECR20/21 Created
' TW        06/12/2006  EP2_323 - Deal with Exclusive Loading values
' TW        09/02/2007  EP2_1298 - Cannot search using FSA ref in Find Introducers from Mortgage Products
' TW        26/04/2007  EP2_2599 - Run Time Error - From Find Exclusives for Networks
' TW        30/07/2007  VR654 - Code change for linking exclusive products to firms
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const mc_strTHIS_MODULE As String = "frmMaintainProductExclusivity"

' Private data
Private m_clsMortProdExclusivityTable As MortProdExclusivityTable
Private m_clsMortProdLanguageTable As MortProdLanguageTable
Private m_clsMortgageClubNetAssocTable As MortgageClubNetAssocTable
Private m_clsPrincipalFirmTable As PrincipalFirmTable

Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Private strSearch As String
Private strWhere As String

Private intLaunchMode As Integer     'Defines where the screen was launched
Private intType As Integer

' TW 30/07/2007 VR654
Private intExclusivityPackagerIndicator As Integer
' TW 30/07/2007 VR654 End

Private strProductExclusivitySeqNo As String
Private strPrincipalFirmID As String
Private strClubNetworkAssociationID As String
Private strProductCode As String
Private Function LikeOrEquals(strFieldName As String, strValue As String) As String
' TW 09/02/2007 EP2_1298
Dim strComparison As String
    strComparison = strFieldName & " "
    If InStr(1, strValue, "%") > 0 Then
        strComparison = strComparison & " Like '"
    Else
        strComparison = strComparison & " = '"
    End If
    LikeOrEquals = strComparison & strValue & "' "
' TW 09/02/2007 EP2_1298 End
End Function

Private Sub cmdApply_Click()
Dim colMatchValues As Collection
Dim colFields As Collection

Dim clsTableAccess As TableAccess

Dim strSQLRead As String
Dim strSQLUpdate As String
Dim strSQLWhere As String
Dim strSQL As String
Dim conn As ADODB.Connection

Dim rs As ADODB.Recordset

    On Error GoTo Failed
    
    Set colFields = New Collection
    Set colMatchValues = New Collection
    
    Frame4.Visible = False

    Set conn = g_clsDataAccess.GetActiveConnection

    Set clsTableAccess = m_clsMortProdExclusivityTable

    strSQLRead = "Select PRODUCTEXCLUSIVITYSEQNO from MORTGAGEPRODUCTEXCLUSIVITY Where MORTGAGEPRODUCTCODE = '"

    strSQLUpdate = "Update MORTGAGEPRODUCTEXCLUSIVITY Set PROCFEELOADING = " & txtProcFeeLoading.Text & " Where MORTGAGEPRODUCTCODE = '"

    Select Case intLaunchMode
        Case 1 'Associations/Clubs
            strSQLWhere = lvLinkedItems.SelectedItem & "' AND CLUBNETWORKASSOCIATIONID = '" & strClubNetworkAssociationID & "'"
        Case 2 'Principal Firms/Packagers
            strSQLWhere = lvLinkedItems.SelectedItem & "' AND PRINCIPALFIRMID = '" & strPrincipalFirmID & "'"
        Case 3 'Product
            strSQLWhere = strProductCode & "' AND "
            Select Case intType
                Case 1 'PrincipalFirm/Network
                    strSQLWhere = strSQLWhere & "PRINCIPALFIRMID = '" & lvLinkedItems.SelectedItem & "'"
                Case 3 'Club/Association
                    strSQLWhere = strSQLWhere & "CLUBNETWORKASSOCIATIONID = '" & lvLinkedItems.SelectedItem & "'"
            End Select
    End Select
    
    strSQL = strSQLRead & strSQLWhere
    Set rs = conn.Execute(strSQL)
    colFields.Add "PRODUCTEXCLUSIVITYSEQNO"
    colMatchValues.Add rs(0)
    clsTableAccess.SetKeyMatchFields colFields
    clsTableAccess.SetKeyMatchValues colMatchValues

    strSQL = strSQLUpdate & strSQLWhere
    conn.Execute strSQL

    g_clsHandleUpdates.SaveChangeRequest m_clsMortProdExclusivityTable

    Select Case intLaunchMode
        Case 1, 2
            PopulateAvailableProductsListView
            PopulateLinkedProductsListView
        Case 3
            PopulateLinkedIntroducersListView
    End Select
    Set conn = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub Form_Activate()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    Set m_clsMortProdExclusivityTable = New MortProdExclusivityTable
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    PopulateScreenFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Private Sub cmdAdd_Click()
Dim X As Integer
Dim clsTableAccess As TableAccess

' Link Product via Mortgage Product Exclusivity to PrincipalFirm or MortgageClubAssociation
    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        For X = 1 To lvAvailableItems.ListItems.Count
            If lvAvailableItems.ListItems.Item(X).Selected Then
    
                Call GetProductExcusivitySeqNo
                Set m_colKeys = New Collection
                       
                m_colKeys.Add strProductExclusivitySeqNo
                
                g_clsFormProcessing.CreateNewRecord m_clsMortProdExclusivityTable
                Set clsTableAccess = m_clsMortProdExclusivityTable
                m_clsMortProdExclusivityTable.SetProductExclusivitySeqNo strProductExclusivitySeqNo
                Select Case intLaunchMode
                    Case 1 'Associations/Clubs
                        m_clsMortProdExclusivityTable.SetClubNetworkAssociationId strClubNetworkAssociationID
                        m_clsMortProdExclusivityTable.SetMortgageProductCode lvAvailableItems.ListItems.Item(X)
' TW 30/07/2007 VR654
                        m_clsMortProdExclusivityTable.SetPackagerIndicator intExclusivityPackagerIndicator
' TW 30/07/2007 VR654 End
                    Case 2 'Principal Firms/Packagers
                        m_clsMortProdExclusivityTable.SetPrincipalfirmID strPrincipalFirmID
                        m_clsMortProdExclusivityTable.SetMortgageProductCode lvAvailableItems.ListItems.Item(X)
' TW 30/07/2007 VR654
                        m_clsMortProdExclusivityTable.SetPackagerIndicator intExclusivityPackagerIndicator
' TW 30/07/2007 VR654 End
                    Case 3 'Product
                        m_clsMortProdExclusivityTable.SetMortgageProductCode strProductCode
' TW 30/07/2007 VR654
                        m_clsMortProdExclusivityTable.SetPackagerIndicator lvAvailableItems.ListItems.Item(X).SubItems(6)
' TW 30/07/2007 VR654 End
                        Select Case lvAvailableItems.ListItems.Item(X).SubItems(4)
                            Case 1 'PrincipalFirm/Network
                                m_clsMortProdExclusivityTable.SetPrincipalfirmID lvAvailableItems.ListItems.Item(X)
                            Case 3 'Club/Association
                                m_clsMortProdExclusivityTable.SetClubNetworkAssociationId lvAvailableItems.ListItems.Item(X)
                        End Select
                End Select
                
                clsTableAccess.Update
                SaveChangeRequest
            End If
        Next X
        Select Case intLaunchMode
            Case 1, 2
                PopulateAvailableProductsListView
                PopulateLinkedProductsListView
            Case 3
                PopulateAvailableIntroducersListView
                PopulateLinkedIntroducersListView
        End Select
    
    End If
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError

End Sub
Private Sub cmdCancel_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Called when the user presses the Cancel button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdOK_Click()
Dim bRet As Boolean
    bRet = DoOKProcessing()
    If bRet = True Then
        SetReturnCode
        Hide
    
    End If
End Sub
Private Sub cmdRemove_Click()
Dim X As Integer
Dim colMatchValues As Collection
Dim colFields As Collection
Dim clsTableAccess As TableAccess
Dim strSQL As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

    On Error GoTo Failed

    Set conn = g_clsDataAccess.GetActiveConnection

    Set clsTableAccess = m_clsMortProdExclusivityTable

    If CheckIfOKToSave() Then
        For X = 1 To lvLinkedItems.ListItems.Count
            If lvLinkedItems.ListItems.Item(X).Selected Then
                Set colFields = New Collection
                Set colMatchValues = New Collection
                
                strSQL = "Select PRODUCTEXCLUSIVITYSEQNO from MORTGAGEPRODUCTEXCLUSIVITY Where MORTGAGEPRODUCTCODE = '"
                Select Case intLaunchMode
                    Case 1 'Associations/Clubs
                        strSQL = strSQL & lvLinkedItems.ListItems.Item(X) & "' AND CLUBNETWORKASSOCIATIONID = '" & strClubNetworkAssociationID & "'"
                    Case 2 'Principal Firms/Packagers
                        strSQL = strSQL & lvLinkedItems.ListItems.Item(X) & "' AND PRINCIPALFIRMID = '" & strPrincipalFirmID & "'"
                    Case 3 'Product
                        strSQL = strSQL & strProductCode & "' AND "
                        Select Case lvLinkedItems.ListItems.Item(X).SubItems(4)
                            Case 1 'PrincipalFirm/Network
                                strSQL = strSQL & "PRINCIPALFIRMID = '" & lvLinkedItems.ListItems.Item(X) & "'"
                            Case 3 'Club/Association
                                strSQL = strSQL & "CLUBNETWORKASSOCIATIONID = '" & lvLinkedItems.ListItems.Item(X) & "'"
                        End Select
                End Select
        
                Set rs = conn.Execute(strSQL)
                colFields.Add "PRODUCTEXCLUSIVITYSEQNO"
                colMatchValues.Add rs(0)
                clsTableAccess.SetKeyMatchFields colFields
                clsTableAccess.SetKeyMatchValues colMatchValues
    
                clsTableAccess.DeleteRow colMatchValues
                clsTableAccess.Update
            
                g_clsHandleUpdates.SaveChangeRequest m_clsMortProdExclusivityTable, , PromoteDelete
            End If
        Next X

        Select Case intLaunchMode
            Case 1, 2
                PopulateAvailableProductsListView
                PopulateLinkedProductsListView
            Case 3
                PopulateLinkedIntroducersListView
        End Select
    End If
    Set conn = Nothing
    Set rs = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdSearch_Click()
' TW 14/02/2007 EP2_1366
Dim T As Control
Dim blnSearchOK As Boolean
    For Each T In Me.Controls
        If TypeOf T Is MSGEditBox Then
            If Len(T.Text) > 0 Then
                Select Case T.Name
                    Case "txtUnitID", "txtFSARefNo", "txtCompanyName", "txtTown", "txtPostCode"
                        blnSearchOK = True
                        Exit For
                End Select
            End If
        End If
    Next
    If Not blnSearchOK Then
        MsgBox "No Search Criteria have been entered", vbInformation
        Exit Sub
    End If
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
' TW 14/02/2007 EP2_1366 End
    PopulateAvailableIntroducersListView
' TW 14/02/2007 EP2_1366
    If lvAvailableItems.ListItems.Count = 0 Then
        If Len(txtTown.Text) = 0 Then
            MsgBox "No records found matching your Search Criteria." & vbCrLf & "It could be that the Company concerned has not been allocated a Unit number", vbExclamation
        Else
            MsgBox "No records found with this location", vbExclamation
        End If
    End If
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
' TW 14/02/2007 EP2_1366 End
End Sub

Private Sub lvAvailableItems_Click()
    Frame4.Visible = False
' TW 14/02/2007 EP2_1366
    If lvAvailableItems.ListItems.Count = 0 Then
        cmdAdd.Enabled = False
        Exit Sub
    End If
' TW 14/02/2007 EP2_1366 End
    cmdAdd.Enabled = True
    cmdRemove.Enabled = False
End Sub
Private Sub lvLinkedItems_Click()
Dim intCountSelected As Integer
Dim X As Integer
' TW 14/02/2007 EP2_1366
    If lvLinkedItems.ListItems.Count = 0 Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
' TW 14/02/2007 EP2_1366 End
    cmdAdd.Enabled = False
    cmdApply.Enabled = False
    intType = 0
    For X = 1 To lvLinkedItems.ListItems.Count
        If lvLinkedItems.ListItems.Item(X).Selected Then
            intCountSelected = intCountSelected + 1
            intType = lvLinkedItems.ListItems.Item(X).SubItems(4)
            Exit For
        End If
    Next X
    If intCountSelected = 1 Then
        txtProcFeeLoading.Text = lvLinkedItems.ListItems.Item(X).SubItems(5)
    End If
    Frame4.Visible = (intCountSelected = 1)
End Sub
Private Sub txtCompanyName_LostFocus()
    EnableSearch txtCompanyName
End Sub
Private Sub txtFSARefNo_LostFocus()
    EnableSearch txtFSARefNo
End Sub
Private Sub txtPostCode_LostFocus()
    EnableSearch txtPostCode
End Sub

Private Sub txtProcFeeLoading_KeyPress(KeyAscii As Integer)
    cmdApply.Enabled = (Len(txtProcFeeLoading.Text) > 0)
End Sub

Private Sub txtTown_LostFocus()
    EnableSearch txtTown
End Sub
Private Sub txtUnitID_LostFocus()
    EnableSearch txtUnitID
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub AppendToWhereClause(ByRef strWhere As String, strFieldName As String, strFieldValue As String)
    If Len(strFieldValue) > 0 Then
        If Right$(strWhere, 1) <> "(" And Len(strWhere) > 0 Then
            strWhere = strWhere & " OR "
        End If
' TW 09/02/2007 EP2_1298
'        strWhere = strWhere & strFieldName & " Like '" & strFieldValue & "'"
        strWhere = strWhere & LikeOrEquals(strFieldName, strFieldValue)
' TW 09/02/2007 EP2_1298 End
    End If
End Sub
Private Function CheckIfOKToSave() As Boolean
Dim bRet As Boolean
    On Error GoTo Failed
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    CheckIfOKToSave = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    CheckIfOKToSave = False
End Function
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
'        SaveScreenData
'        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function
Private Sub EnableSearch(C As Control)
Dim blnMutuallyExclusive As Boolean
Dim T As Control
    If TypeOf C Is MSGEditBox Then
        If Len(C.Text) > 0 Then
            Select Case C.Name
                Case "txtFSARefNo"
                    txtUnitID.Text = ""
                    blnMutuallyExclusive = True
                Case "txtUnitID"
                    txtFSARefNo.Text = ""
                    blnMutuallyExclusive = True
                Case Else
                    txtUnitID.Text = ""
                    txtFSARefNo.Text = ""
            End Select
            If blnMutuallyExclusive Then
                txtCompanyName.Text = ""
                txtTown.Text = ""
                txtPostCode.Text = ""
            End If
        End If
    End If
' TW 14/02/2007 EP2_1366
'    For Each T In Me.Controls
'        If TypeOf T Is MSGEditBox Then
'            If Len(T.Text) > 0 Then
'                Select Case T.Name
'                    Case "txtUnitID", "txtFSARefNo", "txtCompanyName", "txtTown", "txtPostCode"
'                        cmdSearch.Enabled = True
'                        Exit For
'                End Select
'            End If
'        End If
'    Next
' TW 14/02/2007 EP2_1366 End
End Sub
Private Sub GetProductExcusivitySeqNo()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo Failed

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETNEXTSEQUENCENUMBER"
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "ProductExclusivitySeqNo")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strProductExclusivitySeqNo = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strProductExclusivitySeqNo) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub
Private Sub PopulateAvailableIntroducersListView()
Dim strCompanyName As String
Dim strFSARef As String
Dim strPostCode As String
Dim strTown As String
Dim strUnitID As String

Dim strPrincipalNetworkSearch As String
Dim strClubAssociationSearch As String

    On Error GoTo Failed

    strCompanyName = Replace(txtCompanyName.Text, "*", "%")
    strFSARef = Replace(txtFSARefNo.Text, "*", "%")
    strPostCode = Replace(txtPostCode.Text, "*", "%")
    strTown = Replace(txtTown.Text, "*", "%")
    strUnitID = Replace(txtUnitID.Text, "*", "%")
    
    strSearch = ""
    
' TW 30/07/2007 VR654
'    strPrincipalNetworkSearch = "SELECT PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
'                "FSAREF, REPLACE(REPLACE(LTRIM(" & _
'                "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
'                "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 1 As TYPE, null As PROCFEELOADING, UNITID " & _
'                "From PRINCIPALFIRM WHERE "
        
    strPrincipalNetworkSearch = "SELECT PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 1 As TYPE, null As PROCFEELOADING, UNITID, ISNULL(PACKAGERINDICATOR, 0) AS PACKAGERINDICATOR  " & _
                "From PRINCIPALFIRM WHERE "
        
'    strClubAssociationSearch = "SELECT CLUBNETWORKASSOCIATIONID AS ID, MORTGAGECLUBNETWORKASSOCNAME AS COMPANYNAME, '' AS FSAREF, " & _
'                "REPLACE(REPLACE(LTRIM(" & _
'                        "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
'                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 3 AS TYPE, null As PROCFEELOADING, '' AS UNITID " & _
'                "From MORTGAGECLUBNETWORKASSOCIATION WHERE "
    
    strClubAssociationSearch = "SELECT CLUBNETWORKASSOCIATIONID AS ID, MORTGAGECLUBNETWORKASSOCNAME AS COMPANYNAME, '' AS FSAREF, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 3 AS TYPE, null As PROCFEELOADING, '' AS UNITID, ISNULL(PACKAGERINDICATOR, 0) AS PACKAGERINDICATOR " & _
                "From MORTGAGECLUBNETWORKASSOCIATION WHERE "
' TW 30/07/2007 VR654 End
            
    If Len(strFSARef) > 0 Then
' TW 09/02/2007 EP2_1298
'        strPrincipalNetworkSearch = strPrincipalNetworkSearch & "FSAREF Like '" & strFSARef & "' "
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & LikeOrEquals("FSAREF", strFSARef)
' TW 09/02/2007 EP2_1298 End
    End If
    If Len(strUnitID) > 0 Then
' TW 09/02/2007 EP2_1298
'        strPrincipalNetworkSearch = strPrincipalNetworkSearch & "UNITID Like '" & strUnitID & "'"
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & LikeOrEquals("UNITID", strUnitID)
' TW 09/02/2007 EP2_1298 End
    End If
    If Len(strCompanyName) > 0 Then
' TW 09/02/2007 EP2_1298
'        strPrincipalNetworkSearch = strPrincipalNetworkSearch & "PRINCIPALFIRMNAME Like '" & strCompanyName & "'"
'        strClubAssociationSearch = strClubAssociationSearch & "MORTGAGECLUBNETWORKASSOCNAME Like '" & strCompanyName & "'"
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & LikeOrEquals("PRINCIPALFIRMNAME", strCompanyName)
        strClubAssociationSearch = strClubAssociationSearch & LikeOrEquals("MORTGAGECLUBNETWORKASSOCNAME", strCompanyName)
' TW 09/02/2007 EP2_1298 End
        If Len(strTown) > 0 Then
            strPrincipalNetworkSearch = strPrincipalNetworkSearch & " AND "
            strClubAssociationSearch = strClubAssociationSearch & " AND "
        End If
    End If
    If Len(strTown) > 0 Then
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & " ("
        strClubAssociationSearch = strClubAssociationSearch & " ("
    End If
    AppendToWhereClause strPrincipalNetworkSearch, "ADDRESSLINE1", strTown
    AppendToWhereClause strPrincipalNetworkSearch, "ADDRESSLINE2", strTown
    AppendToWhereClause strPrincipalNetworkSearch, "ADDRESSLINE3", strTown
    AppendToWhereClause strPrincipalNetworkSearch, "ADDRESSLINE4", strTown
    AppendToWhereClause strPrincipalNetworkSearch, "ADDRESSLINE5", strTown
    
    AppendToWhereClause strClubAssociationSearch, "TOWN", strTown
    
    If Len(strTown) > 0 Then
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & ")"
        strClubAssociationSearch = strClubAssociationSearch & ")"
    End If
    
    If (Len(strTown) > 0 Or Len(strCompanyName) > 0) And Len(strPostCode) > 0 Then
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & " AND "
        strClubAssociationSearch = strClubAssociationSearch & " AND "
    End If
    If Len(strPostCode) > 0 Then
' TW 09/02/2007 EP2_1298
'        strPrincipalNetworkSearch = strPrincipalNetworkSearch & "POSTCODE Like '" & Replace(strPostCode, " ", "") & "'"
'        strClubAssociationSearch = strClubAssociationSearch & "POSTCODE Like '" & Replace(strPostCode, " ", "") & "'"
        strPrincipalNetworkSearch = strPrincipalNetworkSearch & LikeOrEquals("POSTCODE", strPostCode)
        strClubAssociationSearch = strClubAssociationSearch & LikeOrEquals("POSTCODE", strPostCode)
' TW 09/02/2007 EP2_1298 End
    End If
    strPrincipalNetworkSearch = strPrincipalNetworkSearch & " AND PRINCIPALFIRMID Not In "
' TW 09/02/2007 EP2_1298
'    strClubAssociationSearch = strClubAssociationSearch & " AND CLUBNETWORKASSOCIATIONID Not In "
    If Len(strTown) > 0 Or Len(strCompanyName) > 0 Or Len(strPostCode) > 0 Then
        strClubAssociationSearch = strClubAssociationSearch & " AND "
    End If
    strClubAssociationSearch = strClubAssociationSearch & " CLUBNETWORKASSOCIATIONID Not In "
' TW 09/02/2007 EP2_1298 End
    
    strPrincipalNetworkSearch = strPrincipalNetworkSearch & "(SELECT E.PRINCIPALFIRMID From MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN PRINCIPALFIRM P ON E.PRINCIPALFIRMID= P.PRINCIPALFIRMID WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "')"
    strClubAssociationSearch = strClubAssociationSearch & "(SELECT E.CLUBNETWORKASSOCIATIONID From MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN MORTGAGECLUBNETWORKASSOCIATION M ON E.CLUBNETWORKASSOCIATIONID = M.CLUBNETWORKASSOCIATIONID WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "')"
                
' TW 09/02/2007 EP2_1298
'    strSearch = strPrincipalNetworkSearch & " Union " & strClubAssociationSearch & " ORDER BY COMPANYNAME"
    If Len(strFSARef) = 0 And Len(strUnitID) = 0 Then
        strSearch = strPrincipalNetworkSearch & " Union " & strClubAssociationSearch & " ORDER BY COMPANYNAME"
    Else
        strSearch = strPrincipalNetworkSearch
    End If
' TW 09/02/2007 EP2_1298 End

    PopulateListView lvAvailableItems
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateAvailableProductsListView()
    On Error GoTo Failed
    
' TW 30/07/2007 VR654
'    strSearch = "SELECT P.MORTGAGEPRODUCTCODE, P.STARTDATE, ENDDATE, PRODUCTNAME, 0 AS TYPE, null AS PROCFEELOADING, ORGANISATIONID " & _
'                "FROM MORTGAGEPRODUCT P INNER JOIN MORTGAGEPRODUCTLANGUAGE L ON P.MORTGAGEPRODUCTCODE = L.MORTGAGEPRODUCTCODE " & _
'                "AND P.STARTDATE = L.STARTDATE " & _
'                "WHERE EXCLUSIVEIND = 1 " & _
'                "AND P.MORTGAGEPRODUCTCODE NOT IN (SELECT MORTGAGEPRODUCTCODE FROM MORTGAGEPRODUCTEXCLUSIVITY WHERE "
    
    strSearch = "SELECT P.MORTGAGEPRODUCTCODE, P.STARTDATE, ENDDATE, PRODUCTNAME, 0 AS TYPE, null AS PROCFEELOADING, ORGANISATIONID, 0 AS PACKAGERINDICATOR " & _
                "FROM MORTGAGEPRODUCT P INNER JOIN MORTGAGEPRODUCTLANGUAGE L ON P.MORTGAGEPRODUCTCODE = L.MORTGAGEPRODUCTCODE " & _
                "AND P.STARTDATE = L.STARTDATE " & _
                "WHERE EXCLUSIVEIND = 1 " & _
                "AND P.MORTGAGEPRODUCTCODE NOT IN (SELECT MORTGAGEPRODUCTCODE FROM MORTGAGEPRODUCTEXCLUSIVITY WHERE "
' TW 30/07/2007 VR654 End
    
    Select Case intLaunchMode
        Case 1 'Associations, Clubs
            strSearch = strSearch & "CLUBNETWORKASSOCIATIONID = '" & strClubNetworkAssociationID & "')"
        Case 2 'Principal Firms, Packagers
            strSearch = strSearch & "PRINCIPALFIRMID = '" & strPrincipalFirmID & "')"
    End Select
    PopulateListView lvAvailableItems
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateLinkedIntroducersListView()

    On Error GoTo Failed
                
' TW 30/07/2007 VR654
'    strSearch = "SELECT E.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, FSAREF, " & _
'                "REPLACE(REPLACE(LTRIM(" & _
'                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
'                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 1 AS TYPE, PROCFEELOADING, UNITID " & _
'                "From " & _
'                "MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN PRINCIPALFIRM P ON E.PRINCIPALFIRMID = P.PRINCIPALFIRMID " & _
'                "WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "' " & _
'                "Union " & _
'                "SELECT E.CLUBNETWORKASSOCIATIONID AS ID, MORTGAGECLUBNETWORKASSOCNAME AS COMPANYNAME, '' AS FSAREF, " & _
'                "REPLACE(REPLACE(LTRIM(" & _
'                        "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
'                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 3 AS TYPE, PROCFEELOADING, '' AS UNITID " & _
'                "From " & _
'                "MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN MORTGAGECLUBNETWORKASSOCIATION A ON E.CLUBNETWORKASSOCIATIONID = A.CLUBNETWORKASSOCIATIONID " & _
'                "WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "' " & _
'                "ORDER BY COMPANYNAME"
                
    strSearch = "SELECT E.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, FSAREF, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 1 AS TYPE, PROCFEELOADING, UNITID, ISNULL(P.PACKAGERINDICATOR, 0) AS PACKAGERINDICATOR " & _
                "From " & _
                "MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN PRINCIPALFIRM P ON E.PRINCIPALFIRMID = P.PRINCIPALFIRMID " & _
                "WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "' " & _
                "Union " & _
                "SELECT E.CLUBNETWORKASSOCIATIONID AS ID, MORTGAGECLUBNETWORKASSOCNAME AS COMPANYNAME, '' AS FSAREF, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, 3 AS TYPE, PROCFEELOADING, '' AS UNITID, ISNULL(A.PACKAGERINDICATOR, 0) AS PACKAGERINDICATOR " & _
                "From " & _
                "MORTGAGEPRODUCTEXCLUSIVITY E INNER JOIN MORTGAGECLUBNETWORKASSOCIATION A ON E.CLUBNETWORKASSOCIATIONID = A.CLUBNETWORKASSOCIATIONID " & _
                "WHERE MORTGAGEPRODUCTCODE = '" & strProductCode & "' " & _
                "ORDER BY COMPANYNAME"
                
' TW 30/07/2007 VR654 End
                
    PopulateListView lvLinkedItems
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateLinkedProductsListView()
    On Error GoTo Failed
    
' TW 30/07/2007 VR654
'    strSearch = "SELECT P.MORTGAGEPRODUCTCODE, P.STARTDATE, ENDDATE, PRODUCTNAME, 0 As TYPE, PROCFEELOADING, ORGANISATIONID " & _
'                "FROM MORTGAGEPRODUCTEXCLUSIVITY E " & _
'                "INNER JOIN MORTGAGEPRODUCT P ON E.MORTGAGEPRODUCTCODE = P.MORTGAGEPRODUCTCODE " & _
'                "INNER JOIN MORTGAGEPRODUCTLANGUAGE L ON P.MORTGAGEPRODUCTCODE = L.MORTGAGEPRODUCTCODE " & _
'                "AND P.STARTDATE = L.STARTDATE WHERE "
    
    strSearch = "SELECT P.MORTGAGEPRODUCTCODE, P.STARTDATE, ENDDATE, PRODUCTNAME, 0 As TYPE, PROCFEELOADING, ORGANISATIONID, 0 AS PACKAGERINDICATOR " & _
                "FROM MORTGAGEPRODUCTEXCLUSIVITY E " & _
                "INNER JOIN MORTGAGEPRODUCT P ON E.MORTGAGEPRODUCTCODE = P.MORTGAGEPRODUCTCODE " & _
                "INNER JOIN MORTGAGEPRODUCTLANGUAGE L ON P.MORTGAGEPRODUCTCODE = L.MORTGAGEPRODUCTCODE " & _
                "AND P.STARTDATE = L.STARTDATE WHERE "
    
' TW 30/07/2007 VR654 End
    
    Select Case intLaunchMode
        Case 1 'Associations, Clubs
            strSearch = strSearch & "CLUBNETWORKASSOCIATIONID = '" & strClubNetworkAssociationID & "'"
        Case 2 'Principal Firms, Packagers
            strSearch = strSearch & "PRINCIPALFIRMID = '" & strPrincipalFirmID & "'"
    End Select
    PopulateListView lvLinkedItems
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateListView(lvListView As MSGListView)
Dim rs As ADODB.Recordset
Dim colLine As Collection


    On Error GoTo Failed
    
    Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSearch)
    lvListView.ListItems.Clear
            
    Do While Not rs.EOF
        Set colLine = New Collection
        
        colLine.Add Format$(rs(0))
        colLine.Add Format$(rs(1))
        colLine.Add Format$(rs(2))
        colLine.Add Format$(rs(3))
        colLine.Add Format$(rs(4))
        colLine.Add Format$(rs(5), "0.00")
' TW 30/07/2007 VR654
        colLine.Add Format$(rs(7))
' TW 30/07/2007 VR654 End
        lvListView.AddLine colLine
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Private Sub PopulateScreenFields()
Dim colMatchValues As Collection

' TW 30/07/2007 VR654
    intExclusivityPackagerIndicator = 0
' TW 30/07/2007 VR654 End

    Set colMatchValues = New Collection
    colMatchValues.Add m_colKeys(1)

    Select Case m_colKeys(2)
        Case "Introducer"
            Me.Caption = "Introducer - Find Exclusive Products"
            
            Frame3.Caption = "Available Products"
            Frame1.Caption = "Linked Products"
            
            lblDescription(0).Caption = "Introducer Name"
            lblDescription(1).Caption = "Type"
            lblLegend(1).Caption = Left$(m_colKeys(3), Len(m_colKeys(3)) - 1) & " "
            Select Case m_colKeys(3)
                Case "Associations", "Clubs"
                    intLaunchMode = 1
                    strClubNetworkAssociationID = m_colKeys(1)
                    Set m_clsMortgageClubNetAssocTable = New MortgageClubNetAssocTable
            
                    TableAccess(m_clsMortgageClubNetAssocTable).SetKeyMatchValues colMatchValues
                    TableAccess(m_clsMortgageClubNetAssocTable).GetTableData

' TW 30/07/2007 VR654
                    intExclusivityPackagerIndicator = m_clsMortgageClubNetAssocTable.GetPackagerIndicator
' TW 30/07/2007 VR654 End
            
                    lblLegend(0).Caption = m_clsMortgageClubNetAssocTable.GetMortgageClubNetworkAssocName & " "
                
' TW 26/04/2007 EP2_2599
'                Case "Principal Firms", "Packagers"
                Case "Principal Firms/Networks", "Packagers"
' TW 26/04/2007 EP2_2599 End
                    intLaunchMode = 2
                    strPrincipalFirmID = m_colKeys(1)
                    Set m_clsPrincipalFirmTable = New PrincipalFirmTable
                    TableAccess(m_clsPrincipalFirmTable).SetKeyMatchValues colMatchValues
                    TableAccess(m_clsPrincipalFirmTable).GetTableData

' TW 30/07/2007 VR654
                    intExclusivityPackagerIndicator = m_clsPrincipalFirmTable.GetPackagerIndicator
' TW 30/07/2007 VR654 End

                    lblLegend(0).Caption = m_clsPrincipalFirmTable.GetPrincipalFirmName & " "

            End Select
        Case "Product"
            intLaunchMode = 3
            strProductCode = m_colKeys(1)
        
            Set m_clsMortProdExclusivityTable = New MortProdExclusivityTable
            Set m_clsMortProdLanguageTable = New MortProdLanguageTable
            Me.Caption = "Exclusive Products - Find Introducers"
            
            Frame3.Caption = "Available Introducers"
            Frame1.Caption = "Linked Introducers"
            
            colMatchValues.Add m_colKeys(3) 'Start Date
            
            If g_clsVersion.DoesVersioningExist() Then
                colMatchValues.Add m_colKeys(4) 'Language
            End If
    
            TableAccess(m_clsMortProdLanguageTable).SetKeyMatchValues colMatchValues
            TableAccess(m_clsMortProdLanguageTable).GetTableData
            
            lblDescription(0).Caption = "Mortgage Product Name"
            lblLegend(0).Caption = m_clsMortProdLanguageTable.GetProductName & " "
            lblDescription(1).Caption = "Details"
            lblLegend(1).Caption = m_clsMortProdLanguageTable.GetProductDetails & " "
    
    End Select
    lblLegend(0).Left = lblDescription(0).Left + lblDescription(0).Width + 120
    lblDescription(1).Left = lblLegend(0).Left + lblLegend(0).Width + 120
    lblLegend(1).Left = lblDescription(1).Left + lblDescription(1).Width + 120
    
    SetColumnHeaders
    
    Select Case intLaunchMode
        Case 1, 2
            Frame2.Visible = False
            Frame3.Top = Frame2.Top
            Frame1.Top = Frame3.Top + Frame3.Height + 120
            cmdOK.Top = Frame1.Top + Frame1.Height + 120
            cmdCancel.Top = cmdOK.Top
            Me.Height = cmdOK.Top + cmdOK.Height + 480
            PopulateAvailableProductsListView
            PopulateLinkedProductsListView
        Case 3
            Frame2.Visible = True
            PopulateLinkedIntroducersListView
    End Select

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveChangeRequest()
Dim colMatchValues As Collection
Dim clsTableAccess As TableAccess
    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strProductExclusivitySeqNo
    
    Set clsTableAccess = m_clsMortProdExclusivityTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsMortProdExclusivityTable
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetAddState()
    
    On Error GoTo Failed
    
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetColumnHeaders()

    Dim colHeaders As New Collection
    Dim clsHeader As listViewAccess
        
    Select Case intLaunchMode
        Case 1, 2      'Introducer
            
            clsHeader.nWidth = 15
            clsHeader.sName = "Product Code"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 15
            clsHeader.sName = "Start Date"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 15
            clsHeader.sName = "End Date"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 35
            clsHeader.sName = "Product Name"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 0
            clsHeader.sName = ""
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 20
            clsHeader.sName = "Proc Fee Loading"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 0
            clsHeader.sName = ""
            colHeaders.Add clsHeader
            
        Case 3      'Product
            clsHeader.nWidth = 10
            clsHeader.sName = "Id"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 30
            clsHeader.sName = "Company Name"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 10
            clsHeader.sName = "FSA Ref"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 30
            clsHeader.sName = "Address"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 0
            clsHeader.sName = ""    'Type (1 = Principal Firm, 2 = AR Firm, 3 = Club/Association)
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 20
            clsHeader.sName = "Proc Fee Loading"
            colHeaders.Add clsHeader
            
            clsHeader.nWidth = 0
            clsHeader.sName = ""
            colHeaders.Add clsHeader
    End Select
        
        
        
    
    lvAvailableItems.AddHeadings colHeaders
    lvLinkedItems.AddHeadings colHeaders

End Sub
Private Sub SetEditState()
    
    On Error GoTo Failed
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetIsEdit(blnEditStatus As Boolean)
    m_bIsEdit = blnEditStatus
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
