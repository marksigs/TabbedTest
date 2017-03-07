VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMaintainFSAAssociations 
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmMaintainFSAAssociations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   9705
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Linked Items"
      Height          =   2175
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   9495
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin MSGOCX.MSGListView lvLinkedItems 
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2143
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
      Begin MSGOCX.MSGEditBox txtTradingAs 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Trading As"
         Height          =   195
         Index           =   5
         Left            =   3000
         TabIndex        =   24
         Top             =   1680
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Items"
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   9495
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin MSGOCX.MSGListView lvAvailableItems 
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.TextBox txtSearchType 
      Height          =   285
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8340
      TabIndex        =   14
      Top             =   7200
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   7200
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Find"
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   9495
      Begin MSGOCX.MSGEditBox txtFSARefNo 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
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
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   8040
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin MSGOCX.MSGEditBox txtUnitID 
         Height          =   285
         Left            =   3480
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         Caption         =   "Unit ID"
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Post Code"
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   19
         Top             =   1080
         Width           =   735
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
         Caption         =   "Company Name "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FSA Ref No"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label lblIntroducerName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Introducer Name"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmMaintainFSAAssociations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmMaintainFSAAssociations
' Description   :
'
' Change history
' Prog      Date        Description
' TW        17/10/2006  EP2_15 - Created
' TW        18/11/2006  EP2_132 ECR20/21 Changes for wildcard consistency and promotion  of deletes
' TW        06/12/2006  EP2_304 - overflow error associating a principal firm to a da
' TW        06/12/2006  EP2_305 - entering an fsa number doesn't find a match
' TW        10/01/2007  EP2_637 - Case not ingesting from the web onto omiga
' TW        29/01/2007  EP2_863 - Add new col to IntroducerFirm - "Trading As"
' TW        09/02/2007  EP2_1296 - Error On the AR Firms Tab when selected the firm and clicked the Modify Button
' TW        14/02/2007  EP2_1366 - Error amending AR Broker records
' TW        21/04/2007  EP2_2266 - Only one Packager firm should be linked to Individual packager
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Const mc_strTHIS_MODULE As String = "frmMaintainFSAAssociations"

' Private data
Private m_clsOrgUserTable As OrgUserTable
Private m_clsUserRoleTable As UserRoleTable

Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Private strUserID As String
'Private strUnitId As String
Private strSearch As String
Private strWhere As String
Private strIntroducerID As String
Private strPrincipalFirmID As String
Private strARFirmID As String
Private strIntroducerFirmSeqNo As String

Private Function LikeOrEquals(strFieldName As String, strValue As String) As String
' TW 06/12/2006 EP2_305
Dim strComparison As String
    strComparison = strFieldName & " "
    If InStr(1, strValue, "%") > 0 Then
        strComparison = strComparison & " Like '"
    Else
        strComparison = strComparison & " = '"
    End If
    LikeOrEquals = strComparison & strValue & "'"
' TW 06/12/2006 EP2_305 End
End Function



Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub PopulateAvailableFirmsListView()
' TW 18/11/2006 EP2_132
Dim strCompanyName As String
Dim strFSARef As String
Dim strPostCode As String
Dim strTown As String
Dim strUnitID As String

    strCompanyName = Replace(txtCompanyName.Text, "*", "%")
    strFSARef = Replace(txtFSARefNo.Text, "*", "%")
    strPostCode = Replace(txtPostCode.Text, "*", "%")
    strTown = Replace(txtTown.Text, "*", "%")
    strUnitID = Replace(txtUnitID.Text, "*", "%")
' TW 18/11/2006 EP2_132 End
    
    strSearch = ""
    strWhere = ""
    Select Case txtSearchType.Text
        Case "Individual Packager", "DA Individual Broker"
            strSearch = "SELECT UNITID, PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                        "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "0 AS INTRODUCERFIRMSEQNO, " & _
                        "'' AS TRADINGAS " & _
                        "From PRINCIPALFIRM "
        Case "AR Individual Broker"
            strSearch = "SELECT UNITID, ARFIRMID AS ID, ARFIRMNAME AS COMPANYNAME, " & _
                        "FSAARFIRMREF AS FSAREF, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "0 AS INTRODUCERFIRMSEQNO, " & _
                        "'' AS TRADINGAS " & _
                        "From ARFIRM "
    End Select
    Select Case txtSearchType.Text
        Case "Individual Packager"
            strWhere = strWhere & "PACKAGERINDICATOR = 1 AND ("
            If Len(strFSARef) > 0 Then
                strWhere = strWhere & LikeOrEquals("FSAREF", strFSARef)
            End If
        Case "DA Individual Broker"
            strWhere = strWhere & "(PACKAGERINDICATOR = 0 OR PACKAGERINDICATOR IS NULL) AND ("
            If Len(strFSARef) > 0 Then
                strWhere = strWhere & LikeOrEquals("FSAREF", strFSARef)
            End If
        Case "AR Individual Broker"
            If Len(strFSARef) > 0 Then
                strWhere = strWhere & LikeOrEquals("FSAARFIRMREF", strFSARef)
            End If
    End Select
    If Len(strUnitID) > 0 Then
        strWhere = strWhere & LikeOrEquals("UNITID", strUnitID)
    End If
    If Len(strCompanyName) > 0 Then
        Select Case txtSearchType.Text
            Case "Individual Packager", "DA Individual Broker"
                strWhere = strWhere & LikeOrEquals("PRINCIPALFIRMNAME", strCompanyName)
            Case "AR Individual Broker"
                strWhere = strWhere & LikeOrEquals("ARFIRMNAME", strCompanyName)
        End Select
        If Len(strTown) > 0 Then
            strWhere = strWhere & " AND "
        End If
    End If
    If Len(strTown) > 0 Then
        strWhere = strWhere & " ("
    End If
    AppendToWhereClause "ADDRESSLINE1", strTown
    AppendToWhereClause "ADDRESSLINE2", strTown
    AppendToWhereClause "ADDRESSLINE3", strTown
    AppendToWhereClause "ADDRESSLINE4", strTown
    AppendToWhereClause "ADDRESSLINE5", strTown
    If Len(strTown) > 0 Then
        strWhere = strWhere & ")"
    End If
    If (Len(strTown) > 0 Or Len(strCompanyName) > 0) And Len(strPostCode) > 0 Then
        strWhere = strWhere & " AND "
    End If
    If Len(strPostCode) > 0 Then
        strWhere = strWhere & LikeOrEquals("POSTCODE", Replace(strPostCode, " ", ""))
    End If
    strSearch = strSearch & " WHERE ISNULL(UNITID, '0') > '0' AND " & strWhere
    Select Case txtSearchType
        Case "Individual Packager", "DA Individual Broker"
            strSearch = strSearch & ") AND PRINCIPALFIRMID Not In "
        Case Else
            strSearch = strSearch & " AND ARFIRMID Not In "
    End Select
    Select Case txtSearchType.Text
        Case "Individual Packager"
            strSearch = strSearch & "(SELECT P.PRINCIPALFIRMID " & _
                        "From PRINCIPALFIRM P INNER JOIN INTRODUCERFIRM I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
                        "WHERE PACKAGERINDICATOR = 1 AND INTRODUCERID = '" & strIntroducerID & "')"
        Case "DA Individual Broker"
            strSearch = strSearch & "(SELECT P.PRINCIPALFIRMID " & _
                        "From PRINCIPALFIRM P INNER JOIN INTRODUCERFIRM I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
                        "WHERE (PACKAGERINDICATOR = 0 OR PACKAGERINDICATOR IS NULL) AND INTRODUCERID = '" & strIntroducerID & "')"
        Case "AR Individual Broker"
            strSearch = strSearch & "(SELECT A.ARFIRMID " & _
                        "From ARFIRM A INNER JOIN INTRODUCERFIRM I ON A.ARFIRMID = I.ARFIRMID " & _
                        "WHERE INTRODUCERID = '" & strIntroducerID & "')"
    End Select
    PopulateListView lvAvailableItems
End Sub

Public Sub SetIsEdit(blnEditStatus As Boolean)
    m_bIsEdit = blnEditStatus
End Sub

Private Sub AppendToWhereClause(strFieldName As String, strFieldValue As String)
    If Len(strFieldValue) > 0 Then
        If Right$(strWhere, 1) <> "(" And Len(strWhere) > 0 Then
            strWhere = strWhere & " OR "
        End If
' TW 06/12/2006 EP2_305
        'strWhere = strWhere & strFieldName & " Like '" & strFieldValue & "'"
        strWhere = strWhere & LikeOrEquals(strFieldName, strFieldValue)
' TW 06/12/2006 EP2_305 End
    End If
End Sub

Private Sub EnableSearch(C As Control)
Dim blnMutuallyExclusive As Boolean
Dim T As Control
' TW 10/01/2007 EP2_637
' TW 14/02/2007 EP2_1366
'    cmdSearch.Enabled = False
' TW 14/02/2007 EP2_1366 End
' TW 10/01/2007 EP2_637 End
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


Private Sub PopulateListView(lvListView As MSGListView)
Dim rs As ADODB.Recordset
Dim colLine As Collection

    On Error GoTo Failed
    
    Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSearch)
    
    lvListView.ListItems.Clear

            
    Do While Not rs.EOF
        Set colLine = New Collection
        
        colLine.Add Format$(rs(0))  'UNITID
        colLine.Add Format$(rs(1))  'ID
        colLine.Add Format$(rs(2))  'COMPANYNAME
        colLine.Add Format$(rs(3))  'FSAREF
        colLine.Add Format$(rs(4))  'ADDRESS
        colLine.Add Format$(rs(5))  'INTRODUCERFIRMSEQNO
' TW 29/01/2007 EP2_863
        colLine.Add Format$(rs(6))  'TRADINGAS
' TW 29/01/2007 EP2_863 End
        lvListView.AddLine colLine
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub


Private Sub cmdAdd_Click()
' TW 10/01/2007 EP2_637 - Rewritten

Dim X As Long
Dim clsTableAccess As TableAccess
Dim colLine As Collection

' Link Introducer via Introducer Firm to ARFirm or PrincipalFirm
    On Error GoTo Failed:
    For X = lvAvailableItems.ListItems.Count To 1 Step -1
        If lvAvailableItems.ListItems.Item(X).Selected Then
            Set colLine = New Collection
            colLine.Add lvAvailableItems.ListItems.Item(X)
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(1)
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(2)
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(3)
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(4)
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(5)
            
' TW 29/01/2007 EP2_863
            colLine.Add lvAvailableItems.ListItems.Item(X).SubItems(6)
' TW 29/01/2007 EP2_863 End
            
            lvLinkedItems.AddLine colLine
            lvAvailableItems.ListItems.Remove X
        End If
    Next X
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

Private Function DoOKProcessing() As Boolean
Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
    
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function

Private Sub cmdRemove_Click()
' TW 10/01/2007 EP2_637 - Rewritten
Dim X As Long
Dim strSQL As String
Dim strIntroducerFirmSeqNo As String
Dim rs As ADODB.Recordset


    strSQL = "SELECT INTRODUCERFIRMSEQNO FROM INTRODUCER I " & _
             "INNER JOIN INTRODUCERFIRM F ON I.INTRODUCERID = F.INTRODUCERID " & _
             "INNER JOIN PRINCIPALFIRM P ON F.PRINCIPALFIRMID = P.PRINCIPALFIRMID " & _
             "INNER JOIN USERROLE R ON P.UNITID = R.UNITID " & _
             "WHERE I.INTRODUCERID = '" & strIntroducerID & "' AND R.USERID = I.USERID"

    Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSQL)
    If Not rs.EOF Then
        strIntroducerFirmSeqNo = rs!INTRODUCERFIRMSEQNO
    End If
    
    On Error GoTo Failed
    For X = lvLinkedItems.ListItems.Count To 1 Step -1
        If lvLinkedItems.ListItems.Item(X).Selected Then
            If lvLinkedItems.ListItems.Item(X).SubItems(5) = strIntroducerFirmSeqNo Then
                MsgBox "This entry cannot be deleted as it relates to the Omiga User for this Introducer", vbExclamation
            Else
                lvLinkedItems.ListItems.Remove X
            End If
        End If
    Next X
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
' TW 14/02/2007 EP2_1366 End

' TW 06/12/2006 EP2_304
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
'    DoEvents
' TW 06/12/2006 EP2_304 End
    
    PopulateAvailableFirmsListView
' TW 14/02/2007 EP2_1366
    If lvAvailableItems.ListItems.Count = 0 Then
        If Len(txtTown.Text) = 0 Then
            MsgBox "No records found matching your Search Criteria." & vbCrLf & "It could be that the Company concerned has not been allocated a Unit number", vbExclamation
        Else
            MsgBox "No records found with this location", vbExclamation
        End If
    End If
' TW 14/02/2007 EP2_1366 End
' TW 06/12/2006 EP2_304
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
' TW 06/12/2006 EP2_304 End
End Sub

Private Sub Form_Activate()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Select Case txtSearchType.Text
        Case "Individual Packager"
            Me.Caption = "Add Individual Packager to Packaging Firm"
        Case "DA Individual Broker"
            Me.Caption = "Add DA Individual Broker to Principal Firm"
        Case "AR Individual Broker"
            Me.Caption = "Add AR Individual Broker to AR Firm"
    End Select
        
    Set m_clsOrgUserTable = New OrgUserTable
'    g_clsFormProcessing.PopulateCombo "UserRole", Me.cboUserRole
    SetColumnHeaders
    
' TW 10/01/2007 EP2_637
'    PopulateScreenFields
' TW 10/01/2007 EP2_637 End
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetColumnHeaders()

    Dim colHeaders As New Collection
    Dim clsHeader As listViewAccess
        
    clsHeader.nWidth = 0
    clsHeader.sName = ""
    colHeaders.Add clsHeader
        
    clsHeader.nWidth = 0
    clsHeader.sName = ""
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 20
    clsHeader.sName = "Company Name"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 10
    clsHeader.sName = "FSA Ref"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 35
    clsHeader.sName = "Address"
    colHeaders.Add clsHeader
        
    clsHeader.nWidth = 0
    clsHeader.sName = ""
    colHeaders.Add clsHeader

' TW 29/01/2007 EP2_863
    
    clsHeader.nWidth = 15
    clsHeader.sName = "Trading As"
    colHeaders.Add clsHeader
' TW 29/01/2007 EP2_863 End
    
    lvAvailableItems.AddHeadings colHeaders
    lvLinkedItems.AddHeadings colHeaders

End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Private Sub PopulateScreenFields()
Dim colMatchValues As Collection

    Set m_clsOrgUserTable = New OrgUserTable
    Set colMatchValues = New Collection
    colMatchValues.Add m_colKeys(1)
    
    TableAccess(m_clsOrgUserTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsOrgUserTable).GetTableData
    
' TW 14/02/2007 EP2_1366
    If TableAccess(m_clsOrgUserTable).RecordCount = 0 Then
        MsgBox "An ORGANISATIONUSER record has not been found for this " & txtSearchType.Text & "." & vbCrLf & "This suggests that no Unit Number has been allocated"
    Else
        lblIntroducerName.Caption = m_clsOrgUserTable.GetForename & " " & m_clsOrgUserTable.GetSurname
    End If
' TW 14/02/2007 EP2_1366 End

'    txtActiveFrom.Text = Format$(Now, "dd/mm/yyyy")
    Set m_clsOrgUserTable = Nothing
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub SetEditState()
    
    On Error GoTo Failed
' TW 09/02/2007 EP2_1296
'    strIntroducerID = m_colKeys(1)
    strIntroducerID = m_colKeys(2)
' TW 09/02/2007 EP2_1296 End
' TW 10/01/2007 EP2_637
    PopulateScreenFields
' TW 10/01/2007 EP2_637 End
    PopulateLinkedFirmsListView
           
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





Private Sub lvAvailableItems_Click()
' TW 21/04/2007 EP2_2266
    If txtSearchType.Text = "Individual Packager" Then
        If lvLinkedItems.ListItems.Count > 0 Then
            cmdAdd.Enabled = False
            Exit Sub
        End If
    End If
' TW 21/04/2007 EP2_2266 End

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
' TW 14/02/2007 EP2_1366
    If lvLinkedItems.ListItems.Count = 0 Then
        cmdRemove.Enabled = False
        Exit Sub
    End If
' TW 14/02/2007 EP2_1366 End
    
    cmdAdd.Enabled = False
    cmdRemove.Enabled = True
' TW 29/01/2007 EP2_863
    txtTradingAs.Text = lvLinkedItems.ListItems.Item(lvLinkedItems.SelectedItem.Index).SubItems(6)
' TW 29/01/2007 EP2_863 End
End Sub

Private Sub PopulateLinkedFirmsListView()

    On Error GoTo Failed
    
    Select Case txtSearchType.Text
        Case "Individual Packager"
            strSearch = "SELECT UNITID, P.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                        "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "INTRODUCERFIRMSEQNO, TRADINGAS " & _
                        "From PRINCIPALFIRM P INNER JOIN INTRODUCERFIRM I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
                        "WHERE PACKAGERINDICATOR = 1 AND INTRODUCERID = '" & strIntroducerID & "'"
        Case "DA Individual Broker"
            strSearch = "SELECT UNITID, P.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                        "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "INTRODUCERFIRMSEQNO, TRADINGAS " & _
                        "From PRINCIPALFIRM P INNER JOIN INTRODUCERFIRM I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
                        "WHERE (PACKAGERINDICATOR = 0 OR PACKAGERINDICATOR IS NULL) AND INTRODUCERID = '" & strIntroducerID & "'"
        Case "AR Individual Broker"
            strSearch = "SELECT UNITID, A.ARFIRMID AS ID, ARFIRMNAME AS COMPANYNAME, " & _
                        "FSAARFIRMREF AS FSAREF, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "INTRODUCERFIRMSEQNO, TRADINGAS " & _
                        "From ARFIRM A INNER JOIN INTRODUCERFIRM I ON A.ARFIRMID = I.ARFIRMID " & _
                        "WHERE INTRODUCERID = '" & strIntroducerID & "'"
    End Select
    PopulateListView lvLinkedItems
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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


Private Sub txtTown_LostFocus()
    EnableSearch txtTown
End Sub


Private Sub txtTradingAs_Change()
    If lvLinkedItems.SelectedItem.Index > 0 Then
        lvLinkedItems.ListItems.Item(lvLinkedItems.SelectedItem.Index).SubItems(6) = txtTradingAs.Text
    End If
End Sub

Private Sub txtUnitID_LostFocus()
    EnableSearch txtUnitID
End Sub


