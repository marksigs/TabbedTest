VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMaintainFirmLinks 
   Caption         =   "Add AR Firm to Network"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmMaintainFirmLinks .frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9705
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Linked Items"
      Height          =   2655
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   9495
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin MSGOCX.MSGListView lvLinkedItems 
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
   Begin VB.Frame Frame3 
      Caption         =   "Available Items"
      Height          =   2655
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   9495
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin MSGOCX.MSGListView lvAvailableItems 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   7320
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   7320
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Find"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9495
      Begin MSGOCX.MSGEditBox txtFSARefNo 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin MSGOCX.MSGEditBox txtUnitID 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Post Code"
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Town"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company Name "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FSA Ref No"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMaintainFirmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmMaintainFirmLinks
' Description   :
'
' Change history
' Prog      Date        Description
' TW        08/12/2006  EP2_360 - Created
' TW        05/02/2007  EP2_706 - Rewritten because of specification change
' TW        13/02/2007  EP2_1334 - AR firms not holding value in firm status
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Const mc_strTHIS_MODULE As String = "frmMaintainFirmLinks"

' Private data
Private m_clsAppointmentsTable As AppointmentsTable

Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Private strUserID As String
Private strUnitID As String
Private strSearch As String
Private strWhere As String
Private strPrincipalFirmID As String
Private strARFirmID As String
Private strAppointmentID As String
Private Function LikeOrEquals(strFieldName As String, strValue As String) As String
Dim strComparison As String
    strComparison = strFieldName & " "
    If InStr(1, strValue, "%") > 0 Then
        strComparison = strComparison & " Like '"
    Else
        strComparison = strComparison & " = '"
    End If
    LikeOrEquals = strComparison & strValue & "'"
End Function


Private Sub SaveChangeRequest()
Dim sProductNumber As String
Dim colMatchValues As Collection
Dim clsTableAccess As TableAccess
    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strAppointmentID
    
    Set clsTableAccess = m_clsAppointmentsTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsAppointmentsTable
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub PopulateAvailableFirmsListView()
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
    
    strWhere = ""
    strSearch = "SELECT UNITID, PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
                "From PRINCIPALFIRM "
            
    strWhere = strWhere & "ISNULL(PACKAGERINDICATOR, 0) = 0 AND ("
    If Len(strFSARef) > 0 Then
        strWhere = strWhere & LikeOrEquals("FSAREF", strFSARef)
    End If
    
    If Len(strUnitID) > 0 Then
        strWhere = strWhere & LikeOrEquals("UNITID", strUnitID)
    End If
    If Len(strCompanyName) > 0 Then
        strWhere = strWhere & LikeOrEquals("PRINCIPALFIRMNAME", strCompanyName)
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
    strSearch = strSearch & " WHERE " & strWhere
    strSearch = strSearch & ") AND PRINCIPALFIRMID Not In "
' TW 13/02/2007 EP2_1334
'    strSearch = strSearch & "(SELECT P.PRINCIPALFIRMID " & _
'                "From PRINCIPALFIRM P INNER JOIN APPOINTMENTS I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
'                "WHERE ISNULL(PACKAGERINDICATOR, 0) = 0 AND ARFIRMID = '" & strARFirmID & "')"
    
    strSearch = strSearch & "(SELECT PRINCIPALFIRMID " & _
                "From APPOINTMENTS " & _
                "WHERE ARFIRMID = '" & strARFirmID & "')"
' TW 13/02/2007 EP2_1334 End
    
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
        strWhere = strWhere & LikeOrEquals(strFieldName, strFieldValue)
    End If
End Sub

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


Private Sub PopulateListView(lvListView As MSGListView)
Dim rs As ADODB.Recordset
Dim colLine As Collection


    On Error GoTo Failed
    
    Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSearch)
    
    lvListView.ListItems.Clear
            
    Do While Not rs.EOF
        Set colLine = New Collection
        
        colLine.Add Format$(rs(1))
        colLine.Add Format$(rs(2))
        colLine.Add Format$(rs(3))
        colLine.Add Format$(rs(4))
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
Dim X As Long
Dim clsTableAccess As TableAccess

' Link AR Firm via Appointments to PrincipalFirm
    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        For X = 1 To lvAvailableItems.ListItems.Count
            If lvAvailableItems.ListItems.Item(X).Selected Then
    
                Call GetAppointmentID
                Set m_colKeys = New Collection
                  
                g_clsFormProcessing.CreateNewRecord m_clsAppointmentsTable
                Set clsTableAccess = m_clsAppointmentsTable
                m_clsAppointmentsTable.SetAppointmentID strAppointmentID
                m_clsAppointmentsTable.SetPrincipalfirmID lvAvailableItems.ListItems.Item(X)
                m_clsAppointmentsTable.SetARFirmID strARFirmID
                
                clsTableAccess.Update
                SaveChangeRequest
            End If
        Next X
    
        Call PopulateLinkedFirmsListView
        Call PopulateAvailableFirmsListView
    End If
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError

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


Private Sub GetAppointmentID()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "AppointmentID")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strAppointmentID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strAppointmentID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
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
'        SaveScreenData
'        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function

Private Sub cmdRemove_Click()
Dim X As Long
Dim colMatchValues As Collection
Dim colFields As Collection
Dim clsTableAccess As TableAccess
Dim strSQL As String
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset

    On Error GoTo Failed

    Set conn = g_clsDataAccess.GetActiveConnection

    Set clsTableAccess = m_clsAppointmentsTable

    If CheckIfOKToSave() Then
        For X = 1 To lvLinkedItems.ListItems.Count
            If lvLinkedItems.ListItems.Item(X).Selected Then
                       
                Set colFields = New Collection
                Set colMatchValues = New Collection
            
                strSQL = "Select AppointmentID from APPOINTMENTS Where ARFIRMID = '" & strARFirmID & "' AND "
                strSQL = strSQL & "PRINCIPALFIRMID = '"
                strSQL = strSQL & lvLinkedItems.ListItems.Item(X) & "'"
                
                Set rs = conn.Execute(strSQL)
                colFields.Add "AppointmentID"
                colMatchValues.Add rs(0)
                clsTableAccess.SetKeyMatchFields colFields
                clsTableAccess.SetKeyMatchValues colMatchValues
    
                clsTableAccess.DeleteRow colMatchValues
                clsTableAccess.Update
            
                g_clsHandleUpdates.SaveChangeRequest m_clsAppointmentsTable, , PromoteDelete
            End If
        Next X

        Call PopulateLinkedFirmsListView
        If lvAvailableItems.ListItems.Count > 0 Then
            Call PopulateAvailableFirmsListView
        End If
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
' TW 14/02/2007 EP2_1366 End
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
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
    cmdSearch.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    On Error GoTo Failed
    SetReturnCode MSGFailure
        
    Set m_clsAppointmentsTable = New AppointmentsTable
    SetColumnHeaders
    
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
    
    clsHeader.nWidth = 20
    clsHeader.sName = "Company Name"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 10
    clsHeader.sName = "FSA Ref"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 50
    clsHeader.sName = "Address"
    colHeaders.Add clsHeader
    
    lvAvailableItems.AddHeadings colHeaders
    lvLinkedItems.AddHeadings colHeaders

End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub SetEditState()
    
    On Error GoTo Failed
    strARFirmID = m_colKeys(1)
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
End Sub

Private Sub PopulateLinkedFirmsListView()
    On Error GoTo Failed
    
' TW 13/02/2007 EP2_1334
'    strSearch = "SELECT UNITID, P.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
'                "FSAREF, REPLACE(REPLACE(LTRIM(" & _
'                "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
'                "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
'                "From PRINCIPALFIRM P INNER JOIN APPOINTMENTS I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
'                "WHERE ISNULL(PACKAGERINDICATOR, 0) = 0 AND ARFIRMID = '" & strARFirmID & "' AND STATUSCODE = 'Registered'"

    strSearch = "SELECT UNITID, P.PRINCIPALFIRMID AS ID, PRINCIPALFIRMNAME AS COMPANYNAME, " & _
                "FSAREF, REPLACE(REPLACE(LTRIM(" & _
                "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS " & _
                "From PRINCIPALFIRM P INNER JOIN APPOINTMENTS I ON P.PRINCIPALFIRMID = I.PRINCIPALFIRMID " & _
                "WHERE ISNULL(PACKAGERINDICATOR, 0) = 0 AND ARFIRMID = '" & strARFirmID & "'"
' TW 13/02/2007 EP2_1334 End
    
    
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


Private Sub txtUnitID_LostFocus()
    EnableSearch txtUnitID
End Sub


