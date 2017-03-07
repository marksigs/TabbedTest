VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX"
Begin VB.Form frmNewCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmNewCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin MSGOCX.MSGEditBox TxtValidationType 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         MaxValue        =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   20
      End
      Begin MSGOCX.MSGEditBox TxtValueName 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         MaxValue        =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   50
      End
      Begin MSGOCX.MSGEditBox TxtValueId 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         TextType        =   6
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         MaxValue        =   "32767"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   5
      End
      Begin MSGOCX.MSGEditBox TxtGroupName 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Enabled         =   0   'False
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
      Begin VB.Label LblValidationType 
         AutoSize        =   -1  'True
         Caption         =   "ValidationType"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label LblValueName 
         AutoSize        =   -1  'True
         Caption         =   "ValueName"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label LblValueId 
         AutoSize        =   -1  'True
         Caption         =   "ValueId"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   540
      End
      Begin VB.Label LblGroupName 
         AutoSize        =   -1  'True
         Caption         =   "GroupName"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNewCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmNewCombo
' Description   :   Form which allows the user add Combo Values / Combo Validation
'
' Prog      Date        AQR     Description
' MV        15/01/2003  BM0085  Created
' MV        15/01/03    BM0085  Amended to support Creating audit Records
' MV        04/04/2003  BM0402  Amended CreateComboValidationAuditRecord() and CreateComboValueAuditRecord()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsComboValueTable As ComboValueTable
Private m_clsComboValidationTable As ComboValidationTable
Private m_colKeys As Collection
Private m_sTableName As String
Private blnFormLoaded As Boolean
Private blnIsChanged As Boolean
Private m_clsComboGroup As ComboValueGroupTable
Private Function CreateComboValueRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValueName As String) As Boolean
    
    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + "INSERT INTO COMBOVALUE (GROUPNAME,VALUEID,VALUENAME) "
    sSQL = sSQL + " VALUES ( "
    sSQL = sSQL + "'" & sGroupName & "', "
    sSQL = sSQL + sValueId & ", "
    
    If sValueName = "" Then
        sSQL = sSQL + "Null" + ") "
    Else
        sSQL = sSQL + "'" & sValueName & "' ) "
    End If
    
    g_clsDataAccess.ExecuteCommand sSQL
        
    CreateComboValueRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function


Private Function CreateComboValidationRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValidationType As String) As Boolean
    
    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + "INSERT INTO COMBOVALIDATION (GROUPNAME,VALUEID,VALIDATIONTYPE) "
    sSQL = sSQL + " VALUES ( "
    sSQL = sSQL + "'" & sGroupName & "', "
    sSQL = sSQL + sValueId & ",'"
    sSQL = sSQL + sValidationType & "' )"
    
    g_clsDataAccess.ExecuteCommand sSQL
        
    CreateComboValidationRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function



Private Function DeleteComboValidationRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValidationType As String) As Boolean
    
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim bRet  As Boolean
    
    
    sSQL = ""
    sSQL = sSQL + " DELETE FROM COMBOVALIDATION WHERE  GROUPNAME = '" + sGroupName + "' "
    sSQL = sSQL + " AND VALUEID  = " + sValueId
    sSQL = sSQL + " AND VALIDATIONTYPE = '" + sValidationType + "' "
            
    g_clsDataAccess.ExecuteCommand sSQL
        
    CreateComboValidationAuditRecord sGroupName, sValueId, g_sSupervisorUser, sValidationType, "", "D"
        
    DeleteComboValidationRecord = True
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
        
End Function

Private Function DeleteComboValueRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValueName As String) As Boolean
    
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim ComboValidationRS As ADODB.Recordset
    
    
'Stage1:
    
    sSQL = ""
    sSQL = sSQL + " SELECT * FROM COMBOVALIDATION WHERE  GROUPNAME = '" + sGroupName + "' "
    sSQL = sSQL + " AND VALUEID  = " + sValueId
    
    Set ComboValidationRS = g_clsDataAccess.GetTableData("COMBOVALIDATION", sSQL)
       
    If ComboValidationRS.RecordCount > 0 Then
        
        While Not ComboValidationRS.EOF
            
            sSQL = ""
            sSQL = sSQL + " DELETE FROM COMBOVALIDATION WHERE  GROUPNAME = '" + sGroupName + "' "
            sSQL = sSQL + " AND VALUEID  = " + sValueId
            sSQL = sSQL + " AND VALIDATIONTYPE = '" + ComboValidationRS.fields("VALIDATIONTYPE").Value + "' "
            
            g_clsDataAccess.ExecuteCommand sSQL
            
            CreateComboValidationAuditRecord sGroupName, sValueId, g_sSupervisorUser, ComboValidationRS.fields("VALIDATIONTYPE").Value, "", "D"
            
            ComboValidationRS.MoveNext
            
        Wend
        
    End If
    
'Stage2:

    sSQL = ""
    sSQL = sSQL + " DELETE FROM COMBOVALUE WHERE  GROUPNAME = '" + sGroupName + "' "
    sSQL = sSQL + " AND VALUEID  = " + sValueId
    sSQL = sSQL + " AND VALUENAME  = '" + sValueName + "' "
    
    g_clsDataAccess.ExecuteCommand sSQL
    
    CreateComboValueAuditRecord sGroupName, sValueId, g_sSupervisorUser, sValueName, "", "D"
        
    DeleteComboValueRecord = True
        
    Exit Function
        
Failed:

    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function DoesRecordExists(ByVal sfrmEditComboMode As String) As Boolean
        
    On Error GoTo Failed
    Dim sSQL As String
    Dim TempRs  As ADODB.Recordset
    sSQL = ""
    If sfrmEditComboMode = "COMBOVALUE" Then
            
        If frmEditCombo.m_sfrmEditComboOperationMode = "ADD" Then
            sSQL = ""
            sSQL = sSQL + " SELECT * FROM COMBOVALUE WHERE GROUPNAME = '" & TxtGroupName.Text
            sSQL = sSQL + "' AND VALUEID = " & TxtValueId.Text
        End If
            
    ElseIf sfrmEditComboMode = "COMBOVALIDATION" Then
            
        sSQL = ""
        sSQL = sSQL + " SELECT * FROM COMBOVALIDATION WHERE GROUPNAME = '" & TxtGroupName.Text & "' AND VALUEID = " & TxtValueId.Text
        sSQL = sSQL + " AND VALIDATIONTYPE = '" & TxtValidationType.Text & "' "
        
    End If
        
    If sSQL <> "" Then
        Set TempRs = g_clsDataAccess.GetTableData(frmEditCombo.m_sfrmEditComboMode, sSQL)
        If TempRs.RecordCount <= 0 Or IsNull(TempRs) Then
            DoesRecordExists = False
        Else
            DoesRecordExists = True
        End If
    Else
        DoesRecordExists = False
    End If
    
    Exit Function
    
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Private Sub SetScreenData()
    
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As New Collection
    
    If frmEditCombo.m_sfrmEditComboMode = "COMBOVALUE" Then
        
        m_sTableName = "COMBOVALUE"
        
        Set clsTableAccess = New ComboValueTable
        
        colMatchValues.Add frmEditCombo.m_sGroupName
        colMatchValues.Add frmEditCombo.m_sValueId
        
        clsTableAccess.SetKeyMatchValues colMatchValues
        
        If frmEditCombo.m_sfrmEditComboOperationMode = "ADD" Then
            Me.Caption = "Add Combo Value"
            Me.LblValidationType.Visible = False
            Me.TxtValidationType.Visible = False
            Me.LblValueName.Visible = True
            Me.TxtValueName.Visible = True
            Me.cmdDelete.Enabled = False
        End If
        
        If frmEditCombo.m_sfrmEditComboOperationMode = "EDIT" Then
            Me.Caption = "Edit Combo Value"
            Me.LblValidationType.Visible = False
            Me.TxtValidationType.Visible = False
            Me.LblValueName.Visible = True
            Me.TxtValueName.Visible = True
            Me.cmdDelete.Enabled = True
            Me.TxtValueId.Enabled = False
            Me.TxtValueId.Text = frmEditCombo.m_sValueId
            Me.TxtValueName.Text = frmEditCombo.m_sValueName
        End If
        
    ElseIf frmEditCombo.m_sfrmEditComboMode = "COMBOVALIDATION" Then
            
        m_sTableName = "COMBOVALIDATION"
        
        Set clsTableAccess = New ComboValidationTable
        
        colMatchValues.Add frmEditCombo.m_sGroupName
        colMatchValues.Add frmEditCombo.m_sValueId
        colMatchValues.Add frmEditCombo.m_sValidationType
        
        clsTableAccess.SetKeyMatchValues colMatchValues
          
        If frmEditCombo.m_sfrmEditComboOperationMode = "ADD" Then
            Me.Caption = "Add Combo Validation"
            Me.TxtValueId.Enabled = False
            Me.LblValueName.Visible = False
            Me.TxtValueName.Visible = False
            Me.TxtValueId.Enabled = False
            Me.LblValidationType.Visible = True
            Me.TxtValidationType.Visible = True
            Me.TxtValueId.Text = frmEditCombo.m_sValueId
            Me.cmdDelete.Enabled = False
        End If
        
        If frmEditCombo.m_sfrmEditComboOperationMode = "EDIT" Then
            Me.Caption = "Edit Combo Validation"
            Me.LblValueName.Visible = False
            Me.TxtValueName.Visible = False
            Me.TxtValueId.Text = frmEditCombo.m_sValueId
            'Me.TxtValueId.Enabled = False
            Me.TxtValueId.Enabled = False
            Me.LblValidationType.Visible = True
            Me.TxtValidationType.Visible = True
            Me.TxtValidationType.Text = frmEditCombo.m_sValidationType
        End If
        
    End If
    
End Sub

Private Function UpdateComboValueRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValueName As String, ByVal sPrevValueName As String) As Boolean
    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + " UPDATE COMBOVALUE "
    sSQL = sSQL + " SET VALUENAME = '" + sValueName + "' "
    sSQL = sSQL + " WHERE GROUPNAME = '" + sGroupName + "' AND VALUEID = " + sValueId
    If sPrevValueName <> "" Then
        sSQL = sSQL + " AND VALUENAME = '" + sPrevValueName + "' "
    End If
    
    g_clsDataAccess.ExecuteCommand sSQL
        
    UpdateComboValueRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Private Function UpdateComboValidationRecord(ByVal sGroupName As String, ByVal sValueId As String, ByVal sValidationType As String, ByVal sPrevValidationType As String) As Boolean
    
    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + " UPDATE COMBOVALIDATION "
    sSQL = sSQL + " SET VALIDATIONTYPE = '" + sValidationType + "' "
    sSQL = sSQL + " WHERE GROUPNAME = '" + sGroupName + "' AND VALUEID = " + sValueId
    sSQL = sSQL + " AND VALIDATIONTYPE = '" + sPrevValidationType + "' "
    
    g_clsDataAccess.ExecuteCommand sSQL
        
    UpdateComboValidationRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
Public Function ValidateData(ByVal sfrmEditComboMode As String) As Boolean
    
    On Error GoTo Failed
    
    If sfrmEditComboMode = "COMBOVALUE" Then
        If Len(TxtValueId.Text) <= 0 Then
            g_clsErrorHandling.DisplayError "ValueId is empty"
            ValidateData = False
            TxtValueId.SetFocus
            Exit Function
        End If
        If Len(TxtValueName.Text) <= 0 Then
            g_clsErrorHandling.DisplayError "ValueName is empty"
            ValidateData = False
            TxtValueName.SetFocus
            Exit Function
        End If
    ElseIf sfrmEditComboMode = "COMBOVALIDATION" Then
        If Len(TxtValidationType.Text) <= 0 Then
            g_clsErrorHandling.DisplayError "ValidationType is empty"
            ValidateData = False
            TxtValidationType.SetFocus
            Exit Function
        End If
    End If
    
    ValidateData = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    On Error GoTo Failed
    
    Dim nResponse As Integer
    Dim sSQL As String
    Dim bRet  As Boolean
    Dim sValueName As String
    Dim sValidationType As String
    
    nResponse = MsgBox("Delete the selected record?", vbQuestion + vbYesNo)
            
    If nResponse = vbYes Then
    
        If m_sTableName = "COMBOVALUE" Then
            
            If blnIsChanged Or TxtValueName.Text = "" Then
               sValueName = frmEditCombo.m_sValueName
            Else
               sValueName = TxtValueName.Text
            End If
        
            bRet = DeleteComboValueRecord(TxtGroupName.Text, TxtValueId.Text, sValueName)
            If bRet = False Then
                g_clsErrorHandling.RaiseError errGeneralError, "Error in Deleting ComboValueRecord", ErrorUser
                Exit Sub
            End If
        
        ElseIf m_sTableName = "COMBOVALIDATION" Then
            
            If blnIsChanged Or TxtValidationType.Text = "" Then
                sValidationType = frmEditCombo.m_sValidationType
            Else
                sValidationType = TxtValidationType.Text
            End If
        
            bRet = DeleteComboValidationRecord(TxtGroupName.Text, TxtValueId.Text, sValidationType)
            If bRet = False Then
                g_clsErrorHandling.RaiseError errGeneralError, "Error in Deleting ComboValidation Record", ErrorUser
                Exit Sub
            End If
        
        End If
            
        frmEditCombo.blnSuccess = True
        Unload Me
    End If
     
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub cmdOK_Click()
    
    On Error GoTo Failed
    
    Dim TempRs As ADODB.Recordset
    Dim bRet  As Boolean
        
   
    
    If blnIsChanged Then
        
        bRet = ValidateData(frmEditCombo.m_sfrmEditComboMode)
        
        If bRet Then
            
            bRet = DoesRecordExists(frmEditCombo.m_sfrmEditComboMode)
            
            If bRet Then
                
                If frmEditCombo.m_sfrmEditComboMode = "COMBOVALUE" Then
                    g_clsErrorHandling.DisplayError "Entry for " + TxtValueId.Text + " already exists"
                    Me.TxtValueId.SetFocus
                ElseIf frmEditCombo.m_sfrmEditComboMode = "COMBOVALIDATION" Then
                    g_clsErrorHandling.DisplayError "Entry for " + TxtValidationType.Text + " already exists"
                    Me.TxtValidationType.SetFocus
                End If
            
            Else
                
                SaveScreenData
                frmEditCombo.blnSuccess = True
                Unload Me
                                
            End If
        
        End If
    Else
        Unload Me
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function CreateComboValidationAuditRecord(ByVal sGroupName As String, ByVal sValueId As String, _
                        ByVal sChangeUser As String, ByVal sValidationType As String, _
                        ByVal sPrevValidationType As String, ByVal sOperation As String, Optional strDBName As String) As Boolean

    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + "INSERT INTO COMBOVALIDATIONAUDIT (GROUPNAME,VALUEID,AUDITDATE,CHANGEUSER,VALIDATIONTYPE,PREVVALIDATIONTYPE,OPERATION) "
    sSQL = sSQL + " VALUES ( "
    sSQL = sSQL + "'" & sGroupName & "', "
    sSQL = sSQL + sValueId & ", "
    sSQL = sSQL + g_clsSQLAssistSP.GetSystemDate & ", "
    sSQL = sSQL + "'" & sChangeUser & "', "
    
    If sValidationType = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sValidationType & "', "
    End If
    
    If sPrevValidationType = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sPrevValidationType & "', "
    End If
    
    
    sSQL = sSQL + "'" + sOperation + "' ) "
    
    g_clsDataAccess.ExecuteCommand sSQL, strDBName
        
    CreateComboValidationAuditRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    

End Function


Public Function CreateComboValueAuditRecord(ByVal sGroupName As String, ByVal sValueId As String, _
                        ByVal sChangeUser As String, ByVal sValueName As String, ByVal sPrevValueName As String, _
                        ByVal sOperation As String, Optional strDBName As String) As Boolean
    
    On Error GoTo Failed

    Dim sSQL As String

    sSQL = ""
    sSQL = sSQL + "INSERT INTO COMBOVALUEAUDIT (GROUPNAME,VALUEID,AUDITDATE,CHANGEUSER,VALUENAME,PREVVALUENAME,OPERATION) "
    sSQL = sSQL + " VALUES ( "
    sSQL = sSQL + "'" & sGroupName & "', "
    sSQL = sSQL + sValueId & ", "
    sSQL = sSQL + g_clsSQLAssistSP.GetSystemDate & ", "
    sSQL = sSQL + "'" & sChangeUser & "', "
    
    If sValueName = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sValueName & "', "
    End If
    
    If sPrevValueName = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sPrevValueName & "', "
    End If

    sSQL = sSQL + "'" + sOperation + "' ) "
    
    g_clsDataAccess.ExecuteCommand sSQL, strDBName
    
    CreateComboValueAuditRecord = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function





Private Sub Form_Load()
    
    On Error GoTo Failed
    blnFormLoaded = False
    blnIsChanged = False
    Set m_clsComboValueTable = New ComboValueTable
    Set m_clsComboValidationTable = New ComboValidationTable
    Set m_clsComboGroup = New ComboValueGroupTable
    
    TxtGroupName.Text = frmEditCombo.m_sGroupName
    
    SetScreenData
    blnFormLoaded = True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
        
End Sub

Private Sub DoGroupUpdate()
'    On Error GoTo Failed
'    Dim bRet As Boolean
'    Dim sGroup As String
'    Dim sNotes As String
'    Dim rs As ADODB.Recordset
'    Dim clsTableAccess As TableAccess
'    Dim bUpdate As Boolean
'    Dim col As New Collection
'
'    On Error GoTo Failed
'    sGroup = TxtGroupName.Text
'    'sNotes = txtComboNotes.Text
'    bUpdate = False
'
'    Set clsTableAccess = m_clsComboGroup
'
'    If Len(sGroup) > 0 Then
'        ' Do the group duplicate search if adding (it's new) or if it's changed
'        'If m_bIsEdit = False Or (m_bIsEdit = True And m_sGroupNameOrig <> sGroup) Then
'            DoesGroupExist
'        'End If
'
'        'If m_bIsEdit = False Then
'            ' Adding, so add a new record
'            g_clsFormProcessing.CreateNewRecord clsTableAccess
'        'End If
'
'        m_clsComboGroup.SetGroupName sGroup
'        m_clsComboGroup.SetGroupNote sNotes
'        bUpdate = True
'    End If
'
'    If bUpdate = True Then
'        clsTableAccess.Update
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub DoesGroupExist()
'    On Error GoTo Failed
'
'    Dim bExists As Boolean
'    Dim sGroup As String
'    Dim col As New Collection
'    Dim clsTableAccess As TableAccess
'
'    sGroup = TxtGroupName.Text
'    Set clsTableAccess = m_clsComboGroup
'
'    If Len(sGroup) > 0 Then
'        col.Add sGroup
'        clsTableAccess.SetKeyMatchValues col
'
'        bExists = Not clsTableAccess.DoesRecordExist(col)
'
'        If bExists = False Then
'            Me.TxtGroupName.SetFocus
'            g_clsErrorHandling.RaiseError errGeneralError, "Entry for " + sGroup + " already exists"
'        End If
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SaveScreenData()

    On Error GoTo Failed
    Dim vTmp As Variant
    
    Dim clsTableAccess As TableAccess
    
    If frmEditCombo.m_sfrmEditComboMode = "COMBOVALUE" Then
        
        If frmEditCombo.m_sfrmEditComboOperationMode = "ADD" Then
        
            CreateComboValueRecord TxtGroupName.Text, TxtValueId.Text, TxtValueName.Text
            CreateComboValueAuditRecord TxtGroupName.Text, TxtValueId.Text, g_sSupervisorUser, TxtValueName.Text, "", "C"
            
        ElseIf frmEditCombo.m_sfrmEditComboOperationMode = "EDIT" Then
        
            UpdateComboValueRecord TxtGroupName.Text, TxtValueId.Text, TxtValueName.Text, frmEditCombo.m_sValueName
            CreateComboValueAuditRecord TxtGroupName.Text, TxtValueId.Text, g_sSupervisorUser, TxtValueName.Text, frmEditCombo.m_sValueName, "U"
            
        End If
        
    ElseIf frmEditCombo.m_sfrmEditComboMode = "COMBOVALIDATION" Then
        
        If frmEditCombo.m_sfrmEditComboOperationMode = "ADD" Then
        
            CreateComboValidationRecord TxtGroupName.Text, TxtValueId.Text, TxtValidationType.Text
            CreateComboValidationAuditRecord TxtGroupName.Text, TxtValueId.Text, g_sSupervisorUser, TxtValidationType.Text, "", "C"
            
        ElseIf frmEditCombo.m_sfrmEditComboOperationMode = "EDIT" Then
        
            UpdateComboValidationRecord TxtGroupName.Text, TxtValueId.Text, TxtValidationType.Text, frmEditCombo.m_sValidationType
            CreateComboValidationAuditRecord TxtGroupName.Text, TxtValueId.Text, g_sSupervisorUser, TxtValidationType.Text, frmEditCombo.m_sValidationType, "U"
            
        End If
    
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub









Private Sub TxtValidationType_Change()
'    If blnFormLoaded Then
'        blnIsChanged = True
'    End If
End Sub

Private Sub TxtValidationType_KeyPress(KeyAscii As Integer)
    If blnFormLoaded Then
        blnIsChanged = True
    End If
End Sub

Private Sub TxtValueId_KeyPress(KeyAscii As Integer)
    If blnFormLoaded Then
        blnIsChanged = True
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtValueName_KeyPress(KeyAscii As Integer)
    If blnFormLoaded Then
        blnIsChanged = True
    End If
End Sub


