VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl MSGDataGrid 
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   Enabled         =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5400
   Begin VB.Timer tmrCheckKeyPress 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   0
   End
   Begin MSGOCX.MSGMulti txtStringField 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      BorderStyle     =   0
   End
   Begin MSGOCX.MSGEditBox txtDateField 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Mask            =   "##/##/####"
      TextType        =   1
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
      MaxLength       =   10
      BorderStyle     =   0
   End
   Begin MSGOCX.MSGComboBox combo1 
      Height          =   315
      Left            =   4020
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Text            =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3420
      Top             =   -180
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2835
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   "HH:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MSGDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' ActiveX control that performs extra processing for the datagrid control, such as managing mandatory data,
' and processing of fields (such as allowing combo boxes and text edits to be used)
'
' Change history
' Prog      Date        Description
' AA        06/12/00    Added CheckDayOfWeek
' AA/DJP    08/12/00    Added code to show textbox for all field edits. Previously
'                       only dates were editing with a textbox, but now we are using
'                       a textbox for all edits. This is because the grid itself has
'                       a number of bugs which mean the grid isn't drawn correctly
'                       under some circumstances.
' AA/DJP    04/01/01    Make txtDateField_LostFocus not display textbox error, and
'                       add debugging routine.
' AA        05/01/01    Added IsColumnEnabled as part of Distribution Channel Enhancement
' AA        27/02/01    AQR SYS1893
' AA        09/03/01    Added tmrCheckKeyPress, as a fix for "Client Site not available"
' AA        09/03/01    Added code to function HandleEdit Box to resize the text box SMALLER
'                       rather than the same size as the grid. This is due to the texbox
'                       border being removed.
' DJP/AA    16/03/01    Updated IsTxtDateField to correctly determine if the grid text field
'                       is a date or not (use typename)
' DJP       23/03/01    HandleTermination
' DJP       18/06/01    SQL Server port. Also, when adding a new record move to the first field we can see
' STB       03/12/01    SYS3264 - AllowDelete property now effects underlying datagrid.
' SA        20/12/01    SYS3542 - In Validatechar Allow minus amounts
' CL        30/02/02    SYS4787 - Changed datagrid to stop multiple appearances.
' GHun      16/11/2005  MAR312 Expose Bound recordset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_clsErrorHandling As ErrorHandling
Private m_nLastRow As Variant
Private m_nLastCol As Integer
Private m_colFields As Collection
Private m_sRegSetting As String
Private m_colColStatus As Collection
Private m_rsData As adodb.Recordset
Private m_nComboRow As Integer
Private m_nComboCol As Integer
Private m_bSetTabs As Boolean
Private m_nTextRow As Integer
Private m_nTextCol As Integer
Private m_bAddingRecord As Boolean
Private m_bValidating As Boolean
Private m_bLeavingGrid As Boolean
Private m_bUpdating As Boolean
Private txtGridField As Control
Private m_bControlLostFocus As Boolean

' Member variable that will contain a count of enabled and unlocked fields
Private m_nEnabledCols As Integer

'Default Property Values:
Const m_def_AllowDelete = True
Const m_def_AllowAdd = True
Const m_def_SelStartCol = 0
Const m_def_SelEndCol = 0
Const m_def_Row = 0
Const m_def_Col = 0
Const m_def_Text = "0"
Const m_def_ContainerHwnd = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Private m_sRegLocation As String

'Property Variables:
Dim m_AllowDelete As Boolean
Dim m_AllowAdd As Boolean
Dim m_SelStartCol As Integer
Dim m_SelEndCol As Integer
Dim m_Row As Integer
Dim m_Col As Integer
Dim m_Text As String
Dim m_ContainerHwnd As Long
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

' Constants
Private Const COL_WIDTH As String = "GridFieldWidth"
Private Const TIMER_COMBO_FOCUS = "ComboFocus"
Private Const TIMER_GRID_FOCUS = "GridFocus"
Private Const TIMER_TEXT_FOCUS = "TextFocus"
Private Const TIMER_DELETE_STATE = "TimerTabs"
Private Const TYPE_SHORT = 5
Private Const TYPE_LONG = 10
Private Const TYPE_DOUBLE = 38
Private Const TYPE_DOUBLE_SQL_SERVER = 15
Private Const TYPE_BOOLEAN = 1
Private Const MAX_SHORT = 32767
Private Const MAX_LONG = 2147483647
Private Const MAX_DOUBLE = 2147483647
Private Const MIN_DOUBLE = -2147483647
Private m_clsGenAssist As GeneralAssist
Private Const ALT_KEY_CODE = 164

'Event Declarations:
Event BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer) 'MappingInfo=DataGrid1,DataGrid1,-1,BeforeColEdit
Event BeforeDelete(Cancel As Integer) 'MappingInfo=DataGrid1,DataGrid1,-1,BeforeDelete
Event BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)

Event RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Event AfterAdd()
Event AddButtonClicked(bCancel As Boolean, nCurrentRow As Integer)
Event AfterColumnUpdated(ByVal columnIndex As Long, ByVal nCurrentRow As Long)
Event AfterRowInsert(ByVal nCurrentRow As Long)
Event AfterRowUpdated(ByVal nCurrentRow As Long)
Event BeforeColumnUpdated(ByVal columnIndex As Long, ByVal oldValue As Variant, Cancel As Integer)
Event BeforeRowInsert(ByVal nCurrentRow As Long, ByVal Cancel As Integer)
Event BeforeRowUpdate(ByVal nCurrentRow As Long, ByVal Cancel As Integer)
Event ColumnEdit(ByVal columnIndex As Long, ByVal nCurrentRow As Long)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History:
'
'Prog   Date        AQR     Description
'AW     20/05/2002  BM087   Override cmdDelete_Click()
'MV     07/01/2003  BM0085  Amended AddNewRecord()
'MV     15/01/2003  BM0085  Added New function GetcurrentRow()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideAll()
    On Error GoTo Failed
    
    txtDateField.Visible = False
    txtStringField.Visible = False
    combo1.Visible = False
    UserControl.Refresh
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Function ValidateRows() As Boolean
    Dim bRet As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    
    HideAll

    bRet = True
    
    m_bValidating = True

    nCount = Rows()
            
    nThisItem = 0
    
    While bRet = True And nThisItem < nCount
        bRet = ValidateRow(nThisItem)

        nThisItem = nThisItem + 1
    Wend
    m_bValidating = False

    ValidateRows = bRet
End Function

Public Function ValidateRow(nRow As Integer) As Boolean
    Dim bRet As Boolean
    Dim bValidating As Boolean
    Dim Field As FieldData
    Dim sError As String
    Dim nFirst As Integer
    Const ROW_VALID As Integer = -1
    
    On Error GoTo Failed
    HideAll
    bRet = True
    
    ' We don't want any of the Row Change events to occur here, so disable it
    bValidating = m_bValidating
    m_bValidating = True
    
    sError = "Data Required"
    nFirst = ROW_VALID
    
    EnableTabs ' So we can put focus into the datagrid fields

    If txtGridField.Visible = True Then
        If IsDateValid(txtGridField.Text) = False Then
            ShowTextBox False, True
        Else
            If IstxtDateField = True Then
                bRet = txtGridField.ValidateData()
            Else
                bRet = True
            End If
                    
            If bRet = False Then
                SetTextFocus
            End If
        End If
    End If
    
    If bRet = True Then
        ' Move to the row we want
        m_rsData.Move nRow, 1
        
        If Not m_colColStatus Is Nothing Then
            ' Variant so it can be NULL
            Dim sVal As Variant
            Dim col As Column

            On Error Resume Next
            
            ' Now loop through each column in the row
            For Each col In DataGrid1.Columns
                DataGrid1.col = col.ColIndex
                Field = m_colColStatus(GetKey(col.ColIndex))
                
                ' If it's required and it's visible we want to check that data has been entered.
                If Field.bRequired = True And col.Visible = True Then
                    sVal = col.Text
                    If Len(sVal) = 0 Then
                        If nFirst = ROW_VALID Then
                            nFirst = col.ColIndex
                        End If
                        
                        If Len(Field.sError) > 0 Then
                            sError = Field.sError
                        End If
                        
                        'MsgBox sError, vbCritical
                        
                        bRet = False
                    End If
                End If
            Next
            
            ' If we have an error, put the focus into the offending column
            If bRet = False Then
                DataGrid1.Row = nRow
                If nFirst <> ROW_VALID Then
                    EditActive nFirst
                End If
            End If
        Else
            m_clsErrorHandling.RaiseError errGeneralError, "MSGDataGrid: Can't validate row row, column collection is null"
        End If
    End If
    
    If bValidating = False Then
        m_bValidating = False
    End If
    
    ValidateRow = bRet
    
    Exit Function
Failed:
    ValidateRow = False
    m_clsErrorHandling.DisplayError
End Function

Public Function GetCurrentRow() As Integer
    GetCurrentRow = DataGrid1.Row + DataGrid1.FirstRow - 1
End Function

Private Sub AddNewRecord(Optional bSetFocus As Boolean = True)
    
    Dim bCancel As Boolean
    Dim nRow As Long
    Dim bFirst As Boolean
    Dim nFirstCol As Integer
    Dim Field As FieldData
    
    On Error GoTo Failed

    bCancel = False
    nRow = DataGrid1.Row
    bFirst = True
    nFirstCol = -1
    
    'MV - 07/01/2003 - BM0085 - CC002
    'If nRow >= 0 Then
        RaiseEvent BeforeAdd(bCancel, DataGrid1.Row)
    'End If

    If bCancel = False Then
        ' Need to add a new row first.
        On Error GoTo Failed
        If Not m_colColStatus Is Nothing Then
            If Not m_rsData Is Nothing Then
                m_bAddingRecord = True
                m_rsData.AddNew

                ' For some reason I've yet to figure out, if the user deletes the last row, so there
                ' are no rows left in the grid, then adds a new record, the record count after the AddNew
                ' above is zero! The only way I can make it work is by cloning the recordset, resetting it
                ' to the datagrid, then setting all the headers again.
                
                If m_rsData.RecordCount = 0 Then
                    ' ?????? Zero?
                    Set m_rsData = m_rsData.Clone
                    Set DataGrid1.DataSource = m_rsData
                    SetColumns m_colFields, m_sRegSetting
                    ' All ok now...
                End If
                
                ' Now, need to take a look at each column, and if it's hidden we need to copy it's
                ' default value in.

                Dim sVal As Variant
                Dim col As Column

                On Error Resume Next

                For Each col In DataGrid1.Columns
' AQR SYS1893
' Added functionaility to have default values on VISIBLE Columns
                    Field = m_colColStatus(GetKey(col.ColIndex))
                    
                    If col.Visible = False Then
                        col.Visible = True
                        DataGrid1.col = col.ColIndex

                        col.Visible = False
                    Else
                        If bFirst = True Then
                            nFirstCol = col.ColIndex
                            bFirst = False
                        End If
                    End If
                    sVal = Field.sDefault
                        
                    If Len(sVal) > 0 Then
                        col.Value = sVal
                    End If
                Next
            Else
                MsgBox "Add: DataSource is empty", vbCritical
            End If
        Else
            MsgBox "MSGDataGrid: Can't add row, column collection is null", vbCritical
        End If
    End If

    If bFirst = False And bSetFocus = True Then
        EditActive nFirstCol
        DataGrid1.LeftCol = nFirstCol
    End If
    m_bAddingRecord = False
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Failed
    
    AddNewRecord
    '*=Event Raised to control from the source application
    'AQR 597
    RaiseEvent AddButtonClicked(False, Me.Rows - 1)
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub


Public Function SetAtRowCol(nRow As Integer, nCol As Integer, vVal As Variant) As Boolean
    Dim bNull As Boolean
    Dim nAttribute As Long
    Dim nPrevious As Long

    On Error GoTo Failed
    
    nPrevious = m_rsData.AbsolutePosition - 1
    m_rsData.Move nRow, 1
        
    nAttribute = m_rsData.Fields(nCol).Attributes
    
    ' If the field we are trying to update accepts null values and the value we are
    ' passing in is either null itself, or has an empty string value, set the field
    ' to null. If the field does not accept nulls, leave it alone.
    bNull = CBool(nAttribute And adFldIsNullable)
    
    If Not IsNull(vVal) Then
        If Len(vVal) = 0 Then
            vVal = Null
        End If
    End If
    
    If (Not IsNull(vVal)) Or (IsNull(vVal) And bNull) Then
        m_rsData(nCol).Value = vVal
    End If
    
    m_rsData.Move nPrevious, 1
    
    SetAtRowCol = True
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Function GetAtRowCol(nRow As Integer, nCol As Integer) As String
    Dim sVal As Variant
    Dim colVal As Column
    
    On Error GoTo Failed
    
    m_rsData.Move nRow, 1
        
    Set colVal = DataGrid1.Columns(nCol)
    
    sVal = colVal.Text
    
    If IsNull(sVal) Then
        sVal = ""
    End If
    
    GetAtRowCol = sVal
    
    Exit Function
Failed:
    m_clsErrorHandling.DisplayError
    GetAtRowCol = ""
End Function

Private Sub cmdDelete_Click()

    On Error GoTo Failed
    
    'AW 20/05/02    BM087
    Dim nResponse As Integer
    nResponse = MsgBox("Delete the selected row(s)?", vbQuestion + vbYesNo)
        
    If nResponse = vbNo Then
        Exit Sub
    End If
    ' Due to problems associated with deleting rows from the datagrid, the way that works
    ' best is to actually make the grid itself do the delete - this is done by just pressing
    ' the DELETE key on the grid!
    
    DataGrid1.SetFocus
    SendKeys "{DEL}"
    
    ' The state of the delete button can't be set here - it has to be done after all deletion processing
    ' has taken place, so setup a timer event to handle it.
    Timer1.Enabled = True
    Timer1.Tag = TIMER_DELETE_STATE
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub HandleDeleteState(bEnable As Boolean)
    On Error GoTo Failed
    If Not m_rsData Is Nothing Then
        If m_rsData.EOF = False And m_rsData.BOF = False Then
            cmdDelete.Enabled = bEnable
        Else
            cmdDelete.Enabled = False
        End If
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub Combo1_Click()
    On Error GoTo Failed
    
    UpdateComboValue
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    
    Dim bColEnabled As Boolean
    
    If KeyAscii = vbKeyRight Then
        HandleKeyMove "{RIGHT}", KeyAscii
    ElseIf KeyAscii = vbKeyLeft Then
        HandleKeyMove "{LEFT}", KeyAscii
    ElseIf KeyAscii = vbKeyTab Then
        
        ' If it's not the end of a row, send the tab, otherwise tab onto the next control
        bColEnabled = IsColumnEnabled(m_nComboCol + 1)
        If m_nComboCol + 1 > m_nEnabledCols Then
            m_bLeavingGrid = True
            
            combo1.Visible = False
            EnableTabs
            HandleKeyMove "{TAB}", KeyAscii, False
        ElseIf Not bColEnabled And m_nComboCol + 1 < m_nEnabledCols Then
            HandleKeyMove "{TAB}{TAB}", KeyAscii
        Else
            HandleKeyMove "{TAB}", KeyAscii
        End If
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Function FindFirstVisibleColumn() As Integer
    On Error GoTo Failed
    Dim col As Column
    Dim nFirstVisible As Integer
    
    nFirstVisible = -1
    
    For Each col In DataGrid1.Columns
        If col.Visible = True Then
            nFirstVisible = col.ColIndex
            Exit For
        End If
    Next
    
    FindFirstVisibleColumn = nFirstVisible
    
    Exit Function
Failed:
    FindFirstVisibleColumn = nFirstVisible
End Function

Private Sub ActivateFirstColumn()
    On Error GoTo Failed
    Dim nFirstActive As Integer
    
    If DataGrid1.Columns(DataGrid1.col).Visible = False Then
        nFirstActive = FindFirstVisibleColumn()
        EditActive nFirstActive
    End If
    
    Exit Sub
Failed:
End Sub

'*=Expose All Required events to the source application

Private Sub DataGrid1_AfterInsert()
    '*=Implement here
    RaiseEvent AfterRowInsert(DataGrid1.Row)
    
End Sub

Private Sub DataGrid1_AfterUpdate()
    '*=Implement here
    RaiseEvent AfterRowUpdated(DataGrid1.Row)
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, oldValue As Variant, Cancel As Integer)
    '*=Implement here
    RaiseEvent BeforeColumnUpdated(ColIndex, oldValue, Cancel)
    
End Sub

Private Sub DataGrid1_BeforeInsert(Cancel As Integer)
    '*=Implement here
    RaiseEvent BeforeRowInsert(DataGrid1.Row, Cancel)
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    '*=Implement here
    RaiseEvent BeforeRowUpdate(DataGrid1.Row, Cancel)
    
End Sub

Private Sub DataGrid1_ColEdit(ByVal ColIndex As Integer)
    RaiseEvent ColumnEdit(ColIndex, DataGrid1.Row)
End Sub

Private Sub DataGrid1_GotFocus()
    On Error GoTo Failed
    Dim bComboVisible As Boolean

    'If m_bLeavingGrid = False Then
        ActivateFirstColumn
        
        ' Combo?
        bComboVisible = HandleCombo()
        
        If ValidateTextBox() = True Then
            HandleEditBox False, False
        End If
    
    'End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub TextGotFocus()
    ' If we're on a tabbed dialog, due to a VB bug, we need to ensure we can only tab around the tab
    ' we're on. This is done by actually disabling every tab stop!
    EnableTabs False
    tmrCheckKeyPress.Enabled = True
End Sub

Private Sub combo1_GotFocus()
    ' If we're on a tabbed dialog, due to a VB bug, we need to ensure we can only tab around the tab
    ' we're on. This is done by actually disabling every tab stop!
    
    EnableTabs False
End Sub

Private Sub TextKeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    Dim bSuccess As Boolean
    Dim Cancel As Integer
    Dim bColEnabled As Boolean
    
    Cancel = 0
    
    If KeyAscii = vbKeyTab Then
        bColEnabled = IsColumnEnabled(m_nTextCol + 1)
        If m_nTextCol + 1 > m_nEnabledCols Then
            
                m_bLeavingGrid = True
                UpdateTextBox
                
                ShowTextBox False, False
                
                ' Enable tabstops again.
                EnableTabs
                HandleKeyMove "{TAB}", KeyAscii, False
            
        ElseIf m_nTextCol + 1 <= m_nEnabledCols And Not bColEnabled Then
                HandleKeyMove "{TAB}{TAB}", KeyAscii
        Else
                HandleKeyMove "{TAB}", KeyAscii
        End If
    Else
        ' Need to let things outside this control know that we are changing a cell.
        RaiseEvent BeforeColEdit(m_nTextCol, KeyAscii, Cancel)
        If Cancel <> 0 Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub TextLostFocus()
    On Error GoTo Failed
    
    EnableTabs True
    'ShowTextBox False, False
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub ValidateKeyPress(KeyAscii As Integer)
    Dim bRet As Boolean
    Dim sChar As String
    On Error GoTo Failed
    
    bRet = True
    sChar = Chr(KeyAscii)

    If IsControl(KeyAscii) = False Then
        bRet = ValidateChar(KeyAscii)
    End If

    If bRet = False Then
        KeyAscii = 0
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub HandleKeyMove(sCmd As String, Optional KeyCode As Integer, Optional bGridFocus As Boolean = True)
    On Error GoTo Failed
    If bGridFocus = True Then
        DataGrid1.SetFocus
    End If
    SendKeys sCmd
    
    If Not IsMissing(KeyCode) Then
        KeyCode = 0
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Sub DataGrid1_Click()
    On Error GoTo Failed
    Dim bEnable As Boolean

    If DataGrid1.SelBookmarks.Count = 1 Then
        combo1.Visible = False
        ShowTextBox False, True
        bEnable = True
    End If

    HandleDeleteState bEnable
    DataGrid1.Refresh
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo Failed
    
    'combo1.Visible = False
    'ShowTextBox False, True
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Function ValidateChar(nkeycode As Integer) As Boolean
    Dim vCol As Variant
    Dim col As Column
    Dim bValid As Boolean
    Dim nType As DataTypeEnum
    Dim sValue As String
    Dim sChar As String
    Dim nPrecision As Long
    
    vCol = DataGrid1.col
    
    On Error Resume Next
    
    bValid = True
    sChar = Chr(nkeycode)
    vCol = DataGrid1.col
    
    If Not IsEmpty(vCol) Then
        Set col = DataGrid1.Columns(CLng(vCol))
        
        'sValue = col.Value
        sValue = txtGridField.Text
        nType = m_rsData(vCol).Type
        
        Select Case nType

        ' Whole number only
        ' DJP SQL Server port, add adInteger
        Case adNumeric, adInteger
            'AA - bValid = IsNumeric(sChar)
            bValid = CheckMinus(nkeycode)
            If Not bValid Then
                bValid = IsNumeric(sChar)
            End If
        ' DJP SQL Server port, add adDouble
        Case adVarNumeric, adDouble
           ' bValid = (IsNumeric(sChar)) Or _
                    ((Chr(nKeyCode) = "." And InStr(1, sValue, ".", vbTextCompare) = 0) Or Len(txtGridField.SelText) > 0)
            'SA SYS3542 Check for minus figures first.
            bValid = CheckMinus(nkeycode)
            If Not bValid Then
                bValid = ValidateDecimalPoint(nkeycode)
            End If
        ' String
        ' DJP SQL Server port, add adVarWChar
        Case adChar, adVarChar, adVarWChar
                Dim dValue As Double
                Dim dMaxValue As Double
                
                ' Check the size isn't too big for the field
                dValue = Len(CStr(txtGridField.Text)) + 1
                dMaxValue = m_rsData(vCol).DefinedSize

                If dValue > dMaxValue Then
                    nkeycode = 0
                End If
                If nkeycode > 0 And IsNumeric(sChar) Or sChar = "." Then
                    bValid = ValidateDecimalPoint(nkeycode)
                End If
        End Select
    End If
    
    ValidateChar = bValid
End Function

Public Sub HandlePoint(sVal As Variant, nAfterPoint As Integer)
    On Error GoTo Failed
    
    Dim nPos As Integer
    Dim nLen As Integer
    Dim nLenAfterPoint As Integer
    Dim sAfterPoint As String
    Dim nDiff As Integer

    nLen = Len(sVal)
    nPos = InStr(1, sVal, ".")

    If nPos > 0 Then
        sAfterPoint = Right(sVal, nLen - nPos)
        nLenAfterPoint = Len(sAfterPoint)

        If nLenAfterPoint > 0 Then
            nDiff = nLenAfterPoint - nAfterPoint
        Else
            nDiff = 1
        End If
        sVal = Left(sVal, nLen - nDiff)

    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Function ValidateStringTextBox(dValue As Variant) As Boolean
    On Error GoTo Failed
    Dim dMaxValue As Double
    Dim dMinValue As Double
    Dim nColWidth As Integer
    Dim bCheck As Boolean
    Dim vCol As Variant
    Dim col As Column
    Dim nType As DataTypeEnum
    Dim bValid As Boolean
    Dim nPrecision As Integer
    Dim Field As FieldData
    Dim bCheckMin As Boolean
    
    bCheckMin = False
    bValid = True
    bCheck = False
    
    vCol = m_nTextCol

    If Not IsEmpty(vCol) And Not m_bControlLostFocus Then
        Set col = DataGrid1.Columns(vCol)
        
        Field = m_colColStatus.Item(GetKey(col.ColIndex))
'AA - There is an error here. m_rsdata is not set, and some dodgy msg is displayed
        nType = m_rsData(vCol).Type

        Select Case nType

        ' DJP SQL Server port, add adDouble and ad Integer
        Case adNumeric, adVarNumeric, adDouble, adInteger
            nPrecision = m_rsData(vCol).Precision
            bCheck = True
            
            Select Case nPrecision
                Case TYPE_SHORT
                    If Len(Field.sMaxValue) > 0 Then
                        dMaxValue = CDbl(Field.sMaxValue)
                    Else
                        dMaxValue = MAX_SHORT
                    End If
                                        
                    If Len(Field.sMinValue) > 0 Then
                        dMinValue = CDbl(Field.sMinValue)
                        bCheckMin = True
                    End If
                
                Case TYPE_LONG
                    dMaxValue = MAX_LONG
                
                Case TYPE_DOUBLE, TYPE_DOUBLE_SQL_SERVER
                    dMaxValue = MAX_DOUBLE
                    dMinValue = MIN_DOUBLE
                    HandlePoint dValue, 2

                Case 0
                    bCheck = False
                Case TYPE_BOOLEAN
                    dMaxValue = 1
                
                Case Else
                    MsgBox "Unknown Numeric Type: " & nPrecision
                    bValid = False
            End Select
        Case Else
            bValid = True
        End Select
        
        If bCheck = True Then
            If Val(dValue) > dMaxValue Then
                MsgBox "Field cannot be more than " & dMaxValue, vbCritical
                'EditActive CInt(vCol)
                bValid = False
            End If
' AA - 26/02/01 Removed the checkmin check, so that all values are checked against the minvalue set on the field type
'            If bCheckMin = True Then
                If Val(dValue) < dMinValue Then
                    MsgBox "Field cannot be less than " & dMinValue, vbCritical
                    EditActive CInt(vCol)
                    bValid = False
                End If
 '           End If
        End If
    End If

    ValidateStringTextBox = bValid
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Private Function IsControl(nkeycode As Integer) As Boolean
' AA Keycode constants do not match KeyAscii character codes! - eg VbKeyEnd = #....
    If nkeycode = vbKeyBack Then
        IsControl = True
    Else
        IsControl = False
    End If

End Function

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DataGrid1.SelBookmarks.Count = 1 Then
        DataGrid1.SelBookmarks.Remove (0)
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Failed
    Dim bComboVisible As Boolean

    If m_bValidating = False And m_bUpdating = False Then
        bComboVisible = HandleCombo()
        
        If Not bComboVisible Then
            If ValidateTextBox() = True Then
                HandleEditBox False, True
            End If
        End If
        RaiseEvent RowColChange(LastRow, LastCol)
    
    End If

    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub UpdateTextBox()
    Dim nNewRow As Integer
    Dim nNewCol As Integer
    Dim Field As FieldData
    Dim bSuccess As Boolean
    On Error GoTo Failed
    bSuccess = True
    
    
    If txtGridField.Visible = True Then
        If m_nTextRow <> -1 And m_nTextCol <> -1 Then

            ' First, validate the text box

            bSuccess = txtGridField.ValidateData()
            If bSuccess = True Then

                Dim vText As Variant
                nNewRow = m_rsData.AbsolutePosition - 1
                nNewCol = DataGrid1.col
                
                vText = txtGridField.Text
                
                If IstxtDateField() Then
                    If Len(txtGridField.ClipText) > 0 Then
                        bSuccess = IsDateValid(CStr(vText))
                    Else
                        vText = txtGridField.ClipText
                    End If
                End If
                
                If bSuccess Then
                    m_bUpdating = True
    
                    ' Update the current cell with the value obtained from the textedit
                    If m_nTextRow >= 0 And m_nTextCol >= 0 Then
                        SetAtRowCol m_nTextRow, m_nTextCol, vText
    
                        Field = m_colColStatus(GetKey(m_nTextCol))
                        If Not IsNull(vText) And Len(vText) > 0 Then
                            If Field.bDateField = True And Len(Field.sOtherField) > 0 Then
                                CheckDayOfWeek
                            End If
                        End If
                    End If
                    
                    ' Move to the new position
                    If nNewRow >= 0 Then
                        m_rsData.Move nNewRow, 1
                    End If
                
                    If nNewCol >= 0 Then
                        DataGrid1.col = nNewCol
                    End If
                End If
                m_bUpdating = False
            Else

                Timer1.Tag = TIMER_TEXT_FOCUS
                Timer1.Enabled = True
            End If
        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Sub UpdateComboValue()
    Dim col As Column
    Dim nOldRow As Integer
    Dim Field As FieldData
    Dim sOtherField As String
    Dim sUpdateVal As String
    Dim rs As adodb.Recordset
    Dim colIDS As New Collection
    
    On Error GoTo Failed
    
    If combo1.Visible = True Then
        If m_nComboCol <> -1 And m_nComboRow <> -1 Then
            
            ' Is this column a combo?
            nOldRow = DataGrid1.Row
            DataGrid1.Row = m_nComboRow
            
            Set col = DataGrid1.Columns(m_nComboCol)
            col.Value = combo1.SelText
            
            If nOldRow >= 0 Then
                DataGrid1.Row = nOldRow
            End If
            
            ' Get the information associated with the current field
            Field = m_colColStatus.Item(GetKey(col.ColIndex))
            
            ' This is the field that will contain the numberic value to be written
            sOtherField = Field.sOtherField
            
            If Len(sOtherField) > 0 Then
                Set rs = DataGrid1.DataSource
                
                If combo1.ListIndex >= 0 Then
                    
                    ' This is the list of combo value ID's associated with the combo we are using.
                    Set colIDS = Field.colComboIDS
                    Dim bUseValues As Boolean
                    
                    bUseValues = False
                    
                    If Not colIDS Is Nothing Then
                        If colIDS.Count > 0 Then
                            sUpdateVal = colIDS(combo1.ListIndex + 1)
                        Else
                            bUseValues = True
                        End If
                    Else
                        bUseValues = True
                    End If
                
                    If bUseValues = True Then
                        sUpdateVal = combo1.SelText
                    End If
                    
                    ' Update the numeric field with the numeric value we have obtained
                    rs.Fields(sOtherField).Value = sUpdateVal
                End If
            End If

        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Function HandleCombo(Optional bScroll As Boolean) As Boolean
    Dim nLeft As Integer
    Dim Field As FieldData
    Dim colValues As New Collection
    Dim colIDS As New Collection
    Dim col As Column
    Dim bComboVisible As Boolean
    
    On Error Resume Next

    ' Does the current column have a combo associated with it?
    
    ' First, hide the combo if the current (old) cell is showing it.
    combo1.Visible = False
    bComboVisible = False
    'txtGridField.Visible = False
    
    If DataGrid1.Row <> -1 And DataGrid1.col <> -1 Then
       
        ' Get the field information associated with the new cell
        Field = m_colColStatus(GetKey(DataGrid1.col))
              
        If Err.Number = 0 Then
            Set colValues = Field.colComboValues
            Set colIDS = Field.colComboIDS
            
            If colValues.Count > 0 Then
                ' Get the current column
                Set col = DataGrid1.Columns(DataGrid1.col)
                
                m_nComboCol = DataGrid1.col
                m_nComboRow = DataGrid1.Row
                
                ' Get it's coordinates so we can place the combobox
                If Not DataGrid1.CurrentCellVisible And Not bScroll Or (col.Left + 400) > DataGrid1.Width And Not bScroll Then
                    'Scroll to position
                    DataGrid1.Scroll 1, 1
                    DataGrid1.Scroll m_nComboCol - DataGrid1.LeftCol, 0
                    'DataGrid1.Scroll -222, -2223
                End If
                nLeft = col.Left
                nLeft = nLeft + DataGrid1.Left
                
                ' Place the combo
'                If ((nLeft + col.Width) > DataGrid1.Width) Then
'                    combo1.Width = col.Width - ((nLeft + col.Width) - DataGrid1.Width)
'                Else
'                    combo1.Width = col.Width
'                End If
                
                combo1.Width = col.Width
                combo1.Left = nLeft
                combo1.Top = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
                
                ' Hide the comob if a) we can't see the current cell (if we've scrolled it out of
                ' position for example) or b) we are scrolling - easier to hide it than make it scroll too.
                'If DataGrid1.CurrentCellVisible = False Or bScroll = True Then
                'bScroll = True
                If bScroll = True Then
                    combo1.Visible = False
                Else
                    combo1.Visible = True
                    If bScroll = False Then
                        
                        ' Set the combo values up
                        combo1.SetListTextFromCollection colValues, colIDS
                                            
                        ' If the current cell has a value, then set the combo selectionto be the same
                        If Len(col.Value) > 0 Then
                            combo1.Text = col.Value
                        End If
                    End If
                End If
            End If
        Else
            Err.Clear
        End If
    End If

    ' Setting the combo focus only works when we've exited this function
    If combo1.Visible = True Then
        txtGridField.Visible = False
        
        bComboVisible = True
        Timer1.Interval = 1
        Timer1.Enabled = True
        Timer1.Tag = TIMER_COMBO_FOCUS
    End If

    HandleCombo = bComboVisible

End Function

Private Sub HandleEditBox(bScroll As Boolean, bClear As Boolean)
    Dim nLeft As Integer
    Dim bRet As Boolean
    Dim Field As FieldData
    Dim col As Column
    Dim sOtherField As String
    Dim bUseTextBox As Boolean
    On Error Resume Next
    
    bUseTextBox = False

    ' If the old cell has a textbox associated with it so save that value first.
    If txtGridField.Visible = True Then
        UpdateTextBox
    End If
    
    ' Hide the textbox as a default
    ShowTextBox False, bClear
    
    If DataGrid1.Row <> -1 And DataGrid1.col <> -1 Then
       Field = m_colColStatus(GetKey(DataGrid1.col))

        If Err.Number = 0 Then
            ' Check to see if we are a date field or not. If we are we want to use the MSGEditBox
            ' with it's texttype set to CONTROL_DATE, if it's not, use the normal edit box.
            
            If Field.bDateField = True Then
                Set txtGridField = txtDateField
                bUseTextBox = True
            Else
                If combo1.Visible = False And DataGrid1.Columns(DataGrid1.col).Locked = False Then
                    Set txtGridField = txtStringField
                    bUseTextBox = True
                End If
            End If

            If bUseTextBox Then
                Set col = DataGrid1.Columns(DataGrid1.col)
                
                m_nTextCol = DataGrid1.col
                m_nTextRow = m_rsData.AbsolutePosition - 1
                
                ' Get it's coordinates so we can place the combobox
                If Not DataGrid1.CurrentCellVisible And Not bScroll Or (col.Left + 200) > DataGrid1.Width And Not bScroll Then
                    'Scroll to position
                    DataGrid1.Scroll 1, 1
                    DataGrid1.Scroll m_nTextCol - DataGrid1.LeftCol, 0 ' m_nTextRow
                End If
                
                nLeft = col.Left
                nLeft = nLeft + DataGrid1.Left
    
                ' Position the textbox
                If ((nLeft + col.Width) > DataGrid1.Width) Then
                    'The column width is greater than the datagrid so, make the width of the textbox the len of remaining space between the textbox and the datagrid width
                    txtGridField.Width = (col.Width - 10) - ((nLeft + col.Width) - DataGrid1.Width)
                Else
                    txtGridField.Width = (col.Width - 10)
                End If
                
                txtGridField.Left = nLeft + 10
                txtGridField.Top = (DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top) + 20
'Shorten the height of the textbox so that the gridlines are not covered
                txtGridField.Height = DataGrid1.RowHeight - 10
                ' Show it
                ShowTextBox True, bClear
                
                If bScroll = False Then
                    'If Len(col.Value) > 0 Then
                        If IsNull(col.Value) Then
                            txtGridField.Text = ""
                        Else
                            txtGridField.Text = col.Value
                        End If
                        
                    'End If
                Else
                    If DataGrid1.CurrentCellVisible = False Then
                        
                        ShowTextBox False, bClear
                    Else
                        
                        ShowTextBox True, bClear
                    End If
                End If
            'Endif
            End If
        Else
            Err.Clear
        End If
    End If

    If txtGridField.Visible = True Then
        Timer1.Interval = 1
        Timer1.Enabled = True
        Timer1.Tag = TIMER_TEXT_FOCUS
    End If

End Sub

Public Sub SetColumns(colFields As Collection, Optional sRegSetting As String, Optional sTitle As String)
    Dim Field As FieldData
    Dim nCount As Integer
    Dim sField As String
    Dim nThisField As Integer
    Dim col As Column
    
    Dim bFound As Boolean
    Dim sMatchField As String
    Dim columnField As FieldData
    Dim colNew As New Collection
    On Error GoTo Failed
    
    m_bValidating = False
    m_bUpdating = False
    
    Set m_colFields = colFields
    m_sRegSetting = sRegSetting
    
    On Error Resume Next

    If Len(sTitle) > 0 Then
        App.Title = sTitle
    End If

    ' Hide the combo and date controls
    combo1.Visible = False
    txtGridField.Visible = False
    
    nCount = colFields.Count
    
    If nCount > 0 Then
        For nThisField = 1 To nCount
            bFound = False
            Field = colFields(nThisField)
            sField = Field.sField
        
            If Len(sField) > 0 Then
                ' Does the field exist?
                For Each col In DataGrid1.Columns
                    
                    columnField = m_colColStatus.Item(GetKey(col.ColIndex))
                    
                    If Err.Number <> 0 Then
                        Err.Clear
                        sMatchField = col.Caption
                    Else
                        If Len(columnField.sField) > 0 Then
                            sMatchField = columnField.sField
                        Else
                            sMatchField = col.Caption
                        End If
                    
                    End If
                    
                    If UCase(sMatchField) = UCase(sField) Then
                        bFound = True
                        
                        If Len(Field.sTitle) > 0 Then
                            col.Caption = Field.sTitle
                        End If
                        
                        SetColumnState col.ColIndex, Field
                        col.AllowSizing = True
                        
                        Exit For
                        
                    End If
                Next col
            End If
            
            If bFound = False Then
                m_clsErrorHandling.RaiseError errGeneralError, "Field " + sField + " does not exist"
            End If
        Next
        
        LoadColumnDetails sRegSetting
        ProcessColumns
        SetInitialSize
        
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Sub ProcessColumns()

    Dim col As Column
    m_nEnabledCols = 0

    For Each col In DataGrid1.Columns
        SetLockedState col
        
        If col.Locked = False And col.Visible = True Then
            m_nEnabledCols = col.ColIndex
        End If
    Next


End Sub

Private Sub SetInitialSize()
    Dim colLast As Column
    Dim dDiff As Double
    Dim colThisCol As Column
    Dim dSize As Double
    Dim dColWidth As Double
    Dim dGridWidth As Double
    Dim dIncrease As Double
    Dim dColPercent As Double
    
    On Error GoTo Failed
    Set colLast = GetLastColumn()

    If Not colLast Is Nothing Then
        If colLast.Left > 0 Then
            dDiff = DataGrid1.Width - (colLast.Left + colLast.Width)
    
            If dDiff > 0 Then
                dGridWidth = DataGrid1.Width
                
                For Each colThisCol In DataGrid1.Columns
                    If colThisCol.Visible = True Then
                        dColWidth = colThisCol.Width
                        dColPercent = dColWidth * 100 / dGridWidth
                        dIncrease = (dDiff / 100) * (dColPercent - 3.1)
                        dSize = dColWidth + dIncrease
                        
                        colThisCol.Width = dSize
                    End If
                Next
            End If
        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Function GetLastColumn() As Column
    On Error GoTo Failed
    Dim colLast As Column
    Dim colHeader As Column
    
    For Each colHeader In DataGrid1.Columns
        If colHeader.Visible = True Then
            Set colLast = colHeader
        End If
    Next

    If Not colLast Is Nothing Then
        Set GetLastColumn = colLast
    Else
        m_clsErrorHandling.RaiseError errGeneralError, "GetLastColumn - no columns visible"
    End If
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Sub LoadColumnDetails(sRegLocation As String)
    On Error GoTo Failed
    Dim colHeader As Column
    Dim sWidth As String
    
    m_sRegLocation = sRegLocation
    
    If Len(sRegLocation) > 0 Then
        ' Loop around all columns reading the width - if set, set the column width
        For Each colHeader In DataGrid1.Columns

            If colHeader.Visible = True Then
                sWidth = QueryValue(HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, sRegLocation & COL_WIDTH & colHeader.Caption)
                
                If IsNumeric(sWidth) Then
                    colHeader.Width = CLng(sWidth)
                End If
            End If
        Next
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Public Sub SaveColumnDetails()
    On Error GoTo Failed
    Dim colHeader As Column
    Dim nWidth As Long
    
    If Len(m_sRegLocation) > 0 Then
        For Each colHeader In DataGrid1.Columns
            If colHeader.Visible = True Then
                nWidth = colHeader.Width
            
                SetKeyValue HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, m_sRegLocation & COL_WIDTH & colHeader.Caption, CStr(nWidth), REG_SZ
            End If
        Next
    End If

    m_sRegLocation = ""
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Public Sub SetColumnState(nCol As Integer, Field As FieldData)
    On Error Resume Next
    Dim columnField As FieldData
    
    If Len(Field.sMinValue) = 0 Then
        Field.sMinValue = "0"
    End If
    
    If Not DataGrid1.DataSource Is Nothing Then
        DataGrid1.Columns(nCol).Visible = Field.bVisible
        DataGrid1.RowHeight = combo1.Height
        
        columnField = m_colColStatus.Item(GetKey(nCol))

        If Err.Number <> 0 Then
            Err.Clear
            columnField = Field
            m_colColStatus.Add columnField, GetKey(nCol)
        Else
            columnField = Field
            m_colColStatus.Remove (GetKey(nCol))
            m_colColStatus.Add columnField, GetKey(nCol)
        End If
    Else
        MsgBox "MSGDataGrid: Not connected to Datasource"
    End If
End Sub

Private Function GetKey(nNum As Integer) As String
    GetKey = Trim("ColKey" & nNum)
End Function

Public Sub EditActive(Optional nCol As Integer = -1, Optional bActive As Boolean = True)
    On Error GoTo Failed
    Dim nFirst As Integer
    
    ' Need to set the leftmost column to the first visible column
    nFirst = FindFirstVisibleColumn()
    
    If nCol = -1 Then
        nCol = nFirst
    End If
    
    DataGrid1.col = nCol
    DataGrid1.EditActive = bActive
    DataGrid1.LeftCol = nFirst
    
    ' Use a timer to set the focus (so the grid can do it's work first)
    Timer1.Interval = 1
    Timer1.Enabled = True
    Timer1.Tag = TIMER_GRID_FOCUS

    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Sub AddRow(Optional bSetFocus As Boolean = True)
    AddNewRecord bSetFocus
    ' SQL Server port, when adding a new record move to the first field we can see
    EditActive
End Sub

Private Sub EnableTabsParent(ctrlParent As Object, Optional bEnable As Boolean = True)
    On Error Resume Next
    Dim i As Integer
    Dim ctrl As Object

    For i = 0 To ctrlParent.Controls.Count - 1   ' Use the Controls collection
        Set ctrl = ctrlParent.Controls(i)
        Dim a As Frame
        
        If m_bSetTabs = False Then
            If Len(ctrl.Tag) = 0 Then
                ctrl.Tag = CStr(ctrl.TabStop)
            End If
        End If
        
        If bEnable = True Then
            If Len(CStr(ctrl.Tag)) > 0 Then
                If ctrl.Visible = True Then
                    If CBool(ctrl.Tag) = True Then
                        ctrl.TabStop = CBool(ctrl.Tag)
                    Else
                        ctrl.TabStop = False
                    End If
                End If
            End If
        Else
            ctrl.TabStop = False
        End If
    Next
End Sub

Private Sub EnableTabs(Optional bEnable As Boolean = True)
    On Error Resume Next
    Dim i As Integer
    Dim ctrl As Object
    Dim ctrlParent As Object
    
    For i = 0 To UserControl.Controls.Count - 1   ' Use the Controls collection
        Set ctrl = UserControl.Controls(i).TabStop
        If m_bSetTabs = False Then
            If Len(ctrl.Tag) = 0 Then
                ctrl.Tag = CStr(UserControl.Controls(i).TabStop)
            End If
        End If

        If bEnable = True Then
            If Len(CStr(ctrl.Tag)) > 0 Then
                If ctrl.Visible = True Then
                    If CBool(ctrl.Tag) = True Then
                        ctrl.TabStop = CBool(ctrl.Tag)
                    Else
                        ctrl.TabStop = False
                    End If
                End If
            End If
        Else
            ctrl.TabStop = False
        End If
    Next
    
    Dim bIsTabbedDialog As Boolean
    bIsTabbedDialog = False
    
    If bEnable Then
        ' Find the parent form
        Dim bDone As Boolean
        
        bDone = False
        
'        Set ctrlParent = Extender.Container

        While bDone = False And Not ctrlParent Is Nothing
            If TypeOf ctrlParent Is SSTab Then
                bIsTabbedDialog = True
                bDone = True
            Else
                Set ctrlParent = ctrlParent.Container
                If Err.Number <> 0 Then
                    Err.Clear
                    If TypeOf ctrlParent Is Form Then
                        bDone = True
                    End If
                End If
            End If
        Wend
    
        If bIsTabbedDialog Then
            SetTabstops UserControl.Parent
        End If
    End If
    
    If Not bIsTabbedDialog Or bEnable = False Then
        EnableTabsParent UserControl.Parent, bEnable
    End If
    
    If m_bSetTabs = False Then
        m_bSetTabs = True
    End If
End Sub

Public Sub SetGridFocus()
    On Error GoTo Failed
    
    If DataGrid1.Enabled = True Then
        DataGrid1.SetFocus
    End If
    
    If Rows() = 0 Then
        If cmdAdd.Enabled = True Then
            cmdAdd.SetFocus
        End If
    End If
    
    Exit Sub
Failed:
    Err.Clear
End Sub

Public Function DoControlValidation() As Boolean
    Dim bRet As Boolean
    bRet = True
    
    If txtGridField.Visible = True Then
        bRet = txtGridField.ValidateData()
    
        If bRet = True Then
            UpdateTextBox
        Else
            Timer1.Tag = TIMER_TEXT_FOCUS
            Timer1.Enabled = True
        End If
    End If

    DoControlValidation = bRet
End Function

Private Sub ShowTextBox(bVisible As Boolean, bClear As Boolean)
    Dim sMask As String
    Dim col As Column
    On Error GoTo Failed
    
    If bClear = True Then
        If IstxtDateField = True Then
            ' If we have a mask, clear it.
            sMask = txtGridField.mask
            txtGridField.mask = ""
            txtGridField.Text = ""
            txtGridField.mask = sMask
        Else
            txtGridField.Text = ""
        End If
    End If

    
    txtGridField.Visible = bVisible
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Function ValidateTextBox() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
    bRet = True
    
   If txtGridField.Visible = True Then
        
        bRet = txtGridField.ValidateData()
        
        If bRet = False Then
            If m_nTextRow >= 0 Then
                DataGrid1.Row = m_nTextRow
            End If
            
            If m_nTextCol >= 0 Then
                DataGrid1.col = m_nTextCol
            End If
            SetTextFocus
        End If
    End If

    ValidateTextBox = bRet
    Exit Function
Failed:
    ValidateTextBox = False
End Function

Private Sub DataGrid1_RowResize(Cancel As Integer)
    combo1.Visible = False
    ShowTextBox False, True
End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)
    
    ' This will be called when we are adding a record, but we only want to do this if we aren't adding a record
    If Not m_bAddingRecord Then
        HandleCombo True
        
        ' If we are scrolling, we will be hiding the textbox so save it's contents first.
        UpdateTextBox
        ShowTextBox False, False
    End If
End Sub

Private Sub tmrCheckKeyPress_Timer()
    'While the focus is set on a textbox, this function is called to check if
    'The alt key has been pressed. If it has then we need to update the textbox
    ' DJP SQL Server port, add error handling
    On Error GoTo Failed
    If GetAsyncKeyState(ALT_KEY_CODE) < 0 And txtGridField.Visible = True Then
        UpdateTextBox
        tmrCheckKeyPress.Enabled = False
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub txtDateField_GotFocus()
    On Error GoTo Failed
    
    ' Call generic  text got focus function
    TextGotFocus
Failed:
End Sub

Private Sub txtDateField_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed

    ' Call generic  text keypress
    TextKeyPress KeyAscii
Failed:
End Sub

Private Function DoesControlCauseValidation() As Boolean
    On Error GoTo Failed
    Dim bDone As Boolean
    Dim bCausesValidation As Boolean
    Dim ctrlParent As Object
    
    bDone = False
    bCausesValidation = False
    DoesControlCauseValidation = False
    
   ' If Not UserControl.Parent Is Nothing Then
        On Error Resume Next
        'Resume next if Client site not available
        Set ctrlParent = UserControl.Parent
        If Err.Number <> 0 Then
            Err.Clear
            DoesControlCauseValidation = True
            Exit Function
        End If
   ' End If
    
    While bDone = False And Not ctrlParent Is Nothing
        
        If TypeOf ctrlParent Is Form Then
            bCausesValidation = ctrlParent.ActiveControl.CausesValidation
            bDone = True
        Else
            Set ctrlParent = ctrlParent.Container
        End If
    Wend
    
    DoesControlCauseValidation = bCausesValidation
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Private Sub combo1_LostFocus()
    On Error GoTo Failed
    
    If m_nComboCol = m_nEnabledCols Then
        m_bLeavingGrid = False
    End If
    
    EnableTabs
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub txtDateField_LostFocus()
    On Error GoTo Failed
    Dim bCausesValidation
    
    bCausesValidation = DoesControlCauseValidation()
                
    If bCausesValidation Then
        ' DJP, the textbox does it's own message display, so pass False in
        ' so we just check to see if the data is valid or not, but do not
        ' display another error (get multiple error messages otherwise)
        If txtDateField.ValidateData(False) = True Then
            
            UpdateTextBox
            TextLostFocus
        Else
            Timer1.Tag = TIMER_TEXT_FOCUS
            Timer1.Enabled = True
        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub txtStringField_GotFocus()
    On Error GoTo Failed
    TextGotFocus
Failed:
End Sub

Private Sub txtStringField_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    TextKeyPress KeyAscii
    
    If KeyAscii <> 0 Then
        ValidateKeyPress KeyAscii
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Private Sub txtStringField_LostFocus()
    On Error GoTo Failed
    Dim bCausesValidation
    Dim vNewValue As Variant
    bCausesValidation = DoesControlCauseValidation()
                
    If bCausesValidation Then
        Dim bValid As Boolean
        vNewValue = txtStringField.Text
        bValid = ValidateStringTextBox(vNewValue)
        
        If bValid Then
            txtStringField.Text = vNewValue
            UpdateTextBox
            TextLostFocus
        Else
            Timer1.Tag = TIMER_TEXT_FOCUS
            Timer1.Enabled = True
        End If
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError

End Sub

Private Sub UserControl_GotFocus()
    tmrCheckKeyPress.Enabled = True
End Sub

Private Sub UserControl_Hide()
    On Error GoTo Failed
    HandleTermination
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub UserControl_Initialize()
    Set m_colColStatus = New Collection
    Set m_clsErrorHandling = New ErrorHandling
    Set m_clsGenAssist = New GeneralAssist
    
    m_nComboCol = -1
    m_nComboRow = -1
    m_bValidating = False
    m_bLeavingGrid = False
    m_bSetTabs = False
    m_bUpdating = False
    Set txtGridField = txtStringField
    m_nEnabledCols = 0
    'CL AQR SYS4787
    DataGrid1.Splits(0).SizeMode = dbgExact
    'END AQR SYS4787

    ' Uncomment this for debug file output
    ' Open "C:\Out.txt" For Output As #1
    ' Close
End Sub

Private Sub PrintDebug(str As String)
    Open "C:\Out.txt" For Append As #1
    Print #1, Now & ": " & str
    Close
End Sub

Public Function Rows() As Long
    If Not m_rsData Is Nothing Then
        Rows = m_rsData.RecordCount
    Else
        Rows = 0
    End If
End Function

Public Sub SetSelection(nStart As Integer, nEnd As Integer)
    DataGrid1.SelStartCol = nStart
    DataGrid1.SelEndCol = nEnd
End Sub

Private Sub UserControl_LostFocus()
    'When the control has lost focus, we can disable the timer!
    m_bControlLostFocus = True
    tmrCheckKeyPress.Enabled = False
End Sub


Private Sub UserControl_Resize()
    DataGrid1.Height = UserControl.Height - DataGrid1.Top ' - 850
    '*=MC - Size aligned properly
    DataGrid1.Width = UserControl.Width - 200
End Sub


Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = New_Enabled
    DataGrid1.Enabled = New_Enabled
    cmdAdd.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    
    'cmdDelete.Enabled = New_Enabled
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
     
End Sub

Public Property Get DataSource() As adodb.Recordset
    Set DataSource = DataGrid1.DataSource
    'SetFieldNames
End Property

Public Property Set DataSource(ByVal New_DataSource As adodb.Recordset)
        SaveColumnDetails
        Set DataGrid1.DataSource = New_DataSource
        Set m_rsData = New_DataSource
        PropertyChanged "DataSource"
End Property

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ContainerHwnd = m_def_ContainerHwnd
    m_Text = m_def_Text
    m_Row = m_def_Row
    m_Col = m_def_Col
    m_SelStartCol = m_def_SelStartCol
    m_SelEndCol = m_def_SelEndCol
    m_AllowAdd = m_def_AllowAdd
    m_AllowDelete = m_def_AllowDelete
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    DataGrid1.ColumnHeaders = PropBag.ReadProperty("ColumnHeaders", True)
    m_ContainerHwnd = PropBag.ReadProperty("ContainerHwnd", m_def_ContainerHwnd)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_Row = PropBag.ReadProperty("Row", m_def_Row)
    m_Col = PropBag.ReadProperty("Col", m_def_Col)
    m_SelStartCol = PropBag.ReadProperty("SelStartCol", m_def_SelStartCol)
    m_SelEndCol = PropBag.ReadProperty("SelEndCol", m_def_SelEndCol)
    m_AllowAdd = PropBag.ReadProperty("AllowAdd", m_def_AllowAdd)
    m_AllowDelete = PropBag.ReadProperty("AllowDelete", m_def_AllowDelete)
    cmdAdd.Visible = m_AllowAdd
    cmdDelete.Visible = m_AllowDelete
       
End Sub

Private Function HandleTermination()
    SaveColumnDetails
    tmrCheckKeyPress.Enabled = False
    Set DataGrid1.DataSource = Nothing
    Set m_clsGenAssist = Nothing
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("ColumnHeaders", DataGrid1.ColumnHeaders, True)
    Call PropBag.WriteProperty("ContainerHwnd", m_ContainerHwnd, m_def_ContainerHwnd)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Row", m_Row, m_def_Row)
    Call PropBag.WriteProperty("Col", m_Col, m_def_Col)
    Call PropBag.WriteProperty("SelStartCol", m_SelStartCol, m_def_SelStartCol)
    Call PropBag.WriteProperty("SelEndCol", m_SelEndCol, m_def_SelEndCol)
    Call PropBag.WriteProperty("AllowAdd", m_AllowAdd, m_def_AllowAdd)
    Call PropBag.WriteProperty("AllowDelete", m_AllowDelete, m_def_AllowDelete)
End Sub

Public Property Get Columns() As Columns
    Set Columns = DataGrid1.Columns
End Property

Public Property Get ColumnHeaders() As Boolean
    ColumnHeaders = DataGrid1.ColumnHeaders
End Property

Public Property Let ColumnHeaders(ByVal New_ColumnHeaders As Boolean)
    DataGrid1.ColumnHeaders() = New_ColumnHeaders
    PropertyChanged "ColumnHeaders"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = m_ContainerHwnd
End Property

Public Property Let ContainerHwnd(ByVal New_ContainerHwnd As Long)
    m_ContainerHwnd = New_ContainerHwnd
    PropertyChanged "ContainerHwnd"
End Property


Public Property Get Text() As String
    '*=Design time error fixed.
    If Not DataGrid1.DataSource Is Nothing Then
        Text = DataGrid1.Text
    End If
End Property


Public Property Let Text(ByVal New_Text As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Text = New_Text
    DataGrid1.Text = New_Text
    PropertyChanged "Text"
End Property

Public Property Get Row() As Integer
    Row = DataGrid1.Row
End Property

Public Property Let Row(ByVal New_Row As Integer)
    On Error GoTo Failed
    If Ambient.UserMode = False Then Err.Raise 387
    m_Row = New_Row
    DataGrid1.Row = New_Row
    PropertyChanged "Row"
    Exit Property
Failed:
    'MsgBox "Row: Error is - " + Err.Description
End Property

Public Property Get col() As Integer
    col = DataGrid1.col
End Property

Public Property Let col(ByVal New_Col As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Col = New_Col
    DataGrid1.col = New_Col
    PropertyChanged "Col"
End Property

Public Property Get SelStartCol() As Integer
    SelStartCol = DataGrid1.SelStartCol
End Property

Public Property Let SelStartCol(ByVal New_SelStartCol As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    
    m_SelStartCol = New_SelStartCol
    DataGrid1.SelStartCol = New_SelStartCol
    PropertyChanged "SelStartCol"
End Property

Public Property Get SelEndCol() As Integer
    SelEndCol = DataGrid1.SelEndCol
End Property

Public Property Let SelEndCol(ByVal New_SelEndCol As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    
    m_SelEndCol = New_SelEndCol
    DataGrid1.SelEndCol = New_SelEndCol
    PropertyChanged "SelEndCol"
End Property

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    RaiseEvent BeforeColEdit(ColIndex, KeyAscii, Cancel)
End Sub


Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    Dim col As Column
    Dim nType As DataTypeEnum
    Dim nPrecision As Long
    Dim dValue As Variant
    Dim dOldValue As Variant
    
    ' Truncate and numbers after decimal point to 2 places.
    nType = m_rsData(ColIndex).Type

    Select Case nType

    ' Whole number only
    ' DJP SQL Server port, add adDouble and adInteger
    Case adNumeric, adVarNumeric, adDouble, adInteger
        nPrecision = m_rsData(ColIndex).Precision
        
        Select Case nPrecision
            Case TYPE_DOUBLE
                dValue = m_rsData(ColIndex).Value
                dOldValue = dValue
                HandlePoint dValue, 2
                
                If dValue <> dOldValue Then
                    DataGrid1.Columns(ColIndex).Value = dValue
                End If
        End Select
    End Select
    '*=Raise Event for source application if it require any action to take
    RaiseEvent AfterColumnUpdated(ColIndex, DataGrid1.Row)
    
End Sub


Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    RaiseEvent BeforeDelete(Cancel)
End Sub

Public Property Get SelBookmarks() As SelBookmarks
    Set SelBookmarks = DataGrid1.SelBookmarks
End Property

Private Sub SetTextFocus()
    Timer1.Tag = TIMER_TEXT_FOCUS
    Timer1.Enabled = True
End Sub

Private Sub DataGrid1_LostFocus()
    m_bLeavingGrid = False
End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
End Sub

Private Sub Timer1_Timer()
    Select Case Timer1.Tag
    Case TIMER_COMBO_FOCUS
        If combo1.Visible = True Then
            combo1.SetFocus
        End If
    Case TIMER_GRID_FOCUS
        DataGrid1.EditActive = True
        DataGrid1.SetFocus

    Case TIMER_TEXT_FOCUS
        If txtGridField.Visible = True Then
            txtGridField.SetFocus
        End If
    Case TIMER_DELETE_STATE
        HandleDeleteState False

    End Select
    
    Timer1.Enabled = False
End Sub

Private Function IstxtDateField() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = False
    
    If TypeName(txtGridField) = TypeName(txtDateField) Then
        ' It's an msgeditbox
        If txtGridField.TextType = CONTROL_DATE Then
            bRet = True
        End If
        
    End If
    
    IstxtDateField = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Private Sub CheckDayOfWeek()

    Dim col As Column
    Dim nUpdateRow As Integer
    Dim nUpdateCol As Integer
    Dim nOldRow As Integer
    Dim Field As FieldData
    Dim sOtherFieldText As String
    Dim sOtherFieldNum As String
    Dim sUpdateTextVal As String
    Dim sUpdateNumericVal As String
    Dim rs As adodb.Recordset
    Dim colIDS As New Collection
    Dim nWeekDay As Integer
    Dim nOtherCol As Integer
    Dim bRet As Boolean
    Dim sOtherField As String
    
    On Error GoTo Failed

    nWeekDay = Weekday(CDate(txtGridField.Text), vbSunday)

    ' This is the list of combo value ID's associated with the combo we are using.
    Set col = DataGrid1.Columns(m_nTextCol)

    Field = m_colColStatus.Item(GetKey(col.ColIndex))
    Set colIDS = Field.colComboIDS

    sOtherField = Field.sOtherField
    
    If Len(sOtherField) > 0 Then
        sOtherFieldNum = Field.sOtherField
        nOtherCol = GetOtherFieldColumnNo(sOtherFieldNum)
    
        Field = m_colColStatus.Item(GetKey(nOtherCol))
    
       sOtherFieldText = Field.sOtherField
       
       bRet = SetLockedState(col)
        If bRet = True Then
        
            Set rs = DataGrid1.DataSource
    
            If nWeekDay >= 0 Then
                
                  Dim bUseValues As Boolean
    
                  bUseValues = False
    
                  If Not colIDS Is Nothing Then
                      If Field.colComboIDS.Count > 0 Then
                          sUpdateNumericVal = Field.colComboIDS(nWeekDay)
                          sUpdateTextVal = Field.colComboValues(nWeekDay)
                      Else
                          bUseValues = True
                      End If
                  Else
                      bUseValues = True
                  End If
    
                  ' Update the numeric field with the mumeric value we have obtained
                  rs.Fields(sOtherFieldText).Value = sUpdateTextVal
                ' Update the text field with the text value we have obtained
                  rs.Fields(sOtherFieldNum).Value = sUpdateNumericVal
            End If
        End If
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
    
End Sub

Private Function GetOtherFieldColumnNo(sOtherField As String) As Integer

    Dim Field As FieldData
    Dim nCnt As Integer
    
    'On Error Resume Next
    
    GetOtherFieldColumnNo = 0
    
    For nCnt = 1 To m_colColStatus.Count
        Field = m_colColStatus.Item(GetKey(nCnt))
        If UCase(Field.sField) = UCase(sOtherField) Then
            GetOtherFieldColumnNo = nCnt
            Exit Function
        End If
    Next nCnt
    
    'Otherfield not found
    
    MsgBox "Otherfield not found", vbCritical
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Private Function SetLockedState(col As Column) As Boolean
    Dim sOtherFieldNum As String
    Dim nOtherCol As Integer
    Dim colOtherCol As Column
    Dim sOtherFieldText As String
    Dim bRet As Boolean
    Dim Field As FieldData
    bRet = False
    On Error Resume Next
    Field = m_colColStatus.Item(GetKey(col.ColIndex))
    
    sOtherFieldNum = Field.sOtherField
    If Len(Field.sOtherField) > 0 Then
        nOtherCol = GetOtherFieldColumnNo(sOtherFieldNum)
        Field = m_colColStatus.Item(GetKey(nOtherCol))
        If Err.Number = 0 Then
            sOtherFieldText = Field.sOtherField
    
            Set colOtherCol = DataGrid1.Columns(nOtherCol)
     
            nOtherCol = GetOtherFieldColumnNo(sOtherFieldText)
            Set colOtherCol = DataGrid1.Columns(nOtherCol)
        Else
            Err.Clear
        End If
    End If
    
    If Len(sOtherFieldNum) > 0 And Len(sOtherFieldText) > 0 Then
        colOtherCol.Locked = True
        bRet = True
    End If
    
    SetLockedState = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number
End Function

Public Function GetFieldData(ColIndex As Integer) As FieldData
    Dim Field As FieldData
    On Error GoTo Failed
    
    If ColIndex > -1 Then
        Field = m_colColStatus.Item(GetKey(ColIndex))
    End If
    
    GetFieldData = Field
    
    Exit Function

Failed:
    m_clsErrorHandling.RaiseError Err.Number

End Function

Public Function IsColumnEnabled(nColIndex As Integer) As Boolean

    Dim bRet As Boolean
    Dim Field As FieldData
    On Error GoTo Failed
    
    If nColIndex >= m_nEnabledCols Then
        IsColumnEnabled = False
        Exit Function
    End If
    
    bRet = True
    Field = GetFieldData(nColIndex)
    
   If Field.bDisabled = True Then

        bRet = False
    End If
    
    IsColumnEnabled = bRet
    
    Exit Function
    
Failed:
    m_clsErrorHandling.RaiseError Err.Number

End Function

Public Sub HideGridColumn(nColIndx As Integer, Optional bVisible As Boolean = False)

    On Error GoTo Failed
    
    DataGrid1.Columns(nColIndx).Visible = bVisible
    HandleEditBox False, False
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number

End Sub

Private Function CheckMinus(nkeycode As Integer) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = False
    
    
    If txtGridField.SelStart = 0 And Chr(nkeycode) = "-" And InStr(1, txtGridField.Text, "-") = 0 Or txtGridField.SelLength = Len(txtGridField.Text) And Chr(nkeycode) = "-" Or InStr(1, txtGridField.SelText, "-") > 0 Then
        If Chr(nkeycode) <> "." Then
            bRet = True
        End If
    ElseIf txtGridField.SelStart = 0 And InStr(1, txtGridField.Text, "-") > 0 Then
        bRet = True
    End If
    
    CheckMinus = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number
End Function

Private Function ValidateDecimalPoint(nkeycode As Integer) As Boolean
    
    On Error GoTo Failed
    Dim sTmp As String
    Dim stmp1 As String
    Dim nStart As Long
    Dim nLen As Long
    Dim bRet As Boolean
    Dim sLeft As String
    Dim sRight As String
    Dim nSel As Integer
    bRet = True
    
    If Not IsControl(nkeycode) Then
        nSel = txtGridField.SelLength
        If nSel = 0 And InStr(1, txtGridField.SelText, ".", vbTextCompare) = 0 Then
            nLen = Len(txtGridField.Text)
            nStart = txtGridField.SelStart
            sTmp = Left(txtGridField.Text, nStart)
            stmp1 = sTmp + Chr(nkeycode)
            sTmp = stmp1 + Right(txtGridField.Text, nLen - nStart)
            
            bRet = ((Chr(nkeycode) = "." And InStr(1, txtGridField.Text, ".", vbTextCompare) = 0) Or Len(txtGridField.SelText) > 0) Or IsNumeric(Chr(nkeycode))
            
            If bRet Then
                bRet = CheckDecimalPlaces(sTmp, 2)
            End If
        Else
            
            sLeft = Left(txtGridField.Text, txtGridField.SelStart)
            sRight = Right(txtGridField.Text, Len(txtGridField.Text) - txtGridField.SelStart)
            
            If InStr(1, sLeft, ".") > 0 And Chr(nkeycode) = "." Or InStr(1, sRight, ".") > 0 And Chr(nkeycode) = "." Then
                If InStr(1, txtGridField.SelText, ".") = 0 And txtGridField.SelLength < Len(txtGridField.Text) Or InStr(1, txtGridField.SelText, "-") > 0 And txtGridField.SelLength < Len(txtGridField.Text) Then
                    bRet = False
                End If
            ElseIf Not IsNumeric(Chr(nkeycode)) And Chr(nkeycode) <> "." Then
                bRet = False
            End If
        End If
    End If
    
    If Not bRet Then
        nkeycode = 0
    End If
    
    ValidateDecimalPoint = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number
End Function

Public Function CheckDecimalPlaces(sText As String, nAfterPointValid As Long)
    Dim bRet As Boolean
    Dim nPos As Long
    Dim nSel As Long
    Dim nAfterPoint As Long
    
    On Error GoTo Failed
    bRet = True
    ' Need to work out how many places are after the point. > 1 and return false, else true
    
    ' After Point
    nPos = InStr(1, sText, ".", vbTextCompare)
    
    If (nPos >= 1) Then
        nAfterPoint = Len(sText) - nPos
        
        If (nAfterPoint > nAfterPointValid Or nAfterPointValid = 0) Then
            bRet = False
        End If
    End If

    ' Before Point
    
    If (bRet = True) Then
        If (nPos = 0) Then
          nPos = Len(sText)
         Else
            nPos = nPos - 1
        End If
          
        If (nPos > 0) Then
            Dim sTmp As String
            Dim nLen As Long
            
            sTmp = Left(sText, nPos - 1)
            
        End If
    End If

    CheckDecimalPlaces = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number
End Function

Public Property Get AllowAdd() As Boolean
    AllowAdd = m_AllowAdd
End Property

Public Property Let AllowAdd(ByVal New_AllowAdd As Boolean)
    m_AllowAdd = New_AllowAdd
    PropertyChanged "AllowAdd"
    cmdAdd.Visible = New_AllowAdd
End Property

Public Property Get AllowDelete() As Boolean
    AllowDelete = m_AllowDelete
End Property

Public Property Let AllowDelete(ByVal New_AllowDelete As Boolean)
    m_AllowDelete = New_AllowDelete
    DataGrid1.AllowDelete = New_AllowDelete
    PropertyChanged "AllowDelete"
    cmdDelete.Visible = New_AllowDelete
End Property

'MAR312 GHun
Public Function GetBoundRecordSet() As adodb.Recordset
    Set GetBoundRecordSet = m_rsData
End Function
'Mar312 End
