VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Object = "{FDC96A5C-781A-11D3-AD7B-0060087A1BF0}#3.1#0"; "MSGCheckBoxControl.ocx"

Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   900
      TabIndex        =   0
      Top             =   1800
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmTest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmTest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   5
         Top             =   420
         Width           =   7215
         Begin MSGOCX.MSGDataGrid MSGDataGrid1 
            Height          =   1395
            Left            =   720
            TabIndex        =   6
            Top             =   2280
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   2461
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSGCheckBoxControl.MSGCheckBox MSGCheckBox1 
            Height          =   735
            Left            =   3780
            TabIndex        =   7
            Top             =   2040
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   1296
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSGOCX.MSGComboBox MSGComboBox1 
            Height          =   315
            Left            =   3780
            TabIndex        =   8
            Top             =   420
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSGOCX.MSGEditBox MSGEditBox4 
            Height          =   375
            Left            =   540
            TabIndex        =   9
            Top             =   420
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
         Begin MSGOCX.MSGEditBox MSGEditBox5 
            Height          =   375
            Left            =   540
            TabIndex        =   10
            Top             =   1380
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
         Begin MSGOCX.MSGComboBox MSGComboBox2 
            Height          =   315
            Left            =   3600
            TabIndex        =   11
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   3555
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   6795
         Begin MSGOCX.MSGEditBox MSGEditBox1 
            Height          =   375
            Left            =   780
            TabIndex        =   2
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
         Begin MSGOCX.MSGEditBox MSGEditBox2 
            Height          =   375
            Left            =   780
            TabIndex        =   3
            Top             =   1140
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
         Begin MSGOCX.MSGEditBox MSGEditBox3 
            Height          =   375
            Left            =   780
            TabIndex        =   4
            Top             =   1920
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
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
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SSTab1_Click(PreviousTab As Integer)
SetTabstops Me
End Sub
Public Sub SetTabstops(frmToSet As Form)
    Dim c As Control
    Dim nLeft As Long
    Dim objParent As Object
    Dim bDone As Boolean
    
    bDone = False
    
    For Each c In frmToSet.Controls
        If TypeOf c.Container Is SSTab Then
            'Not all controls have the TabStop property
            On Error Resume Next
            c.TabStop = c.Left > 0
            On Error GoTo 0
        ElseIf TypeOf c.Container Is Frame Then
            Set objParent = c.Container
            Dim bLeft As Boolean
            bDone = False
            While bDone = False
                If TypeOf objParent Is Frame Then
                    If TypeOf objParent.Container Is SSTab Then
                        bDone = True
                        On Error Resume Next
                        
                        nLeft = objParent.Left
                        
                        If Err.Number = 0 Then
                            bLeft = nLeft > 0
                            c.TabStop = bLeft
                            'MsgBox "Setting tabstop of control " & c.Name & " to " & bLeft
                        End If
                    Else
                        Set objParent = objParent.Container
                    End If
                Else
                    Set objParent = objParent.Container
                End If
            Wend
            
        End If
    Next
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
