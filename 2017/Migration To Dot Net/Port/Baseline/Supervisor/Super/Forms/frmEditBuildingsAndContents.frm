VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBuildingsAndContents 
   Caption         =   "Add/Edit Buildings & Contents Product"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   Icon            =   "frmEditBuildingsAndContents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   315
      Left            =   2160
      TabIndex        =   14
      Top             =   2220
      Width           =   1515
      Begin VB.OptionButton optContentsCover 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   555
      End
      Begin VB.OptionButton optContentsCover 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   1740
      Width           =   1515
      Begin VB.OptionButton optBuildingsCover 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   555
      End
      Begin VB.OptionButton optBuildingsCover 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   165
      TabIndex        =   10
      Top             =   3675
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1515
      TabIndex        =   11
      Top             =   3675
      Width           =   1230
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   2865
      TabIndex        =   12
      Top             =   3675
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtBuildingContents 
      Height          =   315
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   2700
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TextType        =   7
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
   End
   Begin MSGOCX.MSGEditBox txtBuildingContents 
      Height          =   315
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   3120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TextType        =   7
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
   End
   Begin MSGOCX.MSGEditBox txtBuildingContents 
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TextType        =   6
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
      MaxLength       =   5
   End
   Begin MSGOCX.MSGTextMulti txtProductName 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   540
      Width           =   2295
      _ExtentX        =   4048
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
      Mandatory       =   -1  'True
      MaxLength       =   255
   End
   Begin MSGOCX.MSGEditBox txtBuildingContents 
      Height          =   315
      HelpContextID   =   2
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
   End
   Begin MSGOCX.MSGEditBox txtBuildingContents 
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   1380
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
   End
   Begin VB.Label lblBuildingContents 
      Caption         =   "Valuables Item Limit"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   22
      Top             =   3180
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Contents Cover?"
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Buildings Cover?"
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Product Number"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label lblMaintIncentives 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   180
      TabIndex        =   17
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblBuildingContents 
      Caption         =   "Valuables Limit"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditBuildingsAndContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditBuildingsAndContents
' Description   :   Form which allows the user to edit and add Buildings and
'                   Contents details
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Changed call
'                       to NextNumber as it has been moved from FormProcessing to DataAccess.
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Local text indexes
Private Const PRODUCT_NUMBER = 0
Private Const START_DATE = 2
Private Const END_DATE = 3
Private Const VALUABLES_LIMIT = 4
Private Const VALUABLES_ITEM_LIMIT = 5

' Local Radio button indexes
Private Const RADIO_YES = 0
Private Const RADIO_NO = 1

' Private data
Private m_bIsEdit As Boolean
Private m_clsBuildingContents As BuildingAndContentsTable
Private m_vNextNumber As Variant
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection


Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user presses the Another button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If

    If bRet = True Then
        txtProductName.SetFocus
        CreateNewRecord
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user clicks the cancel button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


Private Sub ValidateScreenData()
    On Error GoTo Failed
    
    If optBuildingsCover(RADIO_YES).Value = False And Me.optContentsCover(RADIO_YES).Value = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "At least one of Buildings or Contents cover must be selected"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok or
'                   presses Another
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        ValidateScreenData
        SaveScreenData
        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user pressed the OK button. Performs necessary
'                   validation and saves any data that needs to be saved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sProductNumber As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    sProductNumber = txtBuildingContents(PRODUCT_NUMBER).Text
    colMatchValues.Add sProductNumber
    
    Set clsTableAccess = m_clsBuildingContents
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsBuildingContents
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Initialisation to this screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    SetReturnCode MSGFailure
    
    Set m_clsBuildingContents = New BuildingAndContentsTable
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding new B&C details
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    m_vNextNumber = Null
    
    ' Generates a runtime error caught in form load if this fails
    CreateNewRecord
    
    ' Default to yes
    optContentsCover.Item(RADIO_YES).Value = True
    optBuildingsCover.Item(RADIO_YES).Value = True
    EnableContents
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing exisiting B&C details
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGUID As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection
    
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsBuildingContents
    clsTableAccess.SetKeyMatchValues m_colKeys
     
    Set rs = clsTableAccess.GetTableData()
    cmdAnother.Enabled = False

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Building and Content record"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   EnableContents
' Description   :   Sets the state of the fields in this screen depending
'                   on the bEnable parameter passed in. True, and the fields
'                   are enabled, false they are disabled.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableContents(Optional bEnable As Boolean = True)
    txtBuildingContents(VALUABLES_LIMIT).Text = ""
    txtBuildingContents(VALUABLES_ITEM_LIMIT).Text = ""
    txtBuildingContents(VALUABLES_LIMIT).Enabled = bEnable
    txtBuildingContents(VALUABLES_ITEM_LIMIT).Enabled = bEnable
    lblBuildingContents(VALUABLES_LIMIT).Enabled = bEnable
    lblBuildingContents(VALUABLES_ITEM_LIMIT).Enabled = bEnable

    txtBuildingContents(VALUABLES_LIMIT).Mandatory = bEnable
    txtBuildingContents(VALUABLES_ITEM_LIMIT).Mandatory = bEnable
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   optContentsCover_Click
' Description   :   Sets the state of the fields in this screen depending
'                   on the state of the Contents Cover radio button. If it
'                   is no, then the B&C fields are disabled, yes they are
'                   enabled.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optContentsCover_Click(Index As Integer)
    If optContentsCover.Item(RADIO_YES).Value = True Then
        EnableContents
    Else
        EnableContents False
    End If
End Sub
Private Sub txtBuildingContents_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtBuildingContents(Index).ValidateData()
End Sub
Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    'Product Number
    m_vNextNumber = m_clsBuildingContents.GetBCProductNumber()
        
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        txtBuildingContents(PRODUCT_NUMBER).Text = m_vNextNumber
    
        'Product Name
        txtProductName.Text = m_clsBuildingContents.GetProductName()
        
        ' Start Date
        vTmp = m_clsBuildingContents.GetStartDate()
        g_clsFormProcessing.HandleDate txtBuildingContents(START_DATE), vTmp, SET_CONTROL_VALUE
    
        ' End Date
        vTmp = m_clsBuildingContents.GetEndDate()
        g_clsFormProcessing.HandleDate txtBuildingContents(END_DATE), vTmp, SET_CONTROL_VALUE
    
        ' Buildings cover
        vTmp = m_clsBuildingContents.GetIncludesBuildingsCover()
        g_clsFormProcessing.HandleRadioButtons optBuildingsCover(RADIO_YES), optBuildingsCover(RADIO_NO), vTmp, SET_CONTROL_VALUE
    
        ' Contents cover
        vTmp = m_clsBuildingContents.GetIncludesContentsCover()
        g_clsFormProcessing.HandleRadioButtons optContentsCover(RADIO_YES), optContentsCover(RADIO_NO), vTmp, SET_CONTROL_VALUE
    
        EnableContents CBool(vTmp)
    
        'Valuables limit
        txtBuildingContents(VALUABLES_LIMIT).Text = m_clsBuildingContents.GetValuablesLimit()
        
        'Valuables item limit
        txtBuildingContents(VALUABLES_ITEM_LIMIT).Text = m_clsBuildingContents.GetValuablesItemLimit()
    
    End If
    
    PopulateScreenFields = True
    Exit Function
Failed:
    PopulateScreenFields = False
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CreateNewRecord
' Description   :   Creates a new record in the Buildings and contents
'                   table and generates a new "next number" for the key.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateNewRecord()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim sNextField As String
    
    Set clsTableAccess = m_clsBuildingContents
    
    g_clsFormProcessing.CreateNewRecord m_clsBuildingContents
    
    sNextField = m_clsBuildingContents.GetNextNumberField()
    
    If Len(sNextField) > 0 Then
        m_vNextNumber = g_clsDataAccess.NextNumber(clsTableAccess.GetTable(), sNextField)
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "CreateNewRecord: Next number field is empty"
    End If
    ' Set the Product number
    txtBuildingContents(PRODUCT_NUMBER).Text = CStr(m_vNextNumber)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    Dim vNextNumber As Variant
    
    Set clsTableAccess = m_clsBuildingContents
    
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        'Product Number
        m_clsBuildingContents.SetBCProductNumber m_vNextNumber
        
        'Product Name
        m_clsBuildingContents.SetProductName txtProductName.Text
        
        ' Start Date
        g_clsFormProcessing.HandleDate txtBuildingContents(START_DATE), vTmp, GET_CONTROL_VALUE
        m_clsBuildingContents.SetStartDate vTmp
        
        ' End Date
        g_clsFormProcessing.HandleDate txtBuildingContents(END_DATE), vTmp, GET_CONTROL_VALUE
         m_clsBuildingContents.SetEndDate vTmp
    
        'Valuables limit
        m_clsBuildingContents.SetValuablesLimit txtBuildingContents(VALUABLES_LIMIT).Text
        
        'Valuables item limit
        m_clsBuildingContents.SetValuablesItemLimit txtBuildingContents(VALUABLES_ITEM_LIMIT).Text

        ' Buildings cover
        g_clsFormProcessing.HandleRadioButtons optBuildingsCover(RADIO_YES), optBuildingsCover(RADIO_NO), vTmp, GET_CONTROL_VALUE
        m_clsBuildingContents.SetIncludesBuildingsCover CStr(vTmp)
        
        ' Contents cover
        g_clsFormProcessing.HandleRadioButtons optContentsCover(RADIO_YES), optContentsCover(RADIO_NO), vTmp, GET_CONTROL_VALUE
        m_clsBuildingContents.SetIncludesContentsCover CStr(vTmp)
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "BuildingsAndContents - Next Number is empty"
    End If
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
