VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmProductGrouping 
   Caption         =   "Product Grouping"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmProductGrouping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame fraCurrency 
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin MSGOCX.MSGComboBox cboProductLine 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSGOCX.MSGComboBox cboProductGroup 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSGOCX.MSGComboBox cboProductBrandName 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label lblProductGroup 
         Caption         =   "Product Line"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblProductGroup 
         Caption         =   "Product Group"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblProductGroup 
         Caption         =   "Product Brand Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmProductGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmProductGrouping
' Description   :   Form which allows the user to edit and add Product Grouping
'                   details.
'
' Change history
' Prog      Date        Description
' CL        25/04/01    Initial creation of screen
' STB       30/01/02    SYS3771 Use MortgageProductTable rather than the
'                       ProductGroupingTable class. Also revamped to fit-in
'                       with Supervisor standards.
' DJP       20/02/02    DJP SYS3585 Change combo to Brand for Product Brand Name
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Control indexes.
Private Const PRODUCT_BRAND_NAME As Long = 0
Private Const PRODUCT_GROUP      As Long = 1
Private Const PRODUCT_LINE       As Long = 2

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'Underlying table object for the record.
Private m_clsMortgageProductTable As MortgageProductTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    
    On Error GoTo Failed
    
    'Hide the form and return to the caller.
    Hide
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the record, closing the form if everything
'                 is okay.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    On Error GoTo Failed
    
    'Update the underlying table from the GUI control values.
    SaveScreenData
    
    'Indicate to the caller that this form updated some data.
    SetReturnCode MSGSuccess
    
    'Hide the form and return to the caller.
    Hide

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Main startup routine for the pop-up form. The underlying
'                 table object will already be set (with the cursor at the
'                 current record) so we can setup and populate the screen
'                 controls from the underlying data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'The default form's return status is failure.
    SetReturnCode MSGFailure
          
    'Perform state-specific code.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
        
    'Populate with the existing data
    PopulateScreenFields
       
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets whether or not this form is in edit or add mode.
'                 Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return code for the form for any calling method to
'                 check. Defaults to MSG_SUCCESS.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(ByVal uReturnCode As MSGReturnCode)
    m_uReturnCode = uReturnCode
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Add specific code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Edit specific code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Update the GUI controls values from the underlying data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
 
    Dim vProductLine As Variant
    Dim vProductGroup As Variant
    Dim vProductBrandName As Variant
    
    'Populate the combo controls.
    ' DJP SYS3585 Change combo to Brand"
    g_clsFormProcessing.PopulateCombo "Brand", cboProductBrandName
    g_clsFormProcessing.PopulateCombo "ProductGroup", cboProductGroup
    g_clsFormProcessing.PopulateCombo "ProductLine", cboProductLine
    
    'Set the product brand name.
    vProductBrandName = m_clsMortgageProductTable.GetProductBrandName()
    g_clsFormProcessing.HandleComboExtra cboProductBrandName, vProductBrandName, SET_CONTROL_VALUE
    
    'Set the product group.
    vProductGroup = m_clsMortgageProductTable.GetProductGroup()
    g_clsFormProcessing.HandleComboExtra cboProductGroup, vProductGroup, SET_CONTROL_VALUE
    
    'Set the product line.
    vProductLine = m_clsMortgageProductTable.GetProductLine()
    g_clsFormProcessing.HandleComboExtra cboProductLine, vProductLine, SET_CONTROL_VALUE
            
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Update the underlying table object from the screen controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
        
    Dim vProductLine As Variant
    Dim vProductGroup As Variant
    Dim vProductBrandName As Variant
    
    'ProductBrandName.
    g_clsFormProcessing.HandleComboExtra cboProductBrandName, vProductBrandName, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetProductBrandName CStr(vProductBrandName)
    
    'ProductGroup.
    g_clsFormProcessing.HandleComboExtra cboProductGroup, vProductGroup, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetProductGroup CStr(vProductGroup)
    
    'ProductLine.
    g_clsFormProcessing.HandleComboExtra cboProductLine, vProductLine, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetProductLine CStr(vProductLine)
     
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTableClass
' Description   : Associate the underlying table from the parent form with this
'                 form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTableClass(ByRef clsMortgageProductTable As MortgageProductTable)
    Set m_clsMortgageProductTable = clsMortgageProductTable
End Sub
