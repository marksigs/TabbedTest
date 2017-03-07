VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIncMultiples 
   Caption         =   "Add / Edit Income Multiple Set"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      TextType        =   7
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
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
      MaxLength       =   50
   End
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      TextType        =   2
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
   End
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      TextType        =   2
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
   End
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      TextType        =   2
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
   End
   Begin MSGOCX.MSGEditBox txtIncMult 
      Height          =   315
      Index           =   5
      Left            =   1680
      TabIndex        =   5
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      TextType        =   2
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
   End
   Begin VB.Label Label1 
      Caption         =   "Low Earner Mult"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2715
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "High Earner Mult"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   2235
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Joint Inc Mult"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1755
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Single Inc Mult"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Multiplier Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   280
      Width           =   1260
   End
End
Attribute VB_Name = "frmEditIncMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditIncMultiples
' Description   :   Form which allows the user to edit an income multiple set.
'
' Change history
' Prog      Date        Description
'
' AW        14/05/02    BM088   Initial version
' AW        27/05/02    Added SaveChangeRequest()
' AW        28/05/02    Added DoesRecordExist()
' AW        28/09/02    BMIDS00258  Mandatory fields
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_clsIncomeMultipleSet As New IncomeMultipleSetTable
Private m_colKeys As Collection

'Control Indexes.
Private Const INCOME_MULT_CODE      As Long = 0
Private Const INCOME_MULT_DESC      As Long = 1
Private Const INCOME_MULT_SINGLE    As Long = 2
Private Const INCOME_MULT_JOINT     As Long = 3
Private Const INCOME_MULT_HIGH      As Long = 4
Private Const INCOME_MULT_LOW       As Long = 5

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form, the return status will indicate failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate data and attempt to save. If successful, close the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    'Call into the shared routine to save the record(s).
    bRet = DoOKProcessing()

    'If successful, hide this form and return processing to the opener.
    If bRet = True Then
        Hide
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Populate combos, prepare the underlying objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Default return status is failure.
    SetReturnCode MSGFailure
    
    'Create an underlying table object.
    Set m_clsIncomeMultipleSet = New IncomeMultipleSetTable
    
    'Setup the controls according to the forms state.
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Load the desired record into the table object and populate the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    'Cast a generic table interface onto the underlying table object.
    Set clsTableAccess = m_clsIncomeMultipleSet
    
    'Set the keys collection to the one passed in from MainSuper.
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    'Load the record and obtain a reference to the ADO recordset.
    Set rs = clsTableAccess.GetTableData()
    
    'Another is disabled (we're editing).
    cmdAnother.Enabled = False
    
    'Ensure a record was loaded.
    ValidateRecordset rs, "Income Multiples"
    
    'Double ensure? - the populate the screen with the data.
    If rs.RecordCount = 1 Then
        'We can't edit the primary key if we're in an edit state.
        Me.txtIncMult(INCOME_MULT_CODE).Enabled = False
        
        PopulateScreenFields
    Else
        g_clsErrorHandling.RaiseError errRecordNotFound
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populate each screen control from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PopulateScreenFields() As Boolean
    
    Dim vTmp As Variant
    
    On Error GoTo Failed
        
    txtIncMult(INCOME_MULT_CODE).Text = m_clsIncomeMultipleSet.GetIncMultCode()
        
    txtIncMult(INCOME_MULT_DESC).Text = m_clsIncomeMultipleSet.GetIncMultDesc()
        
    txtIncMult(INCOME_MULT_SINGLE).Text = m_clsIncomeMultipleSet.GetSingleIncMult
        
    txtIncMult(INCOME_MULT_JOINT).Text = m_clsIncomeMultipleSet.GetJointIncMult()
        
    txtIncMult(INCOME_MULT_HIGH).Text = m_clsIncomeMultipleSet.GetHighEarnerMult()
        
    txtIncMult(INCOME_MULT_LOW).Text = m_clsIncomeMultipleSet.GetLowEarnerMult()
    
    PopulateScreenFields = True
    
    Exit Function
    
Failed:
    PopulateScreenFields = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Copy all the control values into the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SaveScreenData() As Boolean
    
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    Dim strCode As String
    Dim bRecExists As Boolean
        
    On Error GoTo Failed
    
    bRecExists = False
    
    If m_bIsEdit = False Then
        bRecExists = DoesRecordExist
    End If
    
    If bRecExists = False Then
        Set clsTableAccess = m_clsIncomeMultipleSet
        strCode = txtIncMult(INCOME_MULT_CODE).Text
        
        If Not IsNull(strCode) And Not IsEmpty(strCode) Then
    
            m_clsIncomeMultipleSet.SetIncMultCode strCode
            
            m_clsIncomeMultipleSet.SetIncMultDesc txtIncMult(INCOME_MULT_DESC).Text
            
            m_clsIncomeMultipleSet.SetSingleIncMult txtIncMult(INCOME_MULT_SINGLE).Text
            
            m_clsIncomeMultipleSet.SetJointIncMult txtIncMult(INCOME_MULT_JOINT).Text
            
            m_clsIncomeMultipleSet.SetHighEarnerMult txtIncMult(INCOME_MULT_HIGH).Text
            
            m_clsIncomeMultipleSet.SetLowEarnerMult txtIncMult(INCOME_MULT_LOW).Text
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Multiplier Code is empty"
        End If
        
        clsTableAccess.Update
        
    End If
    
    SaveScreenData = Not bRecExists
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Save the current record and re-initialise the form to add another record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Call into the shared routine to save the record(s).
    bRet = DoOKProcessing()
        
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        'If the record was saved, clear the screen fields ready for another record.
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If
    
    If bRet = True Then
        'Ensure the focus is back in the first field.
        Me.txtIncMult(INCOME_MULT_CODE).SetFocus
        
        'Re-initialise the form state, this will generate a life rate number.
        SetAddState
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateNewRecord
' Description   : Creates a blank record in the underlying table object and generates a new
'                 Life Rate number (the primary key).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateNewRecord()
    
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed

    'Cast a generic interface onto the underlying table object.
    Set clsTableAccess = m_clsIncomeMultipleSet
    
    'Create a new record.
    g_clsFormProcessing.CreateNewRecord m_clsIncomeMultipleSet
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Validate and save the current record then route into the promotion code. This
'                 routine returns true if successful. Note: This routine is present on forms
'                 with an 'Another' button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    'Ensure all mandatory fields have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    'If they have then proceed.
    If bRet = True Then
        'Save the data.
        bRet = SaveScreenData
        
        If bRet = True Then
            'Ensure the record is flagged for promotion.
            SaveChangeRequest
            
            'Set the return status to success.
            SetReturnCode
        End If
    End If

    'Return True (success) or False (failure) to the caller.
    DoOKProcessing = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Ensures the collection of primary key values is set against the underlying
'                 table object and then passes it into the promotion code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
        
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    'Create a new collection.
    Set colMatchValues = New Collection
            
    'Add the Life Cover Rate number (primary key) into it.
    colMatchValues.Add txtIncMult(INCOME_MULT_CODE).Text
    
    'Cast a generic interface onto the table object.
    Set clsTableAccess = m_clsIncomeMultipleSet
    
    'Set the keys collection.
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    'Call into the promotion code to ensure the update marked for promotion.
    g_clsHandleUpdates.SaveChangeRequest m_clsIncomeMultipleSet
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Creates a blank record, assigns a new Life rate number and populates the
'                 screen control with its value.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    
    On Error GoTo Failed
    
    'Ensure the code is cleared.
    'txtIncMult(INCOME_MULT_CODE).SetFocus
    txtIncMult(INCOME_MULT_CODE).Text = ""
    
    CreateNewRecord
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return status of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Returns the sucess/fail status to the form's caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Checks to see if a record exists already with the same
'                   keys. Returns true if it exists, false if not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim sDesc As String
    
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    Dim clsMigFeeSet As MPMigRateSetTable
    
    sSet = txtIncMult(INCOME_MULT_CODE).Text
    
    If Len(sSet) > 0 Then
    
        Set clsTableAccess = m_clsIncomeMultipleSet
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        bRet = clsTableAccess.DoesRecordExist(col)
        
        If bRet = True Then
            MsgBox "Set already exist - please enter a unique combination", vbCritical
            txtIncMult(INCOME_MULT_CODE).SetFocus
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Set must be valid", vbCritical
    End If
    
    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
