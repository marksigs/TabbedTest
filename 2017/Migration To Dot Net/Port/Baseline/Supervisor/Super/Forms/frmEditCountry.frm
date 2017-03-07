VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditCountry 
   Caption         =   "Add/Edit Country"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   Icon            =   "frmEditCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGTextMulti txtCountryName 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   4500
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1755
      TabIndex        =   4
      Top             =   4500
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   435
      TabIndex        =   3
      Top             =   4500
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtCountryNumber 
      Height          =   315
      Left            =   1860
      TabIndex        =   0
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      TextType        =   6
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
   Begin MSGOCX.MSGDataGrid dgBankHoliday 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      Enabled         =   -1  'True
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
   Begin VB.Label lblEditGlobal 
      Caption         =   "Country Number"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label lblEditGlobal 
      Caption         =   "Country Name"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   540
      Width           =   1035
   End
End
Attribute VB_Name = "frmEditCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditCountry
' Description   :   Form which allows the user to edit and add Country details
' Change history
' Prog      Date        Description
'           06/12/00    Added Bankholiday Datagrid
' STB       06/12/01    SYS1942 - Another button commits current transaction.
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sCountryNumber As String
Private m_bIsEdit As Boolean
Private m_clsCountry As CountryTable
Private m_clsBankHoliday As BankHolidayTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_sCountryName As String
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsCountry = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    bRet = DoOkProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If
    
    If bRet = True Then
        SetupDataGrid
    End If
    
    If bRet = True Then
        txtCountryNumber.SetFocus
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function DoOkProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
        End If
    End If

    DoOkProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOkProcessing = False
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOkProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim sCountryNumber As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sDesc = Me.txtCountryName.Text
    sCountryNumber = txtCountryNumber.Text
    colMatchValues.Add sCountryNumber
    Set clsTableAccess = m_clsCountry
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsCountry, sDesc
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    ' Only one of Amount, Percentage, Boolean and string can be entered.
    Dim nCount As Integer
    Dim bSuccess As Boolean
    Dim sCountryName As String
    On Error GoTo Failed
    sCountryName = txtCountryName.Text
    
    bSuccess = Not g_clsFormProcessing.CheckForDuplicates(m_clsBankHoliday)

    If bSuccess = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Bank Holiday Date must be unique"
    End If

    If bSuccess Then
        bSuccess = dgBankHoliday.ValidateRows()
    End If
    
    If m_bIsEdit = False And bSuccess Then
        bSuccess = Not DoesRecordExist()
    Else
        If m_sCountryName <> sCountryName Then
            bSuccess = Not CheckForDuplicateValues
        End If
    End If

    ValidateScreenData = bSuccess
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sCountry As String
    Dim clsCountry As New CountryTable
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection

    On Error GoTo Failed
    sCountry = txtCountryNumber.Text

    If Len(sCountry) > 0 Then
        Set clsTableAccess = clsCountry
        colValues.Add sCountry
        
        
        bRet = clsTableAccess.DoesRecordExist(colValues)
        
        If bRet = True Then
            g_clsErrorHandling.RaiseError errGeneralError, "Country Number " & sCountry & " already exists"
            txtCountryNumber.SetFocus
        End If
        
        If bRet = False Then
            bRet = CheckForDuplicateValues
        End If
    End If
    
    DoesRecordExist = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SetGridFields()
    Dim bRet As Boolean
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsCombo As New ComboUtils
    Dim sCombo As String
    Dim clsTableAccess As TableAccess
    Dim rsBankHoliday As ADODB.Recordset
    
    bRet = True

    Set clsTableAccess = m_clsBankHoliday
    Set rsBankHoliday = clsTableAccess.GetRecordSet()
    Set dgBankHoliday.DataSource = rsBankHoliday

    ' Country
    fields.sField = "CountryNumber"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sCountryNumber
    fields.sError = ""
    fields.sTitle = "Country Number"
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    fields.bDateField = False
    colFields.Add fields

    ' Bank Holiday Date
    fields.sField = "BankHolidayDate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Bank Holiday Date must be entered"
    fields.sTitle = "Bank Holiday Date"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    fields.bDateField = True
    colFields.Add fields

    ' Description
    fields.sField = "BankHolidayDescription"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Bank Holiday Description"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    fields.bDateField = False
    colFields.Add fields

    dgBankHoliday.SetColumns colFields, "BankHoliday", "Bank Holiday"
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    SetReturnCode MSGFailure
        
    Set m_clsCountry = New CountryTable
    Set m_clsBankHoliday = New BankHolidayTable
    
    If m_bIsEdit = True Then
        SetCountryEditState
        SetBankHolidayEditState
        m_sCountryNumber = txtCountryNumber.Text
    Else
        SetCountryAddState
        SetBankHolidayAddState
    End If

    SetGridFields
    
    Me.dgBankHoliday.Enabled = True
    m_sCountryName = txtCountryName.Text
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Public Sub SetBankHolidayAddState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess

    Set clsTableAccess = m_clsBankHoliday
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDataGrid
    
    cmdAnother.Enabled = True
    txtCountryNumber.Enabled = True

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Edit Country - No Countries exist"
        End If
    End If
End Sub
Public Sub SetBankHolidayEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGuid As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection

    ' First, the department table
    Set clsTableAccess = m_clsBankHoliday
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()
    

    
End Sub
Public Sub SetCountryAddState()
    
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGuid As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection

    Set clsTableAccess = m_clsCountry
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDataGrid
    
    cmdAnother.Enabled = True
    txtCountryNumber.Enabled = True
    
    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Edit Country - No Countries exist"
        End If
    End If


End Sub

Public Sub SetCountryEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGuid As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection

    ' First, the department table
    Set clsTableAccess = m_clsCountry
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()
    
    cmdAnother.Enabled = False
    txtCountryNumber.Enabled = False

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Edit Country - No Countries exist"
        End If
    End If
End Sub
Private Sub txtCountryName_Validate(Cancel As Boolean)
    Cancel = Not txtCountryName.ValidateData()
End Sub

Private Sub txtCountryNumber_Validate(Cancel As Boolean)
    Cancel = Not txtCountryNumber.ValidateData()
End Sub

Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    ' Country ID
    txtCountryNumber.Text = m_clsCountry.GetCountryNumber()
    ' Country Name
    txtCountryName.Text = m_clsCountry.GetCountryName()
    
    PopulateScreenFields = True
    Exit Function
Failed:
    PopulateScreenFields = False
End Function
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim bSuccess As Boolean
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    Dim sCountryNumber As String
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsCountry
    End If
    
    ' Country ID
    m_sCountryNumber = txtCountryNumber.Text
    m_clsCountry.SetCountryNumber m_sCountryNumber
    
    ' Country Name
    m_clsCountry.SetCountryName txtCountryName.Text
    Set clsTableAccess = m_clsCountry
    clsTableAccess.Update
        
    SaveBankHolidayDetails
    
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveBankHolidayDetails
' Description   : Saves the bankholiday records associated with the current country record
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveBankHolidayDetails()
    
    On Error GoTo Failed
    Dim colUpdateValues As Collection
    Dim clsBandedTable As BandedTable
        
     If Not m_bIsEdit Then
        Set colUpdateValues = New Collection
        
        colUpdateValues.Add m_sCountryNumber
        Set clsBandedTable = m_clsBankHoliday
        
        clsBandedTable.SetUpdateValues colUpdateValues
        clsBandedTable.SetUpdateSets
        clsBandedTable.DoUpdateSets
        
    End If
    
    TableAccess(m_clsBankHoliday).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
Private Sub SetupDataGrid()
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupDataGrid
' Description   : Populates the Bankholiday datagrid, with an empty recordset
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim clsTableAccess As TableAccess
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsBankHoliday
    
    clsTableAccess.SetKeyMatchValues m_colKeys
        
    clsTableAccess.GetTableData POPULATE_EMPTY
    
    SetGridFields
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function CheckForDuplicateValues() As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckForDuplicateValues
' Description   : Checks for duplicate values such as Country Name. The value
'                 to check for duplicates to is specified in the colvalues collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim colValues As New Collection
    Dim colFields As Collection
    Dim sCountryName As String
    Dim clsTableAccess As TableAccess
    Dim bRet As Boolean
    
    On Error GoTo Failed
    Set clsTableAccess = New CountryTable
    
    sCountryName = txtCountryName.Text
    
    Set colFields = New Collection
    Set colValues = New Collection
    
    If bRet = False Then
        colValues.Add sCountryName
        colFields.Add m_clsCountry.GetNameField
        
        bRet = clsTableAccess.DoesRecordExist(colValues, colFields)
        
        If bRet = True Then
            g_clsErrorHandling.RaiseError errGeneralError, "The country " & sCountryName & " already exists"
            txtCountryName.SetFocus
        End If
    End If
    CheckForDuplicateValues = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
