VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditProductPeriod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Product Period"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmAddEditProductPeriod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtTarget 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BackColor       =   16777215
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
   Begin MSGOCX.MSGEditBox txtTarget 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BackColor       =   16777215
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
   Begin VB.Label lblPeriod 
      Caption         =   "End Date"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblPeriod 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "frmEditProductPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditIntReportDetails
' Description   : Form which allows the user add and edit Intermediary Report details
'
' Change history
' Prog      Date        Description
' AA        26/06/01    Created
' DJP       27/06/01    SQL Server port.
' STB       22/04/02    SYS4401 Don't associate with a product if the fee type is packaging fee.
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current intermediary record.
Private m_colKeys As Collection

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'This table holds periods.
Private m_clsProcFeeTypeTable As IntermediaryProcFeeTable

'The Sequence number of the current period record. This forms part of a
'composite key field.
Private m_sSequence As String

'The procuration fee type which this period corresponds to.
Private m_uFeeType As ProcFeeTypeEnum

'These fields are only used for certain fee-types.

'The Mortgage Lender's Organisation ID.
Private m_vOrganisationID As Variant

'The ProductID which the fee applies to.
Private m_sProductID As String

'Control indexes.
Private Const START_DATE    As Long = 0
Private Const END_DATE      As Long = 1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the record, closing the form if everything is okay.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    bRet = DoOkProcessing

    If bRet Then
        Hide
    End If

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the opener. The return code will be
'                 defaulted to 'failure'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Setup the form for the current or new record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed

    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Create a new period record in the table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    g_clsFormProcessing.CreateNewRecord m_clsProcFeeTypeTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Filter the product table object so just the desired record is available.
'                 Then populate the screen controls from this record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
            
    Dim sFilter As String
        
    'Apply a filter to restrict the data to a single record.
    sFilter = "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    sFilter = sFilter & " AND Type = " & g_clsSQLAssistSP.FormatString(CStr(m_uFeeType))
    sFilter = sFilter & " AND TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(m_sSequence)
    
    'Apply the filter now.
    TableAccess(m_clsProcFeeTypeTable).ApplyFilter sFilter
    
    'Ensure only one record is available.
    If TableAccess(m_clsProcFeeTypeTable).RecordCount = 1 Then
        SetScreenData
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Procuration Fee Type not found or too many records returned"
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Set the type of procuration fee as this forms part of the composite key for
'                 the period record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeType(ByVal uFeeType As ProcFeeTypeEnum)
    m_uFeeType = uFeeType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTypeSequence
' Description   : Sets the period sequence number if we're editing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTypeSequence(ByVal sSequence As String)
    m_sSequence = sSequence
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTypeSequence
' Description   : Return the type sequence to the opener.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTypeSequence() As String
    GetTypeSequence = m_sSequence
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductID
' Description   : Sets the product id for this period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductID(ByVal sProductID As String)
    m_sProductID = sProductID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetOrganisationID
' Description   : Sets the organisation id for this period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetOrganisationId(ByVal vOrganisationID As Variant)
    m_vOrganisationID = vOrganisationID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenData
' Description   : Populate the screen elements from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetScreenData()
    
    'Start Date.
    g_clsFormProcessing.HandleDate txtTarget(START_DATE), m_clsProcFeeTypeTable.GetFeeActiveFrom, SET_CONTROL_VALUE
    
    'End Date.
    g_clsFormProcessing.HandleDate txtTarget(END_DATE), m_clsProcFeeTypeTable.GetFeeActiveTo, SET_CONTROL_VALUE
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Associate the underlying table object with this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeTypeTable(ByRef clsProcFeeTypeTable As IntermediaryProcFeeTable)
    Set m_clsProcFeeTypeTable = clsProcFeeTypeTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOkProcessing
' Description   : Validates the screen data and then saves it if valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
    
    Dim bSuccess As Boolean
    Dim bShowError As Boolean
    
    On Error GoTo Failed
    
    bShowError = True
    bSuccess = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
        
    If bSuccess Then
        ValidateDates
        SaveScreenData
    End If
    
    DoOkProcessing = bSuccess
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateDates
' Description   : Ensure that the dates specified are valid and that the end date does not
'                 preceed the start date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateDates()
    
    Dim dtEndDate As Date
    Dim dtStartDate As Date
    Dim bEndValid As Boolean
    Dim bStartValid As Boolean
        
    On Error GoTo Failed
    
    bStartValid = txtTarget(START_DATE).ValidateData(False)
    bEndValid = txtTarget(END_DATE).ValidateData(False)
    
    If bStartValid = True And bEndValid = True Then
        dtStartDate = CDate(txtTarget(START_DATE).Text)
        dtEndDate = CDate(txtTarget(END_DATE).Text)
    
        If dtEndDate < dtStartDate Then
            g_clsErrorHandling.RaiseError errGeneralError, "Start Date must be before the End Date"
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    
    Dim sGuid As String
    
    On Error GoTo Failed
    
    If Not m_bIsEdit Then
        m_sSequence = m_clsProcFeeTypeTable.GetNextSequenceNumber
    End If
        
    'Sequence.
    m_clsProcFeeTypeTable.SetTypeSequence m_sSequence
    
    'Intermediary GUID.
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_colKeys(INTERMEDIARY_KEY)))
    m_clsProcFeeTypeTable.SetIntermediaryGuid sGuid

    'Type.
    m_clsProcFeeTypeTable.SetFeeType m_uFeeType
    
    'If relevant, save the OrganisationID.
    If m_uFeeType = MortgageFee Then
        m_clsProcFeeTypeTable.SetOrganisationId m_vOrganisationID
    End If
    
    'If relevant, save the ProductID.
    Select Case m_uFeeType
        Case MortgageFee, InsuranceFee
            m_clsProcFeeTypeTable.SetProductID m_sProductID
    End Select
    
    'Active from.
    m_clsProcFeeTypeTable.SetFeeActiveFrom txtTarget(START_DATE).Text
    
    'Active to.
    m_clsProcFeeTypeTable.SetFeeActiveTo txtTarget(END_DATE).Text
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Tidy-up and release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)

    Set m_colKeys = Nothing
    Set m_clsProcFeeTypeTable = Nothing
    
End Sub
