VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterestRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ORGANISATION_ID_FIELD = 1
Private m_rs As ADODB.Recordset
Private m_clsTableAccess As TableAccess
Private m_clsOtherFeesTable As LenderOtherFeesTable
Private m_bIsEdit As Boolean
Private m_sLenderCode As String
Sub Class_Initialize()
    If m_clsOtherFeesTable Is Nothing Then
        Set m_clsOtherFeesTable = New LenderOtherFeesTable
    End If
End Sub
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsOtherFeesTable = clsTableAccess
End Sub
Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub
Public Function IsEdit()
    IsEdit = m_bIsEdit
End Function
Public Function Initialise(Optional bIsEdit As Boolean) As Boolean
    Dim clsTableAccess As TableAccess
    
    
    m_bIsEdit = bIsEdit
    
    Set clsTableAccess = m_clsOtherFeesTable
    Set m_rs = clsTableAccess.GetTableData()
    
    If Not m_rs Is Nothing Then

        Set frmLenderDetails.dgOtherFees.DataSource = m_rs
        frmLenderDetails.dgOtherFees.Enabled = True
        SetGridFields
    End If
    
    Initialise = True
    Exit Function
Failed:
    MsgBox "OtherFees:Initialise - Error is " + Err.Description
    Initialise = False
End Function
Private Sub SetGridFields()
    Dim fields As FieldData
    Dim colFields As New Collection
    
    ' First, OrganisationID. Not visible, but needs the lender code copied in
    fields.sField = "OrganisationID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sLenderCode
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Next, LenderOtherFeeName has to be entered
    fields.sField = "LenderOtherFeeName"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Fee Name must be entered"
    fields.sTitle = "Name"
    colFields.Add fields
    
    ' StartDate has to be entered
    fields.sField = "LenderOtherFeeStartDate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Start Date must be entered"
    fields.sTitle = "Start Date"
    colFields.Add fields
    
    ' Fee Amount
    fields.sField = "LenderOtherFeeAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Fee Amount must be entered"
    fields.sTitle = "Amount"
    colFields.Add fields
    
    frmLenderDetails.dgOtherFees.SetColumns colFields
End Sub
Public Function SetScreenFields()
    Dim bRet As Boolean

    SetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "SetScreenFields: Error is " + Err.Description
    SetScreenFields = False
End Function
Public Function DoUpdates() As Boolean
    Dim bRet As Boolean
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    bRet = True

    On Error GoTo Failed
    bRet = True
    
    Set clsTableAccess = m_clsOtherFeesTable
    
    If Not m_rs Is Nothing Then
        bRet = clsTableAccess.Update()
    End If
    
    DoUpdates = bRet
    Exit Function
Failed:
    MsgBox "DoUpdates: Error is " + Err.Description
    DoUpdates = False

End Function
Public Sub InitTab()

End Sub
Public Sub SetKeyValues(colValues As Collection)
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim sLenderCode As String
    Dim bRet As Boolean

    sLenderCode = colValues.Item(ORGANISATION_ID_FIELD)
    
    If Len(sLenderCode) > 0 Then
        Set clsTableAccess = m_clsOtherFeesTable
        m_sLenderCode = sLenderCode

        clsTableAccess.SetKeyMatchValues colValues
    Else
        MsgBox "LegalFees:LenderCode - Lender Code is empty"
    End If
    
End Sub
Public Function ValidateScreenData()
    ValidateScreenData = frmLenderDetails.dgOtherFees.ValidateRows()
End Function


