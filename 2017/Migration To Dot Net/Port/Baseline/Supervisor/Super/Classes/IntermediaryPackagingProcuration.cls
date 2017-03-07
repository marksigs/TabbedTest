VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntPackagingProcuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : IntermediaryPackagingProcuration
' Description   : Packaging procuration tab
'
' Change history
' Prog      Date        Description
' AA        26/06/2001  Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsIntProcFee As IntProcFeeSplitTable
Private m_clsProcSplit As IntProcFeeSplitForIntTable
Private m_clsProcFee As IntermediaryProcFeeTable
Private m_nFeeValueID As Long
Private m_vIntermediaryGUID As Variant
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Public Sub Initialise(bIsEdit As Boolean, Optional bGetData As Boolean = True)
    On Error GoTo Failed
    
    Set m_clsProcFee = New IntermediaryProcFeeTable
    m_bIsEdit = bIsEdit
    m_vIntermediaryGUID = m_colKeys(1)
    
    If bGetData Then
        GetSplitDetails
    End If
    SetListViewHeaders
    PopulateScreenControls
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetKeys(col As Collection)
    Set m_colKeys = col
End Sub

Public Function GetProcTypeID() As Variant

    Dim clsComboValidation As ComboValidationTable
    Dim nTypeID As Variant
    
    Set clsComboValidation = New ComboValidationTable
    
    nTypeID = clsComboValidation.GetSingleValueFromValidation("ProcurationFeeType", PROCURATION_FEE_NON_PRODUCT_SPECIFIC)
    
    GetProcTypeID = nTypeID
End Function

Public Sub PopulateScreenControls()
    On Error GoTo Failed

    Dim rs As ADODB.Recordset
    Dim clsInsuranceFee As IntInsuranceProcuration
    Dim sText As String
    Dim sComboText As String
    Dim vBookMark As Variant
    Dim clsIntermediary As Intermediary
    Dim sGuid As String
    
    Set clsInsuranceFee = New IntInsuranceProcuration
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    'Get all Records for intermediary and type
    m_clsProcFee.GetFeeTypesForIntermediary sGuid, CStr(m_nFeeValueID)
    
    
    Set clsIntermediary = New Intermediary

    
    'Populate product period
    clsIntermediary.PopulateProcurationCombo frmEditProcurationFees.cboPackagingPeriod, m_clsProcFee, m_bIsEdit
    
    If frmEditProcurationFees.cboPackagingPeriod.RowSource.RecordCount > 0 Then
        sComboText = frmEditProcurationFees.cboPackagingPeriod.RowSource(frmEditProcurationFees.cboInsurancePeriod.ListField)
        frmEditProcurationFees.cboPackagingPeriod.Text = ""
        frmEditProcurationFees.cboPackagingPeriod.Text = sComboText
    End If

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetFeeType(nType As Integer)
    m_nFeeValueID = nType
End Sub

Public Sub SetIntermediaryGuid(vGuid As Variant)
    m_vIntermediaryGUID = vGuid
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetListViewHeaders
' Description   :   Sets the field header titles on the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "£ Range"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Fee"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "HO Split"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Company Split"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 24
    lvHeaders.sName = "Individual Split"
    headers.Add lvHeaders
    
    frmEditProcurationFees.lvPackaging.AddHeadings headers
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSplitTableClass
' Description   : Sets the proc split table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSplitTableClass(clsSplit As IntProcFeeSplitTable)
    On Error GoTo Failed
    
    Set m_clsIntProcFee = clsSplit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the listview with all records from the proc fee split table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed

    Dim clsProcFee As IntProcFeeSplitTable
    Dim sGuid As String
    
    Set clsProcFee = New IntProcFeeSplitTable
    
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    clsProcFee.GetAllFeeSplitsForGuid sGuid, CStr(m_nFeeValueID)
    
    g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvPackaging, clsProcFee, True

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleProductPeriod
' Description   : Called when the insurance period is changed. Updates the LV to only show fees for that period
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleProductPeriod(Optional bGetData As Boolean = True)
    On Error GoTo Failed

    Dim sSequence As String

    sSequence = frmEditProcurationFees.cboPackagingPeriod.BoundText
    
    If Len(sSequence) > 0 Then
        TableAccess(m_clsIntProcFee).ApplyFilter "TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSequence)
        
        g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvPackaging, m_clsIntProcFee, True
        TableAccess(m_clsIntProcFee).CancelFilter
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetSplitTableClass() As IntProcFeeSplitTable
    Set GetSplitTableClass = m_clsIntProcFee
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetSplitDetails
' Description   : Populates the split details
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSplitDetails()
    On Error GoTo Failed
    Dim sGuid As String

    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    m_clsIntProcFee.GetAllFeeSplitsForGuid sGuid, CStr(m_nFeeValueID)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveProduct
' Description   : Saves the proc fee product to the type record
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveProduct()
    On Error GoTo Failed

    Dim vVal As Variant
    Dim clsFeeType As IntermediaryProcFeeTable
    Dim sSequence As String
    
    Set clsFeeType = New IntermediaryProcFeeTable
    
    sSequence = frmEditProcurationFees.cboPackagingPeriod.BoundText
    
    TableAccess(clsFeeType).SetRecordSet frmEditProcurationFees.cboPackagingPeriod.RowSource
    If Len(sSequence) > 0 Then
        TableAccess(clsFeeType).ApplyFilter "TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSequence)
        
        g_clsFormProcessing.HandleComboExtra frmEditProcurationFees.cboPackagingPeriod, vVal, GET_CONTROL_VALUE
    
        If TableAccess(clsFeeType).RecordCount > 0 Then
            clsFeeType.SetProductID vVal
        
            TableAccess(clsFeeType).Update
        End If
    
        
        TableAccess(clsFeeType).CancelFilter
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

