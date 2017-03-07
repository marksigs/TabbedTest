VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntNonSpecificProcuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : IntermediaryNonProdSpecificProcuration
' Description   : Procuration non-mortgage specific tab
'
' Change history
' Prog      Date        Description
' AA        26/06/2001  created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsProcSplit As IntProcFeeSplitForIntTable
Private m_clsProcFee As IntermediaryProcFeeTable
Private m_clsIntProcFee As IntProcFeeSplitTable
Private m_nFeeValueID As Long
Private m_vIntermediaryGUID As Variant
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Public Sub Initialise(bIsEdit As Boolean, Optional bGetData As Boolean = True)
    On Error GoTo Failed
    
    Set m_clsProcFee = New IntermediaryProcFeeTable
    m_bIsEdit = bIsEdit
        
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
    Dim sGuid As String
    Dim clsIntermediary As Intermediary
    Dim sComboText As String
    
    Set clsIntermediary = New Intermediary
    sGuid = g_clsSQLAssistSP.GuidToString(CStr(m_vIntermediaryGUID))
    
    'Get all Records for intermediary and type
    m_clsProcFee.GetFeeTypesForIntermediary sGuid, CStr(m_nFeeValueID)
    clsIntermediary.PopulateProcurationCombo frmEditProcurationFees.cboNonProdSpecificPeriod, m_clsProcFee, m_bIsEdit
    
    
    If frmEditProcurationFees.cboNonProdSpecificPeriod.RowSource.RecordCount > 0 Then
        frmEditProcurationFees.cboNonProdSpecificPeriod.Text = ""
        sComboText = frmEditProcurationFees.cboNonProdSpecificPeriod.RowSource(frmEditProcurationFees.cboNonProdSpecificPeriod.ListField)
        
        frmEditProcurationFees.cboNonProdSpecificPeriod.Text = sComboText
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
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
    
    frmEditProcurationFees.lvnonMortgageSpecific.AddHeadings headers
    
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
' Function      : SaveScreenData
' Description   : SaveScreenData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed

    Dim sGuid As Variant
    Dim vTmp As Variant
    Dim sType As String
    Dim sTypeSequence As String
    
    If Not m_bIsEdit Then
        'guid
        vTmp = m_clsProcFee.GetIntermediaryGUID
        sGuid = g_clsSQLAssistSP.GuidToString(CStr(vTmp))
        'Type
        sType = m_clsProcFee.GetFeeType
        'Type Sequence Number
        sTypeSequence = m_clsProcFee.GetTypeSequence
    End If
    
    
    
    
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
    
    g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvnonMortgageSpecific, clsProcFee, True

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

    sSequence = frmEditProcurationFees.cboNonProdSpecificPeriod.BoundText
    
    If Len(sSequence) > 0 Then
        TableAccess(m_clsIntProcFee).ApplyFilter "TypeSequenceNumber = " & g_clsSQLAssistSP.FormatString(sSequence)
        
        g_clsFormProcessing.PopulateFromRecordset frmEditProcurationFees.lvnonMortgageSpecific, m_clsIntProcFee, True
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
' Function      : Delete
' Description   : Deletes the current Record from the LV
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete(lv As MSGListView)
    On Error GoTo Failed
    
    Dim lvItem As MSComctlLib.ListItem
    Dim sSequence As String
    Dim clsPopulateDetails As PopulateDetails
    

    Set lvItem = lv.SelectedItem
    
    If Not lvItem Is Nothing Then
        Set clsPopulateDetails = lv.GetExtra(lvItem.Index)
        sSequence = clsPopulateDetails.GetExtra
        m_clsIntProcFee.GetSplitFromSequence sSequence
        
        If TableAccess(m_clsIntProcFee).RecordCount > 0 Then
            'TableAccess(m_clsIntProcFee).DeleteAllRows
            lv.ListItems.Remove lvItem.Index
        End If
        
        TableAccess(m_clsIntProcFee).CancelFilter
    End If
    
    
    
    Exit Sub
Failed:
    TableAccess(m_clsIntProcFee).CancelFilter
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
