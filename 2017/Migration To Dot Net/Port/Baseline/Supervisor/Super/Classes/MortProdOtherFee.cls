VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdOtherFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdOtherFee
' Description   : Provides functionality for the Other Fee tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       06/12/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
' STB       11/01/02    SYS1960 DoesSwapValueExist() wasn't resetting nColumnCount
'                       at the appropriate point.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_vOrgID As Variant
Private m_bIsEdit As Boolean
Private m_clsOtherFees As OtherFeeTable
Private m_clsLenderOtherFees As LenderOtherFeesTable
Private m_sProductCode As String
Private m_sProductStartDate As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsOtherFees = New OtherFeeTable
    Set m_clsLenderOtherFees = New LenderOtherFeesTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    m_bIsEdit = bIsEdit

    PopulateSelectedItems
    PopulateAvailableItems frmProductDetails.MSGSwapOtherFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdOtherFee:Initialise - " + Err.DESCRIPTION
End Sub

Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection

    Dim sName As String
    Dim sStartDate As String
    Dim sAmount As String
    
    Set clsTableAccess = m_clsLenderOtherFees

    ' First populate the Other Fees for the Lender associated with this product.
    PopulateLenderOtherFees

    If IsLenderSelected Then
        SwapList.ClearFirst
        SwapList.EnableButtons
        Set rs = clsTableAccess.GetRecordSet

        If clsTableAccess.RecordCount > 0 Then
            clsTableAccess.MoveFirst
    
            Do While Not rs.EOF
                Set colLine = New Collection
                Set clsSwapExtra = New SwapExtra
    
                sName = m_clsLenderOtherFees.GetName()
                sStartDate = m_clsLenderOtherFees.GetLenderOtherFeeStartDate()
                sAmount = m_clsLenderOtherFees.GetAmount()
                    
                clsSwapExtra.SetValueID sName
                clsSwapExtra.SetStartDate sStartDate
                clsSwapExtra.SetAmount sAmount
                
                colLine.Add sName
                colLine.Add sStartDate
                
                ' Does this value exist in the Selected Items Listbox?
                If DoesSwapValueExist(colLine) = False Then
                    SwapList.AddLineFirst colLine, clsSwapExtra
                End If
                clsTableAccess.MoveNext
            Loop
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdOtherFee:Initialise -Err.DESCRIPTION"
End Sub

Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim sName As String
    Dim sStartDate As String
    Dim sAmount As String
    Dim colLine As Collection
    Dim clsTableAccess As TableAccess
    Dim clsLenderOtherFees As TableAccess
    Dim rsOtherFees As ADODB.Recordset
    Dim rsLenderOtherFees As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colValues As New Collection
    
    Set clsTableAccess = m_clsOtherFees
    Set clsLenderOtherFees = m_clsLenderOtherFees

    If IsLenderSelected Then
        colValues.Add m_sProductCode
        colValues.Add m_sProductStartDate

        clsTableAccess.SetKeyMatchValues colValues

        Set rsOtherFees = clsTableAccess.GetTableData()

        clsTableAccess.ValidateData

        If rsOtherFees.RecordCount > 0 Then
            rsOtherFees.MoveFirst

            Do While Not rsOtherFees.EOF
                sName = m_clsOtherFees.GetName()
                sStartDate = m_clsOtherFees.GetOtherFeeStartDate()
                sAmount = m_clsOtherFees.GetAmount()
                
                If Len(sName) > 0 Then
                    Set colLine = New Collection
                    Set clsSwapExtra = New SwapExtra
                    
                    colLine.Add sName
                    colLine.Add sStartDate
                    
                    clsSwapExtra.SetValueID sName
                    clsSwapExtra.SetStartDate sStartDate
                    clsSwapExtra.SetAmount sAmount
                                        
                    frmProductDetails.MSGSwapOtherFees.AddLineSecond colLine, clsSwapExtra
                End If

                rsOtherFees.MoveNext
            Loop
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdPurposeOfLoan:PopulateSelectedItems - " + Err.DESCRIPTION
End Sub

Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim nSelectedCount As Integer
    Dim nThisItem As Integer
    Dim sName As String
    Dim sAmount As String
    Dim sOTherFeeStartDate As String
    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess

    nSelectedCount = frmProductDetails.MSGSwapOtherFees.GetSecondCount()
    Set clsTableAccess = m_clsOtherFees
    
    If IsLenderSelected Then
        If clsTableAccess.RecordCount > 0 Then
            clsTableAccess.DeleteAllRows
        Else
            clsTableAccess.GetTableData POPULATE_EMPTY
        End If
    
        For nThisItem = 1 To nSelectedCount
            Set colValues = frmProductDetails.MSGSwapOtherFees.GetLineSecond(nThisItem, clsSwapExtra)
    
            sName = clsSwapExtra.GetValueID()
            sOTherFeeStartDate = clsSwapExtra.GetStartDate()
            sAmount = clsSwapExtra.GetAmount()
    
            If Len(sName) > 0 Then
                clsTableAccess.AddRow
    
                m_clsOtherFees.SetProductCode m_sProductCode
                m_clsOtherFees.SetStartDate m_sProductStartDate
                m_clsOtherFees.SetName sName
                m_clsOtherFees.SetOtherFeesStartDate sOTherFeeStartDate
                m_clsOtherFees.SetAmount sAmount
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
            End If
        Next

        clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.Number
End Sub

Public Sub DoUpdates()
    On Error GoTo Failed

    TableAccess(m_clsOtherFees).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.Number
End Sub

Public Function SetLender(vOrgID As Variant) As Boolean
    On Error GoTo Failed
    Dim nResp As Integer
    Dim bSuccess As Boolean
    Dim nSelectedItems As Long
    Dim sGUID As String
    Dim sOriginalGUID As String
    bSuccess = True
    
    nSelectedItems = frmProductDetails.MSGSwapOtherFees.GetSecondCount
    
    If Not IsEmpty(vOrgID) And Not IsEmpty(m_vOrgID) Then
        sGUID = g_clsSQLAssistSP.GuidToString(CStr(vOrgID))
        sOriginalGUID = g_clsSQLAssistSP.GuidToString(CStr(m_vOrgID))
        
        If nSelectedItems > 0 And sGUID <> sOriginalGUID Then
            nResp = MsgBox("Changing the Lender will remove all selected Lender Other Fees for this product. Continue?", vbQuestion + vbYesNo)
            
            If nResp = vbYes Then
                SetOtherFeeState
            Else
                bSuccess = False
            End If
        End If
    End If
    
    If bSuccess Then
        m_vOrgID = vOrgID
    End If
    
    SetLender = bSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetProductKeys(colKeyValues As Collection)
    On Error GoTo Failed
    
    m_sProductCode = colKeyValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colKeyValues(PRODUCT_START_DATE_KEY)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function IsLenderSelected() As Boolean
    On Error GoTo Failed
    Dim bSelected As Boolean
    
    bSelected = False
    If Not IsEmpty(m_vOrgID) Then
        If Len(m_vOrgID) > 0 Then
            bSelected = True
        End If
    End If
    
    IsLenderSelected = bSelected
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub PopulateLenderOtherFees()
    On Error GoTo Failed
    Dim bLenderSpecified As Boolean
    Dim colValues As Collection
    
    bLenderSpecified = IsLenderSelected
    
    ' Only repopulate if we have a lender specified and there are no existing lender other fees
    If bLenderSpecified Then
        Set colValues = New Collection
        bLenderSpecified = True
        colValues.Add m_vOrgID
        TableAccess(m_clsLenderOtherFees).SetKeyMatchValues colValues
        TableAccess(m_clsLenderOtherFees).GetTableData
    Else
        SetOtherFeeState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "PopulateLenderOtherFees - " + Err.DESCRIPTION
End Sub

Private Sub SetOtherFeeState()
    On Error GoTo Failed
    Dim bLenderSpecified As Boolean
    Dim colLine As Collection
    Dim bEnable As Boolean
    
    bEnable = True
    bLenderSpecified = IsLenderSelected
    
    frmProductDetails.MSGSwapOtherFees.ClearFirst
    frmProductDetails.MSGSwapOtherFees.ClearSecond
    
    If Not bLenderSpecified Then
        Set colLine = New Collection
        colLine.Add "No Lender Selected"
        frmProductDetails.MSGSwapOtherFees.AddLineFirst colLine
        bEnable = False
    End If
    
    frmProductDetails.MSGSwapOtherFees.EnableButtons bEnable
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoesSwapValueExist(ByRef colFirst As Collection) As Boolean
    
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sFirstValue As String
    Dim sSecondValue As String
    Dim nColumnCount As Integer
    Dim colSecond As Collection
    Dim nThisColumn As Integer
    Dim bNoMatch As Boolean
        
    On Error GoTo Failed
    
    nCount = frmProductDetails.MSGSwapOtherFees.GetSecondCount()
    
    bFound = False
    nThisItem = 1

    nColumnCount = colFirst.Count
    nThisColumn = 1
    
    While bFound = False And nThisItem <= nCount
        'Set colSecond = New Collection
        Set colSecond = frmProductDetails.MSGSwapOtherFees.GetLineSecond(nThisItem)
        
        bNoMatch = False
        nThisColumn = 1
        
        While bNoMatch = False And nThisColumn <= nColumnCount
            sFirstValue = colFirst(nThisColumn)
            sSecondValue = colSecond(nThisColumn)

            If sFirstValue <> sSecondValue Then
                bNoMatch = True
            End If
        
            nThisColumn = nThisColumn + 1
        Wend

        If bNoMatch = False Then
            bFound = True
        Else
            nThisItem = nThisItem + 1
        End If
    Wend
    
    DoesSwapValueExist = bFound
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub InitTab()
    On Error GoTo Failed
    
    PopulateAvailableItems frmProductDetails.MSGSwapOtherFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

