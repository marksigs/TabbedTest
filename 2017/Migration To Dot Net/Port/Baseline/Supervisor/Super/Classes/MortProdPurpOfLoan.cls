VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdPurpOfLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdPurpOfLoan
' Description   : Provides functionality for the Purpose of Loan, Type of Application and Type
'                 Of Buyer functionality on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const m_sComboPurposeOfLoan As String = "PurposeOfLoanSupervisor"

' Private data
Private m_bIsEdit As Boolean
Private m_colProductKeys  As Collection
Private m_sProductCode As String
Private m_sProductStartDate As String

Private m_clsPurpOfLoan As MortProdPurpOfLoanTable
Private m_clsComboTable As New ComboValueTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsPurpOfLoan = New MortProdPurpOfLoanTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Purpose
'                 of Loan, including populating the swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed

    m_bIsEdit = bIsEdit
 
    PopulateSelectedItems frmProductDetails.MSGSwapPurpOfLoan
    
    PopulateAvailableItems frmProductDetails.MSGSwapPurpOfLoan, m_sComboPurposeOfLoan
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateComboTable
' Description   : Populates the combo table class with the Combo Group passed in.
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateComboTable(sGroup As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim colKeys As New Collection
    Dim clsTableAccess As TableAccess
    
    Set m_clsComboTable = New ComboValueTable
    Set clsTableAccess = m_clsComboTable
    
    If Len(sGroup) > 0 Then
        colKeys.Add sGroup
        clsTableAccess.SetKeyMatchValues colKeys
        Set rs = clsTableAccess.GetTableData()
    
        clsTableAccess.ValidateData
        If rs.RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - No entries for " + sGroup + " in Combo table"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - Group is empty"
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list of items available to move from the first swaplist listview
'                 to the second.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList, sComboGroup As String)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection
    Dim sValue As String
    Dim sValueID As String
    
    SwapList.ClearFirst
    
    PopulateComboTable sComboGroup
    
    Set clsTableAccess = m_clsComboTable
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.ValidateData
        
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sValue = m_clsComboTable.GetValueName()
            sValueID = m_clsComboTable.GetValueID()
            clsSwapExtra.SetValueID sValueID
            
            ' Does this value exist in the Selected Items Listbox?
            If DoesSwapValueExist(sValue, SwapList) = False Then
                colLine.Add sValue
                SwapList.AddLineFirst colLine, clsSwapExtra
            End If
            clsTableAccess.MoveNext
        Loop
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesSwapValueExist
' Description   : Checks if the value passed in sValue exists in the second listview of the
'                 swaplist msgSwap.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesSwapValueExist(sValue As String, msgSwap As MSGHorizontalSwapList) As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sSecondValue As String
    Dim colValue As Collection
    
    nCount = msgSwap.GetSecondCount()
    
    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        
        Set colValue = msgSwap.GetLineSecond(nThisItem)
        
        If colValue.Count = 1 Then
            sSecondValue = colValue(1)
        
            If sSecondValue = sValue Then
                bFound = True
            End If
        End If
        
        nThisItem = nThisItem + 1
    Wend
    
    DoesSwapValueExist = bFound
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateSelectedItems
' Description   : Populates a list of items that have been moved from the first listview of the
'                 swaplist to the second.
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateSelectedItems(SwapList As MSGHorizontalSwapList)
    On Error GoTo Failed
    Dim sValue As String
    Dim sValueID As String
    Dim colLine As Collection
    Dim clsComboValues As New ComboValueTable
    Dim clsTableAccess As TableAccess
    Dim rsPurposeOfLoan As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colValues As New Collection
    Dim enumPopulate As PopulateType
    Set clsTableAccess = m_clsPurpOfLoan

    enumPopulate = POPULATE_EMPTY
    
    frmProductDetails.MSGSwapPurpOfLoan.ClearSecond
    Set clsTableAccess = m_clsPurpOfLoan
    
    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulate = POPULATE_KEYS
    End If
    
    Set rsPurposeOfLoan = clsTableAccess.GetTableData(enumPopulate)
    
    If clsTableAccess.RecordCount > 0 Then
        ' Need to read each index, then find the corresponding string from the
        ' combo table.
                    
        If rsPurposeOfLoan.RecordCount > 0 Then
            rsPurposeOfLoan.MoveFirst
            
            Do While Not rsPurposeOfLoan.EOF
                sValueID = m_clsPurpOfLoan.GetPurposeOfLoan()
                
                If Len(sValueID) > 0 Then
                    sValue = clsComboValues.FindComboValue(m_sComboPurposeOfLoan, sValueID)
                
                    If Len(sValue) > 0 Then
                        Set colLine = New Collection
                        Set clsSwapExtra = New SwapExtra
                        colLine.Add sValue
                        
                        clsSwapExtra.SetValueID sValueID
                        frmProductDetails.MSGSwapPurpOfLoan.AddLineSecond colLine, clsSwapExtra
                    End If
                Else
                
                End If
                rsPurposeOfLoan.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called externally, doesn't do anything for this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all selected items on the Property Location swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim nSelectedCount As Integer
    Dim nThisItem As Integer
    Dim sValueID As String
    Dim sValue As String
    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    If frmProductDetails.MSGSwapPurpOfLoan.HasChanged() Then
        Set clsTableAccess = m_clsPurpOfLoan
        
        clsTableAccess.ValidateData "PurposeOfLoan"
        
        clsTableAccess.DeleteAllRows
    
        nSelectedCount = frmProductDetails.MSGSwapPurpOfLoan.GetSecondCount()
    
        For nThisItem = 1 To nSelectedCount
            Set colValues = frmProductDetails.MSGSwapPurpOfLoan.GetLineSecond(nThisItem, clsSwapExtra)
            
            If colValues.Count > 0 Then
                sValue = colValues(1)
            End If
            
            If Not clsSwapExtra Is Nothing Then
                sValueID = clsSwapExtra.GetValueID()
            End If
        
            If Len(sValueID) > 0 Then
                clsTableAccess.AddRow
                
                m_clsPurpOfLoan.SetProductCode m_sProductCode
                m_clsPurpOfLoan.SetStartDate m_sProductStartDate
                m_clsPurpOfLoan.SetPurposeOfLoan sValueID
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
            End If
        Next
        
        clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates the Purpose of Loan table once all values have been written to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    
    bRet = True
    Set clsTableAccess = m_clsPurpOfLoan
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally when this tab is switched to.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Set frmLenderDetails.lvLegalFees.SelectedItem = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Purpose of Loan tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData()
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "AdditionalMortProdParams.SetMortgageParameterKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

