VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypeOfBuyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTypeOfAppElig
' Description   : Provides functionality for the Type of Application Eligibility tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM history
' Prog      Date        Description
' TW        02/02/2007  EP2_1118 - Optimus - not getting products back for a credit limit increase
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
' TW 02/02/2007 EP2_1118
'Private Const m_sComboTypeOfBuyer As String = "TypeOfBuyerNewLoan"
Private Const m_sComboTypeOfBuyer As String = "TypeOfBuyerSupervisor"
' TW 02/02/2007 EP2_1118 End

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_colProductKeys As Collection
Private m_bTypeOfBuyerExists As Boolean
Private m_clsTypeOfBuyer As MortProdTypeOfBuyerTable
Private m_clsComboTable As New ComboValueTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsTypeOfBuyer = New MortProdTypeOfBuyerTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for
'                 Type of Buyer including the swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed

    m_bIsEdit = bIsEdit

    PopulateSelectedItems frmProductDetails.MSGSwapTypeOfBuyer
    PopulateAvailableItems frmProductDetails.MSGSwapTypeOfBuyer, m_sComboTypeOfBuyer

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
    Dim colKeys As New Collection
    Dim clsTableAccess As TableAccess
    
    Set m_clsComboTable = New ComboValueTable
    Set clsTableAccess = m_clsComboTable
    
    If Len(sGroup) > 0 Then
        colKeys.Add sGroup
        clsTableAccess.SetKeyMatchValues colKeys
        clsTableAccess.GetTableData
    
        clsTableAccess.ValidateData

        If clsTableAccess.RecordCount() = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - No entries for " + sGroup + " in Combo table"
        End If
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
        'Set colValue = frmProductDetails.MSGSwapTypeOfBuyer.GetLineSecond(nThisItem)
        
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
    Dim rsTypeOfBuyer As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim enumPopulateType As PopulateType
    Dim colValues As New Collection
    
    enumPopulateType = POPULATE_EMPTY
    frmProductDetails.MSGSwapTypeOfBuyer.ClearSecond
    Set clsTableAccess = m_clsTypeOfBuyer

    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulateType = POPULATE_KEYS
    End If
    
    Set clsTableAccess = m_clsTypeOfBuyer
    Set rsTypeOfBuyer = clsTableAccess.GetTableData(enumPopulateType)
    clsTableAccess.ValidateData
                
    If clsTableAccess.RecordCount > 0 Then
        rsTypeOfBuyer.MoveFirst
        
        Do While Not rsTypeOfBuyer.EOF
            sValueID = m_clsTypeOfBuyer.GetTypeOfBuyer()
            
            If Len(sValueID) > 0 Then
                sValue = clsComboValues.FindComboValue(m_sComboTypeOfBuyer, sValueID)
            
                If Len(sValue) > 0 Then
                    Set colLine = New Collection
                    Set clsSwapExtra = New SwapExtra
                    colLine.Add sValue
                    
                    clsSwapExtra.SetValueID sValueID
                    frmProductDetails.MSGSwapTypeOfBuyer.AddLineSecond colLine, clsSwapExtra
                End If
            Else
            
            End If
        rsTypeOfBuyer.MoveNext
        Loop
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
' Description   : Saves all selected items on the Type of Buyer swaplist.
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
    
    If frmProductDetails.MSGSwapTypeOfBuyer.HasChanged() Then
        Set clsTableAccess = m_clsTypeOfBuyer
        
        clsTableAccess.ValidateData "TypeOfBuyer"
        clsTableAccess.DeleteAllRows
        
        nSelectedCount = frmProductDetails.MSGSwapTypeOfBuyer.GetSecondCount()
    
        For nThisItem = 1 To nSelectedCount
            Set colValues = frmProductDetails.MSGSwapTypeOfBuyer.GetLineSecond(nThisItem, clsSwapExtra)
            
            If colValues.Count > 0 Then
                sValue = colValues(1)
            End If
            
            If Not clsSwapExtra Is Nothing Then
                sValueID = clsSwapExtra.GetValueID()
            End If
        
            If Len(sValueID) > 0 Then
                clsTableAccess.AddRow
            
                m_clsTypeOfBuyer.SetProductCode m_sProductCode
                m_clsTypeOfBuyer.SetStartDate m_sProductStartDate
                m_clsTypeOfBuyer.SetTypeOfBuyer sValueID
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
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdTypeOfBuyer.SetProductKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates the Type of Application Buyer table once all values have been
'                 written to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsTypeOfBuyer).Update
    
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
Public Function ValidateScreenData()
    On Error GoTo Failed
    Dim bRet As Boolean
    bRet = True
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
   
