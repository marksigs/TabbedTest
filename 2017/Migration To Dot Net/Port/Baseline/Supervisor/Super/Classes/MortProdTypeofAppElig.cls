VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTypeofAppElig"
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
Option Explicit

' Constants
Private Const m_sComboName As String = "TypeOfMortgage"

' Private data
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_colProductKeys As Collection
Private m_clsTypeOfAppTable As MortProdTypeofAppEligTable
Private m_clsComboTable As ComboValueTable
Private m_bIsEdit As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsTypeOfAppTable = New MortProdTypeofAppEligTable
    Set m_clsComboTable = New ComboValueTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Type
'                 of Application, including populating the swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed

    m_bIsEdit = bIsEdit

    PopulateComboTable m_sComboName
    PopulateSelectedItems
    
    PopulateAvailableItems frmProductDetails.MSGSwapTypeOfAppEligibility
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateComboTable
' Description   : Populates the combo table class with the Combo Group passed in.
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateComboTable(sGroup)
    On Error GoTo Failed
    Dim colKeys As New Collection
    Dim clsTableAccess As TableAccess
    
    Set m_clsComboTable = New ComboValueTable
    Set clsTableAccess = m_clsComboTable
    
    If Len(m_sComboName) > 0 Then
        colKeys.Add sGroup
        clsTableAccess.SetKeyMatchValues colKeys
        clsTableAccess.GetTableData
    
        clsTableAccess.ValidateData
        
        If clsTableAccess.RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurpOfLoan:PopulateComboTable - No entries for " + sGroup + " in Combo table"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "TypeOfApplicationEligibility: PopulateComboTable - " + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list of items available to move from the first swaplist listview
'                 to the second.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection
    Dim sValue As String
    Dim sValueID As String
    
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
            If DoesSwapValueExist(sValue) = False Then
                colLine.Add sValue
                SwapList.AddLineFirst colLine, clsSwapExtra
            End If
            clsTableAccess.MoveNext
        Loop
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "TypeOfAppEligibility - no avaiable items"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesSwapValueExist
' Description   : Checks if the value passed in sValue exists in the second listview of the
'                 swaplist MSGSwapTypeOfAppEligibility.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesSwapValueExist(sValue As String) As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sSecondValue As String
    Dim colValue As Collection
    
    nCount = frmProductDetails.MSGSwapTypeOfAppEligibility.GetSecondCount()

    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        Set colValue = frmProductDetails.MSGSwapTypeOfAppEligibility.GetLineSecond(nThisItem)
        
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
Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim sValue As String
    Dim sValueID As String
    Dim colLine As Collection
    Dim clsComboValues As New ComboValueTable
    Dim clsTableAccess As TableAccess
    Dim rsAppElig As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colValues As New Collection
    Dim enumPopulateType As PopulateType
    
    Set clsTableAccess = m_clsTypeOfAppTable

    enumPopulateType = POPULATE_EMPTY
    frmProductDetails.MSGSwapTypeOfAppEligibility.ClearSecond
    Set clsTableAccess = m_clsTypeOfAppTable

    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulateType = POPULATE_KEYS
    End If
    
    Set rsAppElig = clsTableAccess.GetTableData(enumPopulateType)
    clsTableAccess.ValidateData
    
    If clsTableAccess.RecordCount > 0 Then
        Set clsTableAccess = m_clsTypeOfAppTable
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        
        Set rsAppElig = clsTableAccess.GetTableData()
            
        clsTableAccess.ValidateData
        ' Need to read each index, then find the corresponding string from the
        ' combo table.
                    
        If rsAppElig.RecordCount > 0 Then
            rsAppElig.MoveFirst
            
            Do While Not rsAppElig.EOF
                sValueID = m_clsTypeOfAppTable.GetApplicationType()
                
                If Len(sValueID) > 0 Then
                    sValue = clsComboValues.FindComboValue(m_sComboName, sValueID)
                
                    If Len(sValue) > 0 Then
                        Set colLine = New Collection
                        Set clsSwapExtra = New SwapExtra
                        colLine.Add sValue
                        
                        clsSwapExtra.SetValueID sValueID
                        frmProductDetails.MSGSwapTypeOfAppEligibility.AddLineSecond colLine, clsSwapExtra
                    End If
                Else
                
                End If
            rsAppElig.MoveNext
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
Public Function SetScreenFields() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
    bRet = True

    SetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "SetScreenFields: Error is " + Err.DESCRIPTION
    SetScreenFields = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all selected items on the Type of Application swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    Dim nSelectedCount As Integer
    Dim nThisItem As Integer
    Dim sValueID As String
    Dim sValue As String
    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    If frmProductDetails.MSGSwapTypeOfAppEligibility.HasChanged() Then
        Set clsTableAccess = m_clsTypeOfAppTable
        
        clsTableAccess.ValidateData
        clsTableAccess.DeleteAllRows
        
        nSelectedCount = frmProductDetails.MSGSwapTypeOfAppEligibility.GetSecondCount()
    
        For nThisItem = 1 To nSelectedCount
            Set colValues = frmProductDetails.MSGSwapTypeOfAppEligibility.GetLineSecond(nThisItem, clsSwapExtra)
            
            If colValues.Count > 0 Then
                sValue = colValues(1)
            End If
            
            If Not clsSwapExtra Is Nothing Then
                sValueID = clsSwapExtra.GetValueID()
            End If
        
            If Len(sValueID) > 0 Then
                clsTableAccess.AddRow
            
                m_clsTypeOfAppTable.SetProductCode m_sProductCode
                m_clsTypeOfAppTable.SetStartDate m_sProductStartDate
                m_clsTypeOfAppTable.SetApplicationType sValueID
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
' Description   : Updates the Type of Application Eligibility table once all values have been
'                 written to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsTypeOfAppTable).Update
    
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
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Type of App tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
End Function
