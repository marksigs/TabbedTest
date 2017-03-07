VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdProdCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdProdCond
' Description   : Provides functionality for the Special Condition tab on Mortgage Product.
' Change history
' Prog      Date        Description
' GD        20/05/02    BMIDS0017 - File Created.
' GD        05/06/02    BMIDS00016 Change Column headers for Swap Boxes
' GD        23/07/02    BMIDS00232 - Property Location selected items were not saving.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const m_sComboConditions As String = "Conditions"

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_sConditionReference As String
Private m_colProductKeys As Collection
Private m_clsMortProdProdCond As MortProdProdCondTable
'Private m_clsComboTable As New ComboValueTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsMortProdProdCond = New MortProdProdCondTable
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Property
'                 Location, including populating the swaplists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
'GD BMIDS0012 Initialise Swap Boxes
    On Error GoTo Failed

    m_bIsEdit = bIsEdit
    
    'Set headers for each list
    SetColumnHeaders

    ' Populate the SwapList
    PopulateSelectedItems frmProductDetails.MSGSwapProdCond
    PopulateAvailableItems frmProductDetails.MSGSwapProdCond

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetColumnHeaders()
'GD BMIDS0012 Set up column headings for listviews (both)
    Dim colHeaders As New Collection
    Dim clsHeader As listViewAccess
    'GD BMIDS00016 - Change column header
    clsHeader.nWidth = 50
    clsHeader.sName = "Ref"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 50
    clsHeader.sName = "Condition Name"
    colHeaders.Add clsHeader
    

    frmProductDetails.MSGSwapProdCond.SetFirstColumnHeaders colHeaders
    frmProductDetails.MSGSwapProdCond.SetSecondColumnHeaders colHeaders

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list(lhs) on data that doesn't exist in list on rhs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
'GD BMIDS0012 Populate a list of items that the user can choose.
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection
    Dim sValue As String
    Dim sValueID As String
    Dim sDescription As String
    
    SwapList.ClearFirst
    Set clsTableAccess = m_clsMortProdProdCond
    m_clsMortProdProdCond.GetAllConditions
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra

            sValueID = rs.fields("CONDITIONREFERENCE").Value
            sDescription = rs.fields("CONDITIONNAME").Value
            clsSwapExtra.SetValueID sValueID

            ' Does this value exist in the Selected Items Listbox?
            If DoesSwapValueExist(sValueID, SwapList) = False Then
                colLine.Add sValueID
                colLine.Add sDescription
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
'                 swaplist MSGSwapPropLocation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesSwapValueExist(sValue As String, msgSwap As MSGHorizontalSwapList) As Boolean
'GD BMIDS0012
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
        If colValue.Count >= 1 Then
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
' Description   : Populates items from MORTGAGEPRODUCTCONDITION associated with this particular product and startdate
'                 in second list view
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateSelectedItems(SwapList As MSGHorizontalSwapList)
'GD BMIDS0012
    On Error GoTo Failed
    Dim sRow As String
    Dim colLine As Collection
    Dim sConditionReference As String
    Dim sConditionName As String
    Dim clsTableAccess As TableAccess
    Dim rsMortProdCond As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra

    Dim enumPopulateType As PopulateType

    Set clsTableAccess = m_clsMortProdProdCond

    enumPopulateType = POPULATE_EMPTY
    ' GD BMIDS00232 frmProductDetails.MSGSwapPropLocation.ClearSecond
    frmProductDetails.MSGSwapProdCond.ClearSecond
    Set clsTableAccess = m_clsMortProdProdCond
    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulateType = POPULATE_KEYS
    End If
    Set rsMortProdCond = clsTableAccess.GetTableData(enumPopulateType)
    clsTableAccess.ValidateData
    If clsTableAccess.RecordCount > 0 Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        Set rsMortProdCond = clsTableAccess.GetTableData(enumPopulateType)
        If Not rsMortProdCond Is Nothing Then
            If rsMortProdCond.RecordCount > 0 Then
                rsMortProdCond.MoveFirst
                Do While Not rsMortProdCond.EOF
                    sConditionReference = m_clsMortProdProdCond.GetConditionReference()
                    'Empty String returned is condition has been deleted(flag set)or validation type
                    'has been set to something other than 'P' since last save
                    If Len(sConditionReference) > 0 Then
                        sConditionName = m_clsMortProdProdCond.GetConditionName(sConditionReference)
                        'Empty String returned is condition has been deleted or validation type
                        'has been set to something other than 'P' since last save
                        If Len(sConditionName) > 0 Then
                            Set colLine = New Collection
                            Set clsSwapExtra = New SwapExtra
                            colLine.Add sConditionReference
                            colLine.Add sConditionName
                            frmProductDetails.MSGSwapProdCond.AddLineSecond colLine, clsSwapExtra
                        End If
                    End If
                rsMortProdCond.MoveNext
                Loop
            End If
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdProdCond:PopulateSelectedItems - " + Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called externally, doesn't do anything for this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdProdCond:SetScreenFields - " + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all selected items on the MortgageProductCondition swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
'GD BMIDS0012
    On Error GoTo Failed
    Dim nSelectedCount As Integer
    Dim nThisItem As Integer
    Dim sValueID As String
    Dim sValue As String
    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim clsMortProdProdCond As MortProdProdCondTable
    Set clsMortProdProdCond = New MortProdProdCondTable


    Set clsTableAccess = m_clsMortProdProdCond
    m_clsMortProdProdCond.GetMortgageProductCondition m_sProductCode, m_sProductStartDate
    
    clsTableAccess.DeleteAllRows

    nSelectedCount = frmProductDetails.MSGSwapProdCond.GetSecondCount()

    For nThisItem = 1 To nSelectedCount
        Set colValues = frmProductDetails.MSGSwapProdCond.GetLineSecond(nThisItem, clsSwapExtra)

        If colValues.Count > 0 Then
            sValue = colValues(1)
        End If
        If Len(sValue) > 0 Then
            clsTableAccess.AddRow
            m_clsMortProdProdCond.SetConditionReference sValue
            m_clsMortProdProdCond.SetProductCode m_sProductCode
            m_clsMortProdProdCond.SetStartDate m_sProductStartDate
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdProdCond:SaveScreenData - 'Value' is empty"
        End If
    Next
    clsTableAccess.Update
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates the table once all values have been written to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsMortProdProdCond).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Property Location tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
'GD BMIDS0012
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
'GD BMIDS0012
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdProdCond.SetProductKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)

    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



