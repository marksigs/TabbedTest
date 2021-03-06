VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncomeStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncomeStatus
' Description   : Support class for the Mortgage Product Income Status tab on the Mortgage Product screen
' Change history
' Prog      Date        Description
' TW        30/11/2006  EP2_253 - Created
' TW        19/03/2007  EP2_1830 - When adding Mortgage Product in Supervisor error message displayed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_sIncomeStatus As String
Private m_colProductKeys As Collection
Private m_clsMortProdIncomeStatus As MortProdIncomeStatusTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsMortProdIncomeStatus = New MortProdIncomeStatusTable
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

    ' Populate the SwapList
    PopulateSelectedItems frmProductDetails.MSGSwapIncomeStatus
    PopulateAvailableItems frmProductDetails.MSGSwapIncomeStatus

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list(lhs) on data that doesn't exist in list on rhs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
''GD BMIDS0012 Populate a list of items that the user can choose.
'    On Error GoTo Failed
'    Dim clsTableAccess As TableAccess
'    Dim rs As ADODB.Recordset
'    Dim clsSwapExtra As SwapExtra
'    Dim colLine As Collection
'    Dim sValue As String
'    Dim sValueId As String
'    Dim sDescription As String
'
'    SwapList.ClearFirst
'    Set clsTableAccess = m_clsMortProdIncomeStatus
'    m_clsMortProdIncomeStatus.GetAvailableIncomeStatusSettings m_colProductKeys(1), m_colProductKeys(2)
'    If clsTableAccess.RecordCount() > 0 Then
'        Set rs = clsTableAccess.GetRecordSet
'
'        clsTableAccess.MoveFirst
'
'        Do While Not rs.EOF
'            Set colLine = New Collection
''            Set clsSwapExtra = New SwapExtra
'
'            sDescription = rs.fields("IncomeStatus").Value
''            clsSwapExtra.SetValueID sValueId
'
'            ' Does this value exist in the Selected Items Listbox?
'            If DoesSwapValueExist(sValueId, SwapList) = False Then
'                colLine.Add sDescription
''                SwapList.AddLineFirst colLine, clsSwapExtra
'            End If
'            clsTableAccess.MoveNext
'        Loop
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

'-----------
Dim colLine As Collection
Dim clsSwapExtra As SwapExtra
Dim clsTableAccess As TableAccess

Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    SwapList.ClearFirst
    Set clsTableAccess = m_clsMortProdIncomeStatus

' TW 19/03/2007 EP2_1830
'    m_clsMortProdIncomeStatus.GetAvailableIncomeStatusSettings m_colProductKeys(1), m_colProductKeys(2)
    If m_bIsEdit Then
        m_clsMortProdIncomeStatus.GetAvailableIncomeStatusSettings m_colProductKeys(1), m_colProductKeys(2)
    Else
        m_clsMortProdIncomeStatus.GetAvailableIncomeStatusSettings "", ""
    End If
' TW 19/03/2007 EP2_1830 End
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra

            clsSwapExtra.SetValueID rs.fields("VALUEID").Value
            colLine.Add rs.fields("VALUENAME").Value
            SwapList.AddLineFirst colLine, clsSwapExtra

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
''GD BMIDS0012
'    On Error GoTo Failed
'    Dim sRow As String
'    Dim colLine As Collection
'    Dim sIncomeStatus As String
'    Dim sConditionName As String
'    Dim clsTableAccess As TableAccess
'    Dim rsMortProdCond As ADODB.Recordset
'    Dim clsSwapExtra As SwapExtra
'
'    Dim enumPopulateType As PopulateType
'
'    Set clsTableAccess = m_clsMortProdIncomeStatus
'
'    enumPopulateType = POPULATE_EMPTY
'    ' GD BMIDS00232 frmProductDetails.MSGSwapPropLocation.ClearSecond
'    frmProductDetails.MSGSwapProdCond.ClearSecond
'    Set clsTableAccess = m_clsMortProdIncomeStatus
'    If Not m_colProductKeys Is Nothing Then
'        clsTableAccess.SetKeyMatchValues m_colProductKeys
'        enumPopulateType = POPULATE_KEYS
'    End If
'    Set rsMortProdCond = clsTableAccess.GetTableData(enumPopulateType)
'    clsTableAccess.ValidateData
'    If clsTableAccess.RecordCount > 0 Then
'        clsTableAccess.SetKeyMatchValues m_colProductKeys
'        Set rsMortProdCond = clsTableAccess.GetTableData(enumPopulateType)
'        If Not rsMortProdCond Is Nothing Then
'            If rsMortProdCond.RecordCount > 0 Then
'                rsMortProdCond.MoveFirst
'                Do While Not rsMortProdCond.EOF
'                    sIncomeStatus = m_clsMortProdIncomeStatus.GetIncomeStatus()
'                    'Empty String returned is condition has been deleted(flag set)or validation type
'                    'has been set to something other than 'P' since last save
'                    If Len(sIncomeStatus) > 0 Then
''                        sConditionName = m_clsMortProdIncomeStatus.GetConditionName(sIncomeStatus)
'                        'Empty String returned is condition has been deleted or validation type
'                        'has been set to something other than 'P' since last save
'                        If Len(sConditionName) > 0 Then
'                            Set colLine = New Collection
'                            Set clsSwapExtra = New SwapExtra
'                            colLine.Add sIncomeStatus
'                            colLine.Add sConditionName
'                            frmProductDetails.MSGSwapProdCond.AddLineSecond colLine, clsSwapExtra
'                        End If
'                    End If
'                rsMortProdCond.MoveNext
'                Loop
'            End If
'        End If
'    End If
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, "MortProdIncomeStatus:PopulateSelectedItems - " + Err.DESCRIPTION
Dim colLine As Collection
Dim clsSwapExtra As SwapExtra
Dim clsTableAccess As TableAccess

Dim rs As ADODB.Recordset

    On Error GoTo Failed
    
    Set clsTableAccess = m_clsMortProdIncomeStatus

    frmProductDetails.MSGSwapProdCond.ClearSecond
' TW 19/03/2007 EP2_1830
    If m_colProductKeys Is Nothing And Not m_bIsEdit Then
        Exit Sub
    End If
' TW 19/03/2007 EP2_1830 End
    Set clsTableAccess = m_clsMortProdIncomeStatus
    m_clsMortProdIncomeStatus.GetLinkedIncomeStatusSettings m_colProductKeys(1), m_colProductKeys(2)
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra

            clsSwapExtra.SetValueID rs.fields("VALUEID").Value
            colLine.Add rs.fields("VALUENAME").Value
            SwapList.AddLineSecond colLine, clsSwapExtra

            clsTableAccess.MoveNext
        Loop
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdProdClass:PopulateSelectedItems - " + Err.DESCRIPTION


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called externally, doesn't do anything for this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdIncomeStatus:SetScreenFields - " + Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
Dim nThisItem As Integer
Dim sValue As String
Dim clsSwapExtra As New SwapExtra
Dim colKeys As Collection
Dim colValues As Collection
Dim clsTableAccess As TableAccess
Dim clsMortProdIncomeStatus As MortProdIncomeStatusTable
    
    On Error GoTo Failed
    Set colKeys = New Collection
    
    colKeys.Add "MORTGAGEPRODUCTCODE"
    colKeys.Add "STARTDATE"
    colKeys.Add "INCOMESTATUS"

    Set clsMortProdIncomeStatus = New MortProdIncomeStatusTable
    
    Set clsTableAccess = clsMortProdIncomeStatus
    clsTableAccess.SetKeyMatchFields colKeys
    
    For nThisItem = 1 To frmProductDetails.MSGSwapIncomeStatus.GetFirstCount
        Set colValues = frmProductDetails.MSGSwapIncomeStatus.GetLineFirst(nThisItem, clsSwapExtra)
        sValue = clsSwapExtra.GetValueID

        Set colValues = New Collection
        colValues.Add m_sProductCode
        colValues.Add m_sProductStartDate
        colValues.Add sValue
        clsTableAccess.SetKeyMatchValues colValues
        clsTableAccess.GetTableData POPULATE_KEYS
        If clsTableAccess.RecordCount > 0 Then
            clsTableAccess.DeleteRow colValues
            clsTableAccess.Update
        End If
            
    Next
    
    clsTableAccess.SetRecordSet Nothing
    
    For nThisItem = 1 To frmProductDetails.MSGSwapIncomeStatus.GetSecondCount()
        Set colValues = frmProductDetails.MSGSwapIncomeStatus.GetLineSecond(nThisItem, clsSwapExtra)
        sValue = clsSwapExtra.GetValueID

        Set colValues = New Collection
        colValues.Add m_sProductCode
        colValues.Add m_sProductStartDate
        colValues.Add sValue
        clsTableAccess.SetKeyMatchValues colValues
        clsTableAccess.GetTableData POPULATE_KEYS
        If clsTableAccess.RecordCount = 0 Then
            clsTableAccess.AddRow
            clsMortProdIncomeStatus.SetMortgageProductCode m_sProductCode
            clsMortProdIncomeStatus.SetStartDate m_sProductStartDate
            clsMortProdIncomeStatus.SetIncomeStatus sValue
            clsTableAccess.Update
        End If
            
    Next
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
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdIncomeStatus.SetProductKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)

    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


