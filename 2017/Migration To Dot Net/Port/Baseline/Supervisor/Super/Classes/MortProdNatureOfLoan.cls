VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdNatureOfLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdNatureOfLoan
' Description   : Support class for the Mortgage Product Nature Of Loan tab on the Mortgage Product screen
' Change history
' Prog      Date        Description
' TW        30/11/2006  EP2_253 - Created
' TW        19/03/2007  EP2_1830 - When adding Mortgage Product in Supervisor error message displayed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_sIncomeStatus As String
Private m_colProductKeys As Collection
Private m_clsMortProdNatureOfLoan As MortProdNatureOfLoanTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsMortProdNatureOfLoan = New MortProdNatureOfLoanTable
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
    PopulateSelectedItems frmProductDetails.MSGSwapNatureOfLoan
    PopulateAvailableItems frmProductDetails.MSGSwapNatureOfLoan

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
Dim colLine As Collection
Dim clsSwapExtra As SwapExtra
Dim clsTableAccess As TableAccess

Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    SwapList.ClearFirst
    Set clsTableAccess = m_clsMortProdNatureOfLoan
' TW 19/03/2007 EP2_1830
'    m_clsMortProdNatureOfLoan.GetAvailableNatureOfLoanSettings m_colProductKeys(1), m_colProductKeys(2)
    If m_bIsEdit Then
        m_clsMortProdNatureOfLoan.GetAvailableNatureOfLoanSettings m_colProductKeys(1), m_colProductKeys(2)
    Else
        m_clsMortProdNatureOfLoan.GetAvailableNatureOfLoanSettings "", ""
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
Private Sub PopulateSelectedItems(SwapList As MSGHorizontalSwapList)
Dim colLine As Collection
Dim clsSwapExtra As SwapExtra
Dim clsTableAccess As TableAccess

Dim rs As ADODB.Recordset

    On Error GoTo Failed
    
    Set clsTableAccess = m_clsMortProdNatureOfLoan

    frmProductDetails.MSGSwapProdCond.ClearSecond
' TW 19/03/2007 EP2_1830
    If m_colProductKeys Is Nothing And Not m_bIsEdit Then
        Exit Sub
    End If
' TW 19/03/2007 EP2_1830 End
    Set clsTableAccess = m_clsMortProdNatureOfLoan
    m_clsMortProdNatureOfLoan.GetLinkedNatureOfLoanSettings m_colProductKeys(1), m_colProductKeys(2)
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
    g_clsErrorHandling.RaiseError Err.Number, "MortProdNatureOfLoan:SetScreenFields - " + Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
Dim nThisItem As Integer
Dim sValue As String
Dim clsSwapExtra As New SwapExtra
Dim colKeys As Collection
Dim colValues As Collection
Dim clsTableAccess As TableAccess
Dim clsMortProdNatureOfLoan As MortProdNatureOfLoanTable
    
    On Error GoTo Failed
    Set colKeys = New Collection
    
    colKeys.Add "MORTGAGEPRODUCTCODE"
    colKeys.Add "STARTDATE"
    colKeys.Add "NATUREOFLOAN"
    
    Set clsMortProdNatureOfLoan = New MortProdNatureOfLoanTable
    
    Set clsTableAccess = clsMortProdNatureOfLoan
    clsTableAccess.SetKeyMatchFields colKeys

    For nThisItem = 1 To frmProductDetails.MSGSwapNatureOfLoan.GetFirstCount
        Set colValues = frmProductDetails.MSGSwapNatureOfLoan.GetLineFirst(nThisItem, clsSwapExtra)
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
    
    For nThisItem = 1 To frmProductDetails.MSGSwapNatureOfLoan.GetSecondCount()
        Set colValues = frmProductDetails.MSGSwapNatureOfLoan.GetLineSecond(nThisItem, clsSwapExtra)
        sValue = clsSwapExtra.GetValueID

        Set colValues = New Collection
        colValues.Add m_sProductCode
        colValues.Add m_sProductStartDate
        colValues.Add sValue
        clsTableAccess.SetKeyMatchValues colValues
        clsTableAccess.GetTableData POPULATE_KEYS
        If clsTableAccess.RecordCount = 0 Then
            clsTableAccess.AddRow
            clsMortProdNatureOfLoan.SetMortgageProductCode m_sProductCode
            clsMortProdNatureOfLoan.SetStartDate m_sProductStartDate
            clsMortProdNatureOfLoan.SetNatureOfLoan sValue
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
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdNatureOfLoan.SetProductKeys"
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



