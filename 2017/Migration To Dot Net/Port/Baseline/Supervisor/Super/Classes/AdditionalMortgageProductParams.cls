VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdditionalMortProdParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : AdditionalMortProdParams
' Description   : Contains all processing for the Additional Product Parameters tab on the Product
'                 Details tab.
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants, added InitTab.
' HMA       09/12/04    BMIDS959  Remove MortgageProductBands table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


' Text control constants
Private Const COMPULSORY_BC_TERM = 0
Private Const COMPULSORY_PP_TERM = 1

' BMIDS959  remove unused constants.

Private Const DIFFERENTIAL_LOWER_LIMIT = 12
Private Const DIFFERENTIAL_UPPER_LIMIT = 13

' Privata data
Private m_bIsEdit As Boolean
Private m_colProductKeys As Collection
Private m_clsProductParameters As MortProdParamsTable
' BMIDS959  Remove MortgageProductBands table

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    If m_clsProductParameters Is Nothing Then
        Set m_clsProductParameters = New MortProdParamsTable
    End If
    
    Set m_colProductKeys = New Collection
    
    ' BMIDS959  Remove MortgageProductBands table
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form to do initialisation of the Additional
'                 Product Details tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed

    m_bIsEdit = bIsEdit
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Sets the status of the Mortgage Product Additional Parameters tab when adding a
'                 product. Need to create the relevant records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    On Error GoTo Failed
    
    ' BMIDS959  Remove MortgageProductBands
    
    g_clsFormProcessing.CreateNewRecord m_clsProductParameters
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Sets the status of the Mortgage Product Additional Parameters tab when editing
'                 a product. Need to read the relevant records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    On Error GoTo Failed
    
    ' Mortgage Product Parameters
    TableAccess(m_clsProductParameters).SetKeyMatchValues m_colProductKeys
    TableAccess(m_clsProductParameters).GetTableData
    
    ' BMIDS959 Remove Mortgage Product Bands
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the Additional Mortgage Parameters tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCompulsoryBC(OPT_YES), frmProductDetails.optCompulsoryBC(OPT_NO), m_clsProductParameters.GetCompulsoryBC(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCompulsoryPP(OPT_YES), frmProductDetails.optCompulsoryPP(OPT_NO), m_clsProductParameters.GetCompulsoryPP(), SET_CONTROL_VALUE
    
    frmProductDetails.txtAdditionalParams(COMPULSORY_BC_TERM).Text = m_clsProductParameters.GetCompulsoryBCTerm()
    frmProductDetails.txtAdditionalParams(COMPULSORY_PP_TERM).Text = m_clsProductParameters.GetCompulsoryPPTerm()

    ' BMIDS959  Remove unused fields
    
    frmProductDetails.txtAdditionalParams(DIFFERENTIAL_LOWER_LIMIT).Text = m_clsProductParameters.GetDifferentialLowerLimit()
    frmProductDetails.txtAdditionalParams(DIFFERENTIAL_UPPER_LIMIT).Text = m_clsProductParameters.GetDifferentialUpperLimit()
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all fields on the Additional Mortgage Parameters tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    Dim vTmp As Variant
    
    ' Save all key values
    SaveKeys
    
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCompulsoryBC(OPT_YES), frmProductDetails.optCompulsoryBC(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsProductParameters.SetCompulsoryBC vTmp
    
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCompulsoryPP(OPT_YES), frmProductDetails.optCompulsoryPP(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsProductParameters.SetCompulsoryPP vTmp
    
    m_clsProductParameters.SetCompulsoryBCTerm frmProductDetails.txtAdditionalParams(COMPULSORY_BC_TERM).Text
    m_clsProductParameters.SetCompulsoryPPTerm frmProductDetails.txtAdditionalParams(COMPULSORY_PP_TERM).Text

    m_clsProductParameters.SetDifferentialLowerLimit frmProductDetails.txtAdditionalParams(DIFFERENTIAL_LOWER_LIMIT).Text
    m_clsProductParameters.SetDifferentialUpperLimit frmProductDetails.txtAdditionalParams(DIFFERENTIAL_UPPER_LIMIT).Text
    
    ' BMIDS959  Remove MortgageProductBands

    DoUpdates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveKeys
' Description   : Saves all key values onto the tables associated with this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveKeys()
    On Error GoTo Failed
    Dim sCode As String
    Dim sDate As String

    ' Read the Product Code and Start Date
    sCode = m_colProductKeys(PRODUCT_CODE_KEY)
    sDate = m_colProductKeys(PRODUCT_START_DATE_KEY)
    
    ' Set the key values
    ' BMIDS959  Remove MortgageProductBands
    m_clsProductParameters.SetStartDate sDate
    m_clsProductParameters.SetMortgageProductCode sCode

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
        g_clsErrorHandling.RaiseError errKeysEmpty, "AdditionalMortProdParams.SetMortgageParameterKeys"
    End If
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' BMIDS959  Remove ValidateEarlyRedemption function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Callled to validate all data on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    ' BMIDS959  Remove call to ValidateEarlyRedemption
    bRet = True
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    MsgBox "Additional MortgageProduct:ValidateScreenData: Error is " + Err.DESCRIPTION
    ValidateScreenData = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Called to update all tables used by this tab, once all data has been written
'                 to the tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsProductParameters).Update
    
    ' BMIDS959  Remove MortgageProductBands
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called whenever this tab is switched to, so perform any initialisation required
'                 by the tab when it gets focus.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

