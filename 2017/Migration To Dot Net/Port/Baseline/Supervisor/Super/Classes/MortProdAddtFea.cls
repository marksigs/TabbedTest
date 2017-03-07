VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdAddtFea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************
'   File            : MortProdAddtFea.cls
'   Description     : Facilitates to manage additional features list for
'                     a mortgage product
'   Change history
'
'   Prog      Date        Description
'
'   MC      25/05/2004      REG018/BMIDS775 File added to implement REG018 change
'
'
'*************************************************************************************

Option Explicit

' Private data
Private m_bIsEdit           As Boolean
Private m_sProductCode      As String
Private m_sProductStartDate As String
Private m_colProductKeys    As Collection

Private m_clsAddlFeatTable  As MortProdAddtFeaTable
'Private m_clsComboTable    As New ComboValueTable
Private m_clsComboTable     As ComboValueTable
' Constants
Private Const m_sCOMBOADDITIONALFEATURES As String = "ADDITIONALFEATURES"

'*************************************************************************************
'   Function      : Class_Initialize
'   Description   : Called when this class is instantiated. Create module data.
'
'*************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrInitialize
    '*=[MC]Initialize Mortgage Product Additional Feature Table Object
    Set m_clsAddlFeatTable = New MortProdAddtFeaTable
    Set m_clsComboTable = New ComboValueTable
ExitInitialize:
    '[MC]Cleanup code here
    Exit Sub
    
ErrInitialize:
    '*=[MC]Handle Errors or Throw
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'*************************************************************************************
' Function      :   SetProductKeys()
' Arguments     :   DB Column keys collection
' Description   :   Called externally to set the keys used to identify rows for this
'                   table.
'*************************************************************************************
Public Sub SetProductKeys(colValues As Collection)
    
    On Error GoTo ErrSetProductKeys
    
    '*=[MC]If Column key collection is blank/null, throw an error
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdAddtFea.SetProductCode"
    End If
    
    '*=[MC]Persist ProductCode and Start Date for further use
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    '*=Persist keyValue collection for further use
    Set m_colProductKeys = colValues
ExitSetProductKeys:
    '[MC]Cleanup code here
    
    Exit Sub
ErrSetProductKeys:
    '*=[MC]Handle Errors or Throw
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'*************************************************************************************
' Function      : Initialise
' Description   : Called to initialise the Product Eligibility tab. Will load all data
'                 required by this tab.
'*************************************************************************************

Public Sub Initialise(bIsEdit As Boolean)
    
    On Error GoTo ErrorHandler
    
    m_bIsEdit = bIsEdit

    '*=[MC]Initialize Available and Selected Swap Lists
    
    If m_bIsEdit Then
        '*=Selected Swap list
        PopulateSelectedItems frmAdditionalFeatures.MSGSwapAddlItems
    End If
    '*=Available Swap list
    PopulateAvailableItems frmAdditionalFeatures.MSGSwapAddlItems, m_sCOMBOADDITIONALFEATURES
    
    


ExitHandler:
    '*=[MC]CLEANUP CODE HERE
    
    Exit Sub
ErrorHandler:
    '*=[MC]HANDLE ERRORS OR THROW
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'*************************************************************************************
' Function      : PopulateComboTable
' Description   : Reads all combovalues from the combovalue table for the
'                 specified group.
'*************************************************************************************

Private Sub PopulateComboTable(sGroup As String)
    
    Dim colKeys As Collection
    
    On Error GoTo ErrorHandler
    
    '*[MC]Initialize ComboValueTable object
    Set m_clsComboTable = New ComboValueTable
    '*=[MC]If ComboGroupName is specified populate combo data
    Set colKeys = New Collection
    
    If Len(sGroup) > 0 Then
        '*=Add ComboGroup Name to collection
        colKeys.Add sGroup
        '*=Set Matching Keys
        TableAccess(m_clsComboTable).SetKeyMatchValues colKeys
        '*=[MC]Get ComboData
        TableAccess(m_clsComboTable).GetTableData
        '*=[MC]Validate populated Combo Data
        TableAccess(m_clsComboTable).ValidateData
        '*=[MC]Combo do not have any list items, throw an error
        If TableAccess(m_clsComboTable).RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, " No entries for " + sGroup + " in Combo table"
        End If
    End If

ExitHandler:
    '[MC]CleanUp Code here
    Set colKeys = Nothing
    Exit Sub
ErrorHandler:
    '*=[MC]Handle Errors
    
    g_clsErrorHandling.RaiseError Err.Number, "MortProdPurpOfLoan:PopulateComboTable - " + Err.DESCRIPTION
End Sub

'*************************************************************************************
' Function      : PopulateAvailableItems
' Description   : Populates all possible available items in the first swaplist box.
'                 Needs to read this list of available items, then only populate
'                 those that don't exist in the second swaplist.
'*************************************************************************************
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList, sComboGroup As String)
    
    Dim RecData             As Recordset
    Dim colLine             As Collection
    Dim clsSwapExtra        As SwapExtra
    Dim clsTableAccess      As TableAccess
    Dim sValue              As String
    Dim sValueId            As String
    
    On Error GoTo ErrorHandler
    
    '*=[MC]Clear Available swaplist items
    SwapList.ClearFirst
    
    '*=[MC]Populate List
    PopulateComboTable sComboGroup
    '*=TableAccess Reference
    Set clsTableAccess = m_clsComboTable
    
    '*=If Records Exist, Process the request
    If clsTableAccess.RecordCount() > 0 Then
        Set RecData = clsTableAccess.GetRecordSet
        '#########################
        'clsTableAccess.MoveFirst
        RecData.MoveFirst
        '*=Iterate through all record items in the resultset
        Do While Not RecData.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sValue = m_clsComboTable.GetValueName()
            sValueId = m_clsComboTable.GetValueID()
            clsSwapExtra.SetValueID sValueId
            
            ' Does this value exist in the Selected Items Listbox? If it doesn't, then add the item
            If SwapList.DoesValueExist(sValue) = False Then
                colLine.Add sValue
                SwapList.AddLineFirst colLine, clsSwapExtra
            End If
            
            RecData.MoveNext
            '##################
            'clsTableAccess.MoveNext
        Loop
    End If
    
    '*=Close RecordSet
    clsTableAccess.CloseRecordSet

ExitHandler:
    '*=[MC]Cleanup Code here
    Set RecData = Nothing
    Exit Sub
ErrorHandler:

    '*=Handle Errors
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'*************************************************************************************
' Function      : PopulateSelectedItems
' Description   : Populates all selected items in the second swaplist box. Needs to read
'                 this list of items and add them to the second swaplist.
'*************************************************************************************

Private Sub PopulateSelectedItems(SwapList As MSGHorizontalSwapList)
    
    Dim sValue          As String
    Dim colLine         As Collection
    Dim sValueId        As String
    'Dim colValues       As New Collection
    Dim clsSwapExtra    As SwapExtra
    Dim clsTableAccess  As TableAccess
    Dim rsPurposeOfLoan As ADODB.Recordset
    Dim rsAddFea        As Recordset
    Dim enumPopulate    As PopulateType
    Dim colValues      As Collection
    
    On Error GoTo ErrorHandler
    
    '*=Initialize collection
    Set colValues = New Collection
    
    enumPopulate = POPULATE_EMPTY

    '*=Clear List from Selected items list
    SwapList.ClearSecond
    
    Set clsTableAccess = m_clsAddlFeatTable
    
    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulate = POPULATE_KEYS
    End If
    
    Set clsTableAccess = m_clsAddlFeatTable
    
    clsTableAccess.SetKeyMatchValues m_colProductKeys
    
    '*=Get Selected Items list from ADDITIONALFEATURES DB Table
    m_clsAddlFeatTable.GetAdditionalFeatures m_colProductKeys, m_sCOMBOADDITIONALFEATURES
    
    If clsTableAccess.RecordCount > 0 Then
        '*=[MC]    Get Recordset
        Set rsAddFea = clsTableAccess.GetRecordSet
        
        If Not rsAddFea Is Nothing Then
            If rsAddFea.RecordCount > 0 Then
                rsAddFea.MoveFirst
                '*=[MC]Iterate through all selected items and populate selected items
                '      swaplist
                Do While Not rsAddFea.EOF
                    sValueId = m_clsAddlFeatTable.GetFeatureID()
                    If Len(sValueId) > 0 Then
                        sValue = m_clsAddlFeatTable.GetFeatureDesc
                        If Len(sValue) > 0 Then
                            Set colLine = New Collection
                            Set clsSwapExtra = New SwapExtra
                            '*=Add value pair to the collection
                            colLine.Add sValue
                            clsSwapExtra.SetValueID sValueId
                            '*=Add item to the swap list
                            SwapList.AddLineSecond colLine, clsSwapExtra
                        End If
                    End If
                    '*=[MC]Move Recordset pointer
                    rsAddFea.MoveNext
                Loop
            End If
        End If
        '*=[MC]
        clsTableAccess.CloseRecordSet
    End If
ExitHandler:
    '*=[MC]Cleanup Code here
    Set rsAddFea = Nothing
    Exit Sub
    
ErrorHandler:
    '*=[MC]handle errors or throw
    g_clsErrorHandling.RaiseError Err.Number, "MortProdAddFea:PopulateSelectedItems-" + Err.DESCRIPTION
End Sub

''''***********************************************************************************
'''' Function      : SetScreenFields
'''' Description   : Called externally to set screen fields
''''***********************************************************************************
'''Public Sub SetScreenFields()
'''    On Error GoTo Failed
'''
'''    Exit Sub
'''Failed:
'''    g_clsErrorHandling.RaiseError Err.Number, Err.Description
'''End Sub


'***********************************************************************************
' Function      : SaveScreenData
' Description   : Saves all data changes from AdditionalFeatures form
'***********************************************************************************
Public Sub SaveScreenData()
    
    On Error GoTo ErrorHandler
    
    '*=[MC]Save Additional Features
    Call SaveAdditionalFeatures
    
ExitHandler:
    '*=[MC] Cleanup here
    Exit Sub
    
ErrorHandler:
    '*=[MC]Handle Errors or Throw
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'***********************************************************************************
' Function      : SaveEmploymentEligibility
' Description   : Saves all data entered in the swaplist to the database
'***********************************************************************************

Private Sub SaveAdditionalFeatures()
    
    Dim sValueId        As String
    Dim sValue          As String
    Dim nThisItem       As Integer
    Dim nSelectedCount  As Integer
    Dim clsTableAccess  As TableAccess
    Dim clsSwapExtra    As SwapExtra
    Dim colValues       As Collection
    
    On Error GoTo ErrorHandler
    
    '*=[MC]Initialize required objected
    Set clsTableAccess = m_clsAddlFeatTable
    Set colValues = New Collection
    Set clsSwapExtra = New SwapExtra
    
    '*=Set Column Criteria
    clsTableAccess.SetKeyMatchValues m_colProductKeys
        
    '*=Delete Existing Rows before saving the list
    clsTableAccess.GetTableData
    clsTableAccess.ValidateData
    clsTableAccess.DeleteAllRows
    
    '*=Save Selected Items to the database
    nSelectedCount = frmAdditionalFeatures.MSGSwapAddlItems.GetSecondCount()
    '*=Iterate through selected items list and persist to the database
    For nThisItem = 1 To nSelectedCount
        Set colValues = frmAdditionalFeatures.MSGSwapAddlItems.GetLineSecond(nThisItem, clsSwapExtra)
        If colValues.Count > 0 Then
            sValue = colValues(1)
        End If
        If Not clsSwapExtra Is Nothing Then
            sValueId = clsSwapExtra.GetValueID()
        End If
        If Len(sValueId) > 0 Then
            clsTableAccess.AddRow
            '*=Save Data
            m_clsAddlFeatTable.SetProductCode m_sProductCode
            m_clsAddlFeatTable.SetStartDate m_sProductStartDate
            m_clsAddlFeatTable.SetFeatureID sValueId
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
        End If
    Next
    
    '*=Update Changes and close recordset
    clsTableAccess.Update
    clsTableAccess.CloseRecordSet
ExitHandler:
    '*=Cleanup code here
    Set colValues = Nothing
    Set clsSwapExtra = Nothing
    Set clsTableAccess = Nothing
    
    Exit Sub
    
ErrorHandler:
    '*=Handle Errors
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'***************************************************************************************
' Function      : ValidateScreenData
' Description   : No Validation required for additional features screen data
'***************************************************************************************

Public Function ValidateScreenData() As Boolean
    
    On Error GoTo ErrorHandler
    
    ValidateScreenData = True
ExitHandler:
    '*=Cleanup Code here
    Exit Function
ErrorHandler:
    '*=Handle errors here
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''***************************************************************************************
'' Function      : InitTab
'' Description   : Called whenever this tab is switched to, so perform any initialisation
''                 required
''                 by the tab when it gets focus.
''***************************************************************************************
'Public Sub InitTab()
'    On Error GoTo Failed
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
   
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
