VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdEmpElig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdEmpElig
' Description   : Provides functionality for the Employment Eligibility tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port, and tidy up
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_colProductKeys As Collection
Private m_clsEmpElig As MortProdEmpEligTable
Private m_clsComboTable As New ComboValueTable

' Constants
Private Const m_sComboEmpStatus As String = "EmploymentStatus"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is instantiated. Create module data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    Set m_clsEmpElig = New MortProdEmpEligTable
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Called externally to set the keys used to identify rows for this
' table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdEmpElig.SetProductCode"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to initialise the Product Eligibility tab. Will load all data required by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    m_bIsEdit = bIsEdit

    ' Initialise swaplists
    If m_bIsEdit Then
        PopulateSelectedItems frmProductDetails.MSGSwapEmpElig
    End If
    PopulateAvailableItems frmProductDetails.MSGSwapEmpElig, m_sComboEmpStatus

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateComboTable
' Description   : Reads all combovalues from the combovalue table for the specified group.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateComboTable(sGroup As String)
    On Error GoTo Failed
    Dim colKeys As New Collection
    On Error GoTo Failed
    
    Set m_clsComboTable = New ComboValueTable
    
    If Len(sGroup) > 0 Then
        colKeys.Add sGroup
        TableAccess(m_clsComboTable).SetKeyMatchValues colKeys
        TableAccess(m_clsComboTable).GetTableData
        TableAccess(m_clsComboTable).ValidateData
        
        If TableAccess(m_clsComboTable).RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, " No entries for " + sGroup + " in Combo table"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdPurpOfLoan:PopulateComboTable - " + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates all possible available items in the first swaplist box. Needs to read
'                 this list of available items, then only populate those that don't exist in the
'                 second swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList, sComboGroup As String)
    ' Comes from the Combo Table.
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sValue As String
    Dim colLine As Collection
    Dim sValueID As String
    Dim clsSwapExtra As SwapExtra
    Dim clsTableAccess As TableAccess
    
    ' Before we start we must remove all items from the first swaplist
    SwapList.ClearFirst
    
    PopulateComboTable sComboGroup
    
    Set clsTableAccess = m_clsComboTable
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.MoveFirst
            
        ' loop through all items returned and add them to the first swaplist
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sValue = m_clsComboTable.GetValueName()
            sValueID = m_clsComboTable.GetValueID()
            clsSwapExtra.SetValueID sValueID
            
            ' Does this value exist in the Selected Items Listbox? If it doesn't, then add the item
            If SwapList.DoesValueExist(sValue) = False Then
                colLine.Add sValue
                SwapList.AddLineFirst colLine, clsSwapExtra
            End If
            
            clsTableAccess.MoveNext
        Loop
    End If
    
    clsTableAccess.CloseRecordSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateSelectedItems
' Description   : Populates all selected items in the second swaplist box. Needs to read
'                 this list of items and add them to the second swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateSelectedItems(SwapList As MSGHorizontalSwapList)
    On Error GoTo Failed
    Dim sValue As String
    Dim colLine As Collection
    Dim sValueID As String
    Dim colValues As New Collection
    Dim clsSwapExtra As SwapExtra
    Dim clsTableAccess As TableAccess
    Dim rsPurposeOfLoan As ADODB.Recordset
    Dim enumPopulate As PopulateType
    
    enumPopulate = POPULATE_EMPTY

    ' First clear all items from the second swaplist
    frmProductDetails.MSGSwapEmpElig.ClearSecond
    Set clsTableAccess = m_clsEmpElig
    
    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulate = POPULATE_KEYS
    End If
    
    Set clsTableAccess = m_clsEmpElig
    clsTableAccess.SetKeyMatchValues m_colProductKeys
    
    ' Read all items form the database
    m_clsEmpElig.GetEmploymentEligibility m_colProductKeys, m_sComboEmpStatus
    If clsTableAccess.RecordCount > 0 Then
        
        Set rsPurposeOfLoan = clsTableAccess.GetRecordSet
        
        If Not rsPurposeOfLoan Is Nothing Then
            If rsPurposeOfLoan.RecordCount > 0 Then
                rsPurposeOfLoan.MoveFirst
                
                Do While Not rsPurposeOfLoan.EOF
                    sValueID = m_clsEmpElig.GetStatus()
                    
                    If Len(sValueID) > 0 Then
                        sValue = m_clsEmpElig.GetStatusText
                    
                        If Len(sValue) > 0 Then
                            Set colLine = New Collection
                            Set clsSwapExtra = New SwapExtra
                            colLine.Add sValue
                            
                            clsSwapExtra.SetValueID sValueID
                            frmProductDetails.MSGSwapEmpElig.AddLineSecond colLine, clsSwapExtra
                        End If
                    Else
                    
                    End If
                rsPurposeOfLoan.MoveNext
                Loop
            End If
        End If
        
        clsTableAccess.CloseRecordSet
    
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProdPurposeOfLoan:SaveScreenData - " + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called externally to set screen fields - this tab doesn't have any fields at
'                 present so the function is here for completeness.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all data entered on this tab to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    
    ' Save the swaplist data
    SaveEmploymentEligibility
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveEmploymentEligibility
' Description   : Saves all data entered in the swaplist to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveEmploymentEligibility()
    On Error GoTo Failed
    
    Dim sValueID As String
    Dim sValue As String
    Dim colValues As New Collection
    Dim nThisItem As Integer
    Dim clsSwapExtra As New SwapExtra
    Dim nSelectedCount As Integer
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsEmpElig
    clsTableAccess.SetKeyMatchValues m_colProductKeys
        
    ' Delete all items form the database
    clsTableAccess.GetTableData
    clsTableAccess.ValidateData
    clsTableAccess.DeleteAllRows
        
    nSelectedCount = frmProductDetails.MSGSwapEmpElig.GetSecondCount()

    ' Add all new values to the database
    For nThisItem = 1 To nSelectedCount
        Set colValues = frmProductDetails.MSGSwapEmpElig.GetLineSecond(nThisItem, clsSwapExtra)
        
        If colValues.Count > 0 Then
            sValue = colValues(1)
        End If
        
        If Not clsSwapExtra Is Nothing Then
            sValueID = clsSwapExtra.GetValueID()
        End If
    
        If Len(sValueID) > 0 Then
            clsTableAccess.AddRow
        
            m_clsEmpElig.SetProductCode m_sProductCode
            m_clsEmpElig.SetStartDate m_sProductStartDate
            m_clsEmpElig.SetStatus sValueID
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
        End If
    Next
            
    clsTableAccess.Update
    clsTableAccess.CloseRecordSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called externally to ensure all data that has to be entered, has been.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    
    ValidateScreenData = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
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
