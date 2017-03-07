VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdChannelElig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdChannelElig
' Description   : Provides functionality for the Channel Eligibility tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port, and tidy up
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
' Constants
Private Const ORGANISATION_ID_FIELD = 1

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_clsChannels As DistributionChannelTable
Private m_colProductKeys As Collection
Private m_clsMortProdChannel As MortProdChannelEligTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    Set m_clsChannels = New DistributionChannelTable
    Set m_clsMortProdChannel = New MortProdChannelEligTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Channel
'                 Eligibility, including populating the swaplists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    m_bIsEdit = bIsEdit

    PopulateDistributionChannels
    PopulateSelectedItems
    
    PopulateAvailableItems frmProductDetails.MSGSwapChannel
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateDistributionChannels
' Description   : Populates a list of Distribution Channels
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateDistributionChannels()
    On Error GoTo Failed
    
    TableAccess(m_clsChannels).GetTableData POPULATE_ALL
    TableAccess(m_clsChannels).ValidateData
    
    If TableAccess(m_clsChannels).RecordCount = 0 Then
        g_clsErrorHandling.RaiseError errRecordNotFound
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortProducChannelElig.PopulateDistributionChannels " & Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAvailableItems
' Description   : Populates a list of items available to move from the first swaplist listview
'                 to the second.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList)
    On Error GoTo Failed
    Dim sValue As String
    Dim sValueID As String
    Dim rs As ADODB.Recordset
    Dim colLine As Collection
    Dim clsSwapExtra As SwapExtra
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsChannels
    
    PopulateDistributionChannels
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet
        
        clsTableAccess.MoveFirst
            
        Do While Not rs.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sValue = m_clsChannels.GetChannelName()
            sValueID = m_clsChannels.GetChannelID
            clsSwapExtra.SetValueID sValueID
            
            ' Does this value exist in the Selected Items Listbox?
            If DoesSwapValueExist(sValue) = False Then
                colLine.Add sValue
                SwapList.AddLineFirst colLine, clsSwapExtra
            End If
            clsTableAccess.MoveNext
        Loop
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Channel Eligibility - No Distribution Channels avaailble"
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesSwapValueExist
' Description   : Checks if the value passed in sValue exists in the second listview of the
'                 swaplist MSGSWapChannel.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesSwapValueExist(sValue As String) As Boolean
    Dim bFound As Boolean
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim sSecondValue As String
    Dim colValue As Collection
    
    nCount = frmProductDetails.MSGSwapChannel.GetSecondCount()

    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        Set colValue = frmProductDetails.MSGSwapChannel.GetLineSecond(nThisItem)
        
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
    Dim clsTableAccess As TableAccess
    Dim rsMortChannelElig As ADODB.Recordset
    Dim rsDistChannels As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colValues As New Collection
    Dim enumPopulate As PopulateType
    
    enumPopulate = POPULATE_EMPTY

    Set clsTableAccess = m_clsMortProdChannel
    
    If Not m_colProductKeys Is Nothing Then
        clsTableAccess.SetKeyMatchValues m_colProductKeys
        enumPopulate = POPULATE_KEYS
    End If
    
    Set rsMortChannelElig = clsTableAccess.GetTableData(enumPopulate)
    
    clsTableAccess.ValidateData
                    
    If clsTableAccess.RecordCount > 0 Then
        rsMortChannelElig.MoveFirst
        
        Do While Not rsMortChannelElig.EOF
            sValueID = m_clsMortProdChannel.GetChannel()
            
            If Len(sValueID) > 0 Then
                Set colValues = New Collection
                colValues.Add sValueID
                
                TableAccess(m_clsChannels).SetKeyMatchValues colValues
                
                Set rsDistChannels = TableAccess(m_clsChannels).GetTableData()
                If rsDistChannels.RecordCount = 1 Then
                    sValue = m_clsChannels.GetChannelName()
                Else
                    sValue = ""
                    g_clsErrorHandling.RaiseError errGeneralError, "Distribution Channel " + sValueID + " does not exist"
                End If
                
                If Len(sValue) > 0 Then
                    Set colLine = New Collection
                    Set clsSwapExtra = New SwapExtra
                    colLine.Add sValue
                    
                    clsSwapExtra.SetValueID sValueID
                    frmProductDetails.MSGSwapChannel.AddLineSecond colLine, clsSwapExtra
                End If
            End If
            rsMortChannelElig.MoveNext
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
' Function      : SaveChannelEligiblity
' Description   : Saves all selected items on the Channel Eligibility swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChannelEligiblity()
    On Error GoTo Failed
    Dim nThisItem As Integer
    Dim nSelectedCount As Integer
    Dim sValue As String
    Dim sValueID As String

    Dim clsSwapExtra As New SwapExtra
    Dim colValues As New Collection
    
    TableAccess(m_clsMortProdChannel).DeleteAllRows
    nSelectedCount = frmProductDetails.MSGSwapChannel.GetSecondCount()

    For nThisItem = 1 To nSelectedCount
        Set colValues = frmProductDetails.MSGSwapChannel.GetLineSecond(nThisItem, clsSwapExtra)
        
        If colValues.Count > 0 Then
            sValue = colValues(1)
        End If
        
        If Not clsSwapExtra Is Nothing Then
            sValueID = clsSwapExtra.GetValueID()
        End If
    
        If Len(sValueID) > 0 Then
            TableAccess(m_clsMortProdChannel).AddRow
        
            m_clsMortProdChannel.SetProductCode m_sProductCode
            m_clsMortProdChannel.SetStartDate m_sProductStartDate
            m_clsMortProdChannel.SetChannel sValueID
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "MortProdPurposeOfLoan:SaveScreenData - ValueID is empty"
        End If
    Next
        
    TableAccess(m_clsMortProdChannel).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all selected items on the Channel Eligibility swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    
    SaveChannelEligiblity
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally when this tab is switched to.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    Set frmLenderDetails.lvLegalFees.SelectedItem = Nothing
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Channel Eligibility tab - called externaly,
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
        g_clsErrorHandling.RaiseError errKeysEmpty, "AdditionalMortProdParams.SetProductKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
