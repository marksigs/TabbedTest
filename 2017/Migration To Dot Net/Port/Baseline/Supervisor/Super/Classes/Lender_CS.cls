VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderCS
' Description   : Client specific version of the Lender class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    SYS2831 Created for client variants
' DJP       13/11/01    SYS2831 Add ability to override Postcode validation
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       14/02/02    SYS4094 Allow Lender to use Base Rate Structure. Added SelectItem and
'                       DeSelectItem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Base class
Implements Lender

' Private data
Private m_bIsEdit As Boolean
Private m_clsLender As Lender
Private m_colTabs As Collection
Private m_colKeys As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_Show
' Description   : Shows the Lender form. First, must Load the form, then show it. Any client specific
'                 code goes in here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_Show()
    On Error GoTo Failed
    
    ' Perform Lender load functionality first
    Lender_Load
    
    ' Then show the form
    m_clsLender.Show
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_Load
' Description   : Loads the Lender form - first perform any Core Load functionality, then create
'                 all tabs (related tab control classes)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_Load()
    On Error GoTo Failed
    
    ' Delete to Core method
    m_clsLender.Load
    
    ' Create tab control classes
    Lender_CreateTabs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_CreateTabs
' Description   : Creates all tab control classes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_CreateTabs()
    m_bIsEdit = m_clsLender.IsEdit
    
    ' Delete to Core method
    m_clsLender.CreateTabs
    
    ' Set the tabs collection for this class to use, now that Core has created all required tabs
    Set m_colTabs = m_clsLender.GetTabsCollection
    
    ' ... Create any client specific tabs here ...
    
    If m_bIsEdit = True Then
        Lender_InitialiseEdit
        ' Initialise any client specific editing
    Else
        ' Need to be able to pass the lender code across to the LegalFee screen so it can be used
        ' to a) display the lender code on the screen in an edit, and also when the user presses
        ' ok we need to use the lender code the user had entered as a key for the LegalFeeSet
        ' table.
        
        Lender_InitialiseAdd
        ' Initialise any client specific Adding
    
    End If
    
    ' Now initialise all tabs (this will include any client tabs added)
    Lender_InitialiseTabs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_InitialiseEdit
' Description   : Performs any Edit processing for the Lender form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_InitialiseEdit()
    On Error GoTo Failed
        
    ' Delegate to the Core method
    m_clsLender.Initialise
    
    ' Get the Core Lender keys for use with client versions
    Set m_colKeys = Lender_GetKeyMatchValues
    
    ' ... Do any client specific processing here ...
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_InitialiseAdd
' Description   : Performs any Add processing for the Lender form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_InitialiseAdd()
    On Error GoTo Failed
    
    ' Delegate to the Core method
    m_clsLender.Initialise

    ' ... Do any client specific processing here ...
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_InitialiseTabs
' Description   : Performs all tab initialisation code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_InitialiseTabs()
    On Error GoTo Failed
    
    ' Delegate to the Core method
    m_clsLender.InitialiseTabs
    
    ' Populate all screen fields from the database
    If m_bIsEdit Then
        Lender_PopulateScreenFields
    End If
    
    ' ... Put any client specific code here ...
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_SaveScreenData
' Description   : Saves all data on the Lender form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_SaveScreenData()
    On Error GoTo Failed
    
    ' First, delegate to the Core method
    m_clsLender.SaveScreenData

    ' Now do any client saves

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Lender_OK
' Description   : Called when the user presses OK on the Lender form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Lender_OK()
    On Error GoTo Failed
    Dim bSuccess As Boolean
    ' First, delegate to Core for Validation
    
    ' Delegate to the Core Validation method first.
    bSuccess = m_clsLender.ValidateScreenData
    
    If bSuccess Then
        ' Delegate to Core to save existing data
        m_clsLender.SaveScreenData
        
        ' ...Do any client specific data saving here...
        m_clsLender.OK
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Lender_Add(nTab As LenderDetailsTabs)
    m_clsLender.Add nTab
End Sub
Private Sub Lender_Cancel()
    m_clsLender.Cancel
End Sub
Private Sub Lender_Delete(nTab As LenderDetailsTabs)
    m_clsLender.Delete nTab
End Sub
Private Sub Lender_Edit(nTab As LenderDetailsTabs)
    m_clsLender.Edit nTab
End Sub
Private Function Lender_GetReturnCode() As MSGReturnCode
    Lender_GetReturnCode = m_clsLender.GetReturnCode
End Function
Private Function Lender_GetTabsCollection() As Collection
    Set Lender_GetTabsCollection = m_clsLender.GetTabsCollection
End Function
Private Sub Lender_HandlePostcode(ByVal nIndex As Integer, nKeyAscii As Integer)
    m_clsLender.HandlePostcode nIndex, nKeyAscii
End Sub

Private Sub Lender_InitialiseTab(nTab As Variant)
    m_clsLender.InitialiseTab nTab
End Sub
Private Function Lender_IsEdit() As Boolean
    Lender_IsEdit = m_clsLender.IsEdit
End Function
Private Sub Lender_PopulateScreenFields()
    m_clsLender.PopulateScreenFields
End Sub
Private Sub Lender_SetActiveTab(Optional nTab As LenderDetailsTabs, Optional bSetNow As Boolean = True)
    m_clsLender.SetActiveTab nTab, bSetNow
End Sub
Private Sub Lender_SetButtonsState(nTab As LenderDetailsTabs)
    m_clsLender.SetButtonsState nTab
End Sub
Private Sub Lender_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsLender.SetIsEdit bEdit
End Sub
Private Sub Lender_SetKeys(colKeys As Collection)
    m_clsLender.SetKeys colKeys
End Sub
Private Sub Lender_UnloadForm()
    m_clsLender.UnloadForm
End Sub
Private Sub Lender_ValidateContactDetailsKey(ByVal nIndex As Integer, nKeyAscii As Integer)
    m_clsLender.ValidateContactDetailsKey nIndex, nKeyAscii
End Sub
Private Sub Lender_ValidatePostcodeKey(ByVal nIndex As Integer, nKeyAscii As Integer)
    m_clsLender.ValidatePostcodeKey nIndex, nKeyAscii
End Sub
Private Sub Class_Initialize()
    Set m_clsLender = New Lender
    frmLenderDetails.SetLender Me
End Sub
Private Sub Class_Terminate()
    Set m_clsLender = Nothing
End Sub
Private Function Lender_GetKeyMatchValues() As Variant
    Set Lender_GetKeyMatchValues = m_clsLender.GetKeyMatchValues
End Function
Private Function Lender_ValidateScreenData() As Boolean
    Lender_ValidateScreenData = m_clsLender.ValidateScreenData
End Function
Private Sub Lender_Initialise()
    m_clsLender.Initialise
End Sub
Private Sub Lender_DeselectItem(nItem As Long)
    m_clsLender.DeselectItem nItem
End Sub
Private Sub Lender_SelectItem(nItem As Long)
    m_clsLender.SelectItem nItem
End Sub

