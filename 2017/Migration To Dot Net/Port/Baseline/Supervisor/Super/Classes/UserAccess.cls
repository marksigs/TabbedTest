VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : UserAccess
' Description   : Handle all controls relevant to the User Access tab.
' History       :
'
' Prog      Date        Description
' STB       21/01/02    SYS2957 Security tab-handler added. Currently, it can
'                       only handle objects with the swap list, operations are
'                       not supported.
' SDS       12/02/02    SYS4033 To show the Available Items in SwapList Control
'                       in the same order that they appear in Supervisor.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mars Specific History:
' Prog   Date        AQR        Decription
' RF     18/01/2006  MAR1000    Added extra error handling
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The column in the swap list with the object id.
Private Const IDX_OBJECT_ID As Long = 1

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The current record's user ID.
Private m_sUserID As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if the form is in an add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the add/edit state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bEdit As Boolean)
    m_bIsEdit = bEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated accoding to the
'                 underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
            
    'Store the add/edit mode.
    m_bIsEdit = bIsEdit
            
    'Setup some column headers.
    SetColumnHeaders
    
    'Set the swaplist contents regardless of add/edit mode.
    PopulateScreenControls
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetColumnHeaders
' Description   : Build the column headers for the swaplist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetColumnHeaders()

    Dim colHeaders As New Collection
    Dim clsHeader As listViewAccess
        
    clsHeader.nWidth = 0
    clsHeader.sName = "ObjectID"
    colHeaders.Add clsHeader
    
    clsHeader.nWidth = 88
    clsHeader.sName = "Supervisor Item"
    colHeaders.Add clsHeader
    
    frmEditUser.swpSecurityResources.SetFirstColumnHeaders colHeaders
    frmEditUser.swpSecurityResources.SetSecondColumnHeaders colHeaders

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Loads data into the swaplists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()

    Dim colLine As Collection
    Dim iManagement As ISecurityManager
    Dim clsResource As SecurityResource
    Dim colResourceList As SecurityResourceList
        
    On Error GoTo Failed
        
    'Load a security object, and populate the swaplists.
    Set iManagement = g_clsSecurityMgr
    
    'Only populate what the user can do if we're editing (otherwise a UserID
    'won't exist yet).
    If m_bIsEdit = True Then
        'Get a resource list which the user can access.
        Set colResourceList = iManagement.GetAllocatedResources(m_sUserID, AgentTypeUser)
        
        'Iterate through each resource and add it to the second list.
        For Each clsResource In colResourceList
            Set colLine = New Collection
            colLine.Add clsResource.ObjectID
            colLine.Add clsResource.ObjectName
            
            'Add the granted resource to the selected list.
            frmEditUser.swpSecurityResources.AddLineSecond colLine
        Next clsResource
    End If
    
    'Get a resource list for every available resource.
    Set colResourceList = iManagement.GetAvailableResources()
    
    'Iterate through each resource and add it to the first list (if it isn't
    'already in the second list).
    For Each clsResource In colResourceList
        If frmEditUser.swpSecurityResources.DoesValueExist(clsResource.ObjectID) = False Then
            Set colLine = New Collection
            colLine.Add clsResource.ObjectID
            colLine.Add clsResource.ObjectName
        
            frmEditUser.swpSecurityResources.AddLineFirst colLine
        End If
    Next clsResource
    
    frmEditUser.swpSecurityResources.SetFirstSorted False
    frmEditUser.swpSecurityResources.SetSecondSorted False
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError errGeneralError, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Loads all details from the tables into the screens controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetScreenFields() As Boolean
    SetScreenFields = True  'Stub.
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all details from the relevant screen controls to table
'                 objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
On Error GoTo Failed

    Dim iIndex As Integer
    Dim colLine As Object
    Dim iManagement As ISecurityManager
    Dim strStatus As String
    
    strStatus = "Initialising security component"
        
    'Cast a management interface onto the security component.
    Set iManagement = g_clsSecurityMgr
    
    strStatus = "Removing existing access"
    
    'Remove all access for the current user.
    iManagement.RemoveAllAccess m_sUserID, AgentTypeUser
    
    strStatus = "Adding selected access"
    
    'Re-add everything which has been selected.
    For iIndex = 1 To frmEditUser.swpSecurityResources.GetSecondCount
        Set colLine = frmEditUser.swpSecurityResources.GetLineSecond(iIndex)
        
        strStatus = "Granting access to item " & CStr(iIndex)
        
        'Grant access to this resource.
        iManagement.SetAccess m_sUserID, AgentTypeUser, True, colLine(IDX_OBJECT_ID), ""
    Next iIndex
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError _
        Err.Number, _
        "Error in UserAccess SaveScreenData (status is " & strStatus & "): " & Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Update all underlying table objects handled by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    'Stub.
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetUserID
' Description   : Sets the user ID at module level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetUserID(ByVal sUserID As String)
    m_sUserID = sUserID
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates any control values relevant to this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Configure the screen controls to work in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Configure the screen controls to work in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    'Stub.
End Sub

