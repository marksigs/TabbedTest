VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Intermediary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : Intermediary
' Description  : Helper-class which contains Intermediary routines related to
'                the tree control found on frmMain.
'
' Change history
' Prog      Date        Description
' AA        27/06/2001  Created
' DJP       30/06/01    SLQ Server port
' STB       10/12/01    SYS2550 SQL Server support.
' STB       11/02/02    SYS3581 Made GetIntermediaryTypeAsText() public so it
'                       could be used by IntermediaryTable.cls
' STB       07/03/02    SYS4246 - Altered FormatKey to stop Right-click crash.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'ListView items have their keys prefixed with this value.
Private Const KEY_PREFIX As String = "INT"

'A GUID representing the current treenode context. Its used whilst building
'the intermediaries treeview nodes.
Private m_sParentNodeKey As String

'A reference to the currently selected tree item.
Private m_clsTreeItem As TreeItem

'These enums match the ValueIDs.
Public Enum IntermediaryTypeEnum
    LeadAgentType = 1
    CompanyType
    AdminCentreType
    IndividualType
    PackagerType ' EP15 pct
    BrokerType ' EP15 pct
End Enum


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateIntermediaries
' Description   : Top-level routine used to begin populating the intermediaries
'                 tree nodes. Only the top-level nodes and their immediate
'                 children are added. Other levels are added dynamically.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateIntermediaries(ByRef tv As MSComctlLib.TreeView, ByRef node As MSComctlLib.node)
    
    Dim clsIntermediaryTable As IntermediaryTable
    
    On Error GoTo Failed
    
    'Create an intermediary table.
    Set clsIntermediaryTable = New IntermediaryTable
        
    'Ensure all added nodes are removed apart from the LeadAgent and Individual nodes.
    If node.Text = INTERMEDIARIES Then
         HandleIntermediaryRefresh False
    End If

' EP15 pct
    'Make the type nodes, bold.
    tv.Nodes(BROKER).Bold = True
    tv.Nodes(PACKAGER).Bold = True

    'Load all top-level Packagers.
    clsIntermediaryTable.GetIntermediaries PackagerType

    'Populate all Packagers
    m_sParentNodeKey = PACKAGER
    PopulateRootIntermediary tv, PackagerType, clsIntermediaryTable, node

    'Load all top-level Brokers
    clsIntermediaryTable.GetIntermediaries BrokerType

    'Populate all Brokers
    m_sParentNodeKey = BROKER
    PopulateRootIntermediary tv, BrokerType, clsIntermediaryTable, node


'    'Make the type nodes, bold.
'    tv.Nodes(INDIVIDUAL).Bold = True
'    tv.Nodes(LEADAGENT).Bold = True
'
'    'Load all top-level LeadAgents.
'    clsIntermediaryTable.GetIntermediaries LeadAgentType
'
'    'Populate all LeadAgents
'    m_sParentNodeKey = LEADAGENT
'
'    PopulateRootIntermediary tv, LeadAgentType, clsIntermediaryTable, node
'
'    'Load all top-level Individuals.
'    clsIntermediaryTable.GetIntermediaries IndividualType
'
'    'Populate all Individuals
'    m_sParentNodeKey = INDIVIDUAL
'    PopulateRootIntermediary tv, IndividualType, clsIntermediaryTable, node
' EP15 pct End

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleInteremediaryRefresh
' Description   : Resets the intermediary tree structure. This is done when the
'                 tree is first being constructed and also if the user chooses
'                 refresh from the menu.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleIntermediaryRefresh(Optional ByVal bPopulate As Boolean = True)
    
    Dim IntNode As MSComctlLib.node
    Dim tmpNode As MSComctlLib.node
    
    On Error GoTo Failed
    
    'Release reference to the current item.
    frmMain.ClearCurrentTreeitem

    'Get a reference to the root intermediary node.
    Set IntNode = frmMain.tvwDB.Nodes(INTERMEDIARIES)
    
    If IntNode Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Intermediary Node not found"
' EP15 pct
    Else
        'Remove and re-add the two root nodes.
        On Error Resume Next
        
        Set tmpNode = frmMain.tvwDB.Nodes(PACKAGER)
        
        If Not tmpNode Is Nothing Then
            frmMain.tvwDB.Nodes.Remove (PACKAGER)
            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, PACKAGER, PACKAGER, "closed"
        End If

        Set tmpNode = frmMain.tvwDB.Nodes(BROKER)

        If Not tmpNode Is Nothing Then
            frmMain.tvwDB.Nodes.Remove (BROKER)
            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, BROKER, BROKER, "closed"
        End If

        On Error GoTo Failed
        
        If bPopulate Then
            'Re-populate the initial tree-structure (i.e. all Packagers and Brokers).
            PopulateIntermediaries frmMain.tvwDB, IntNode
        End If
    End If

'    Else
'        'Remove and re-add the two root nodes.
'        On Error Resume Next
'
'        Set tmpNode = frmMain.tvwDB.Nodes(LEADAGENT)
'
'        If Not tmpNode Is Nothing Then
'            frmMain.tvwDB.Nodes.Remove (LEADAGENT)
'            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, LEADAGENT, LEADAGENT, "closed"
'        End If
'
'        Set tmpNode = frmMain.tvwDB.Nodes(INDIVIDUAL)
'
'        If Not tmpNode Is Nothing Then
'            frmMain.tvwDB.Nodes.Remove (INDIVIDUAL)
'            frmMain.tvwDB.Nodes.Add INTERMEDIARIES, tvwChild, INDIVIDUAL, INDIVIDUAL, "closed"
'        End If
'
'        On Error GoTo Failed
'
'        If bPopulate Then
'            'Re-populate the initial tree-structure (i.e. all individuals and lead agents).
'            PopulateIntermediaries frmMain.tvwDB, IntNode
'        End If
'    End If
' EP15 pct End

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateRootIntermediary
' Description   : Populates all Leadagent and Individual tree nodes at the top-
'                 level of the tree structure on frmMain. This routine is
'                 called once for each type of intermediary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateRootIntermediary(ByRef tv As MSComctlLib.TreeView, ByVal uIntermediaryType As IntermediaryTypeEnum, ByRef clsIntermediaryTable As IntermediaryTable, ByRef node As MSComctlLib.node, Optional ByVal bCheckChildren As Boolean = True)

    Dim sTitle As String
    Dim vGUID  As Variant
    Dim nThisItem As Long
    Dim nChildren As Long
    Dim bNewRec As Boolean
    Dim sCurGUID As String
    Dim bIsChild As Boolean
    Dim vAttachedTo As Variant
    Dim sParentKeyNode  As String
    
    On Error GoTo Failed
    
    'Store the parent key.
    sParentKeyNode = m_sParentNodeKey
    
    'ClsIntermediaryTable contains the GUID, type and ChildCount as this point.
        
    'Are there any intermediaries in the table object?
    If TableAccess(clsIntermediaryTable).RecordCount > 0 Then
        'Move to the first record.
        TableAccess(clsIntermediaryTable).MoveFirst
        
        'Iterate through each intermediary record and add a node.
        For nThisItem = 1 To TableAccess(clsIntermediaryTable).RecordCount
            'Get the record's GUID.
            vGUID = g_clsSQLAssistSP.GuidToString(clsIntermediaryTable.GetIntermediaryGUID)
            
            'Initialise the new record flag.
            bNewRec = False
            
            'If the GUID has altered since the last record then trip the flag.
            If sCurGUID <> vGUID Then
                bNewRec = True
                
                'Reset the parent.
                m_sParentNodeKey = sParentKeyNode
                sCurGUID = vGUID
            End If
            
            'Stop duplicate GUIDs being added in a row.
            If bNewRec Then
                'Obtain the number of child intermediaries from the current one.
                If bCheckChildren Then
                    nChildren = clsIntermediaryTable.GetNoChildren
                Else
                    nChildren = 0
                End If
                                
                
' EP15 pct
                'Get the name of the intermediary, it will be added as a caption.
                Select Case uIntermediaryType
                    Case LeadAgentType: sTitle = clsIntermediaryTable.GetIntermediaryName(LEADAGENT)
                    Case IndividualType: sTitle = clsIntermediaryTable.GetIntermediaryName(INDIVIDUAL)
                    Case PackagerType: sTitle = clsIntermediaryTable.GetIntermediaryName(PACKAGER)
                    Case BrokerType: sTitle = clsIntermediaryTable.GetIntermediaryName(BROKER)
                End Select
                
'                'Get the name of the intermediary, it will be added as a caption.
'                If uIntermediaryType = LeadAgentType Then
'                    sTitle = clsIntermediaryTable.GetIntermediaryName(LEADAGENT)
'                Else
'                    sTitle = clsIntermediaryTable.GetIntermediaryName(INDIVIDUAL)
'                End If

' EP15 pct End
                'Is there a parent record to this one?
                vAttachedTo = clsIntermediaryTable.GetParentIntermediaryGUID
                
                'Initialise a flag to indicate there is no parent.
                bIsChild = False
                
                'If the vAttachedTo GUID is not null then there is a parent.
                If Not IsNull(vAttachedTo) Then
                    If Len(vAttachedTo) > 0 Then
                        bIsChild = True
                    End If
                End If
                
                'If this record has no parent and has a new GUID (from the last record).
                If (Not bIsChild) And bNewRec Then
                    If Not IsNull(vGUID) Then
                        'Add the tree node for this intermediary.
                        AddNode tv, sTitle, CStr(vGUID), True, GetIntermediaryTypeAsText(uIntermediaryType)
                    End If
                End If
                            
                'If this isn't an individual (i.e. it CAN have children) AND
                'it has one or more children AND its has no parent then add a dummy node.
                
                'This dummy added as a child node to enable the '+' symbol.
                If (uIntermediaryType <> IndividualType) And (nChildren > 0) And (Not bIsChild) Then
                    AddNode tv, "", CStr(vGUID), , GetIntermediaryTypeAsText(uIntermediaryType), "Dummy", True
                End If
            End If
            
            'Move onto the next record/intermediary.
            TableAccess(clsIntermediaryTable).MoveNext
        Next
    End If
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Initialise the module-levels.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    m_sParentNodeKey = LEADAGENT
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddNode
' Description   : Adds a new node to a specified treeview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddNode(ByRef tv As MSComctlLib.TreeView, ByRef sTitle As String, Optional ByRef sKey As String = "", Optional ByVal bSetParent As Boolean = False, Optional ByVal sType As String, Optional ByVal sKeySuffix As String = "", Optional ByVal bIsSysteoNode As Boolean = False, Optional ByVal bHeader As Boolean = False)

    Dim sGuid As String
    Dim clsTreeItem As TreeItem
    Dim oNode As MSComctlLib.node
    
    On Error GoTo Failed
    
    'Create a treeitem placeholder object for this tree-node.
    Set clsTreeItem = New TreeItem
    
    'Store the original key value (it will be altered).
    sGuid = sKey
    
    'If no key was specified, then use the specified title instead.
    If Len(sKey) = 0 Then
        sKey = sTitle
    End If
    
    'If a suffix was specified, then append it to the key now.
    If Len(sKeySuffix) > 0 Then
        sKey = sKey & sKeySuffix
    End If
    
    'Add a treenode now to the currently stored parent node.
    Set oNode = tv.Nodes.Add(m_sParentNodeKey, tvwChild, sKey, sTitle, "closed")

    'If this is a header node (i.e. it does not correspond to an actual
    'intermediary) then make it bold.
    tv.Nodes(sKey).Bold = bHeader

    'Store the specified intermediary type in the treeitem object (a literal).
    clsTreeItem.SetType sType
    
    'Store the header flag.
    clsTreeItem.SetIsHeader bHeader
    
    'Store the GUID of the underlying record.
    clsTreeItem.SetTag sGuid
    
    'Store this flag. TODO: What is it used for?
    clsTreeItem.SetIsSupervisorSystemNode bIsSysteoNode
    
    'Store this placeholder object in the node's tag property.
    Set oNode.Tag = clsTreeItem
    
    'Make this node the new parent node.
    If bSetParent Then
        m_sParentNodeKey = sKey
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryTypeAsText
' Description   : From the specified intermediary enumeration return a string
'                 literal.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIntermediaryTypeAsText(ByVal uIntermediaryType As IntermediaryTypeEnum) As String

    Dim sReturn As String
    
    Select Case uIntermediaryType
' EP15 pct
        Case PackagerType
            sReturn = PACKAGER
            
        Case BrokerType
            sReturn = BROKER
' EP15 pct End

        Case LeadAgentType
            sReturn = LEADAGENT
            
        Case IndividualType
            sReturn = INDIVIDUAL
            
        Case CompanyType
            sReturn = INTERMEDIARY_COMPANY
            
        Case AdminCentreType
            sReturn = ADMIN_CENTRE
    End Select
        
    GetIntermediaryTypeAsText = sReturn
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryTypeAsEnum
' Description   : From the specified string literal, return an intermediary type enum.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIntermediaryTypeAsEnum(ByVal sType As String) As IntermediaryTypeEnum

    Dim uReturn As IntermediaryTypeEnum
    
    Select Case sType
' EP15 pct
        Case PACKAGER
            uReturn = PackagerType
            
        Case BROKER
            uReturn = BrokerType
' EP15 pct End

        Case LEADAGENT
            uReturn = LeadAgentType
            
        Case INDIVIDUAL
            uReturn = IndividualType
            
        Case INTERMEDIARY_COMPANY
            uReturn = CompanyType
        
        Case ADMIN_CENTRE
            uReturn = AdminCentreType
    End Select
    
    GetIntermediaryTypeAsEnum = uReturn
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetParentNodeKey
' Description   : Store a reference to the specified node key.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetParentNodeKey(ByRef sKey As String)
    m_sParentNodeKey = sKey
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the frmMain Listview with the intermediary details
'                 from the node selected in the treeview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateListView(ByRef clsTreeItem As TreeItem, ByRef lvListView As MSGListView, ByRef tvNode As MSComctlLib.node)

    Dim sGuid As String
    Dim sType As String
    Dim bHeaderNode As Boolean
    Dim clsIntermediaryTable As IntermediaryTable
    
    On Error GoTo Failed
    
    'Create a new intermediary table used to load with data.
    Set clsIntermediaryTable = New IntermediaryTable
    
    'Read these properties from the treeitem.
    sGuid = clsTreeItem.GetTag
    sType = clsTreeItem.GetType
    bHeaderNode = clsTreeItem.GetIsHeader
    
    'Get the intermediary details for the specified intermediary.
    clsIntermediaryTable.GetIntermediaryDetails sGuid, sType
    
    If Not bHeaderNode Then
        'Populate the listview from the underlying table object.
        g_clsFormProcessing.PopulateFromRecordset lvListView, clsIntermediaryTable
    Else
        'If a header node is clicked on, simply ensure the listview is cleared.
        lvListView.ListItems.Clear
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateSingleIntermediary
' Description   : Populate all child nodes for selected node in the main
'                 treeview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateSingleIntermediary(ByRef tv As MSComctlLib.TreeView, ByRef node As MSComctlLib.node, ByRef clsTreeItem As TreeItem)

    Dim sType As String
    Dim sGuid As String
    Dim sIntType As String
    Dim tvParentNode As MSComctlLib.node
    Dim clsIntermediaryTable As IntermediaryTable
    
    On Error GoTo Failed
        
    'Create an intermediary table object to populate with records.
    Set clsIntermediaryTable = New IntermediaryTable
    
    'Release the reference to the current node.
    frmMain.ClearCurrentTreeitem
    
    'Ensure all added nodes are removed apart from the LeadAgent and Individual nodes.
    If node.Text = INTERMEDIARIES Then
        HandleIntermediaryRefresh False
    End If
    
    'This stops the routine processing the top-level nodes which have no
    'treeitem tags. i.e. The LeadAgent and Individual top-level nodes.
    If Not clsTreeItem Is Nothing Then
        'Read the intermediary type and guid from the treeitem.
        sType = clsTreeItem.GetType
        sGuid = FormatKey(clsTreeItem.GetTag)
        
        'Map the literal to an enum _ although we'll still store it in a string!
        sIntType = GetIntermediaryTypeAsEnum(sType)
         
        'Get a reference to this nodes parent.
        Set tvParentNode = node.Parent
        
        'If the item is a top-level node, use GetIntermediaries, otherwise use
        'the GetChildIntsWithChildren routine.
        If sType = LEADAGENT Then
            'Get the types of any child intermediaries for the current item.
            clsIntermediaryTable.GetIntermediaries , sGuid
        Else
            'Get the child records of the current one.
            clsIntermediaryTable.GetChildIntsWithChildren sGuid, sIntType
        End If

        'If records were returned, then we'll populate the treeview.
        If TableAccess(clsIntermediaryTable).RecordCount > 0 Then
            'Store this node as the current parent.
            m_sParentNodeKey = node.Key
             
            'Populate all the child nodes from the specified data object.
            PopulateChildIntermediaries sGuid, node, TableAccess(clsIntermediaryTable), sType, tv, sIntType
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateChildIntermediaries
' Description   : Populates all child nodes from the specified node when it is expanded for the
'                 first time. This routine is called from PopulateSingleIntermediary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateChildIntermediaries(ByVal sGuid As String, ByRef tvNode As MSComctlLib.node, ByRef clsSrcIntermediary As IntermediaryTable, ByVal sType As String, ByRef tv As MSComctlLib.TreeView, Optional ByVal sIntType As String = "")

    Dim nThisItem As Long
    Dim vCurKey As Variant
    Dim sIntName As String
    Dim sChildKey As String
    Dim sChildType As String
    Dim sChildGUID As String
    Dim clsTreeItem As TreeItem
    Dim vIsAttachedTo As Variant

    On Error GoTo Failed
    
    'Remove all child nodes of the parent.
    If TypeOf tvNode.Child.Tag Is TreeItem Then
        Set clsTreeItem = tvNode.Child.Tag

        If clsTreeItem.GetIsSupervisorSystemNode = True Then
            tv.Nodes.Remove tvNode.Child.Key
        End If
    End If

    'Remove any child nodes from this node.
    RemoveChildNodes tv, tvNode

    'Iterate through each child record.
    For nThisItem = 1 To TableAccess(clsSrcIntermediary).RecordCount
        'Store this node as the current parent.
        m_sParentNodeKey = tvNode.Key
        
        'Store the GUID and Type (literal) for the first child.
        sChildGUID = g_clsSQLAssistSP.GuidToString(clsSrcIntermediary.GetIntermediaryGUID)
        sChildType = GetIntermediaryTypeAsText(CStr(clsSrcIntermediary.GetIntermediaryType))

        'Because the SQL returns all Intermediary child types, the intermediaryguid is not
        'always unique. Therefor a check must be processed to see if this is a different record
        If (sChildGUID & sChildType) <> vCurKey Then
            vCurKey = sChildGUID & sChildType
            vIsAttachedTo = clsSrcIntermediary.GetParentIntermediaryGUID

            'Get the parent guid if this is a child.
            If Not IsNull(vIsAttachedTo) Then
                vIsAttachedTo = g_clsSQLAssistSP.GuidToString(CStr(vIsAttachedTo))
            Else
                vIsAttachedTo = ""
            End If

            'TODO: I don't know what this is supposed to be doing. If time allows,
            'it should be commented....
            If Len(vIsAttachedTo) = 0 And Len(sIntType) > 0 Or sType <> LEADAGENT Then
                If Len(vIsAttachedTo) = 0 And Len(sChildType) > 0 Then
                    sIntName = clsSrcIntermediary.GetIntermediaryName(sChildType)
                Else
                    sIntName = clsSrcIntermediary.GetIntermediaryName(sType)
                End If

                If tvNode.children > 0 And nThisItem = 1 Then
                    m_sParentNodeKey = tvNode.Child.Key
                    sChildKey = m_sParentNodeKey
                Else
                    m_sParentNodeKey = tvNode.Key
                    sChildKey = ""
                End If

                If clsSrcIntermediary.GetNoChildren > 0 Then
                    If Not DoesNodeExist(tv, sChildGUID) Then
                        AddNode tv, sIntName, sChildGUID, True, sType
                    End If

                    If Not DoesNodeExist(tv, sChildGUID & sChildType) Then
                        m_sParentNodeKey = sChildGUID
                        AddNode tv, sChildType, sChildGUID, True, sChildType, sChildType, , True
                        AddNode tv, "", sChildGUID, True, sChildType, sChildType, True
                    End If
                Else
                    If sChildGUID <> sChildKey And Not DoesNodeExist(tv, sChildGUID) Then
                        AddNode tv, sIntName, sChildGUID, False, sType
                    End If
                End If

            Else
                If sGuid = vIsAttachedTo Or Len(vIsAttachedTo) = 0 Then
                    m_sParentNodeKey = sGuid
                End If

                AddNode tv, sChildType, sChildGUID, True, sChildType, sChildType, , True
                AddNode tv, "", sChildGUID, , sChildType, sChildType, True
            End If
        End If

        TableAccess(clsSrcIntermediary).MoveNext
    Next

    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesNodeExist
' Description   : Returns a boolean indicating whether a node exists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesNodeExist(ByRef tv As MSComctlLib.TreeView, ByVal sGuid As String) As Boolean
    
    Dim bRet As Boolean
    Dim tvNode As MSComctlLib.node
    
    On Error Resume Next
    
    bRet = True
    Err.Clear
    
    Set tvNode = tv.Nodes(sGuid)

    If Err.Number <> 0 Then
        bRet = False
    Else
        Err.Clear
    End If
    
    DoesNodeExist = bRet

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateLV
' Description   : Populates the LV from given combogroup for a given LV
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateLV(ByVal sGroup As String, ByRef lv As MSGListView)

    Dim nCnt As Long
    Dim rs As ADODB.Recordset
    Dim clsCombo As ComboValueTable
    Dim lvItem As MSComctlLib.ListItem
    Dim clsComboValidation As ComboValidationTable
    
    On Error GoTo Failed
    
    Set clsCombo = New ComboValueTable
    Set clsComboValidation = New ComboValidationTable
    
    'Get rs for combo group
    Set rs = clsComboValidation.GetComboGroupWithValidationType(sGroup)
    
    TableAccess(clsCombo).SetRecordSet rs
    rs.Sort = "VALUEID ASC"
    
    'Populate LV with ValueName and set the key to ValueID
    For nCnt = 1 To TableAccess(clsCombo).RecordCount
        Set lvItem = lv.ListItems.Add
        
        lvItem.Text = clsCombo.GetValueName
        lvItem.Key = GetKey(clsCombo.GetValueID)
        
        'Put the valueid in the tag of the lv item
        lvItem.Tag = clsCombo.GetValueID
        TableAccess(clsCombo).MoveNext
    Next

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKey
' Description   : Retuns a key to be used to index a listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetKey(ByVal vval As Variant) As String
    
    On Error GoTo Failed

    GetKey = KEY_PREFIX & CStr(vval)
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetCheckedItemsAsCollection
' Description   : Return a collection of all checked items for a given LV
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCheckedItemsAsCollection(ByRef lv As MSGListView) As Collection

    Dim col As Collection
    Dim nThisItem As Long
    Dim lvItem As MSComctlLib.ListItem
    
    On Error GoTo Failed
    
    Set col = New Collection
    
    For Each lvItem In lv.ListItems
        If lvItem.Checked Then
            col.Add lvItem.Tag
        End If
    Next
    
    Set GetCheckedItemsAsCollection = col
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateParentIntCombo
' Description   : Depending on the validation type of the current intermediary, all suitable
'                 parent ints are populated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateParentIntCombo(ByRef cboIntermediaryType As MSGDataCombo, ByVal uIntermediaryType As IntermediaryTypeEnum)
    
    Dim colTypeList As Collection
    Dim rsTypes As ADODB.Recordset
    Dim clsIntermediaryTable As IntermediaryTable
    
    On Error GoTo Failed
    
    'Create a table object to populate the combo from.
    Set clsIntermediaryTable = New IntermediaryTable
    
    'Create a collection to hold valid parent types.
    Set colTypeList = New Collection
    
    'Based upon the intermediary-type, build a collection of validation types.
    Select Case uIntermediaryType
        Case CompanyType
            colTypeList.Add CStr(LeadAgentType)

        Case AdminCentreType
            colTypeList.Add CStr(CompanyType)
            colTypeList.Add CStr(LeadAgentType)

        Case IndividualType
            colTypeList.Add CStr(LeadAgentType)
            colTypeList.Add CStr(CompanyType)
            colTypeList.Add CStr(AdminCentreType)

        Case LeadAgentType
            'A Lead Agent may not have any parents.
    End Select

    'Populate the table object from the specified validation types.
    clsIntermediaryTable.GetValidParentIntermediaries colTypeList
    
    'Get a reference to the underlying recordset.
    Set rsTypes = TableAccess(clsIntermediaryTable).GetRecordSet
        
    'Bind the combo supplied to the recordset.
    Set cboIntermediaryType.RowSource = rsTypes
    cboIntermediaryType.ListField = clsIntermediaryTable.GetParentIntermediaryListField
    cboIntermediaryType.BoundColumn = clsIntermediaryTable.GetIntermediaryGUIDField
    
    'Ensure the data is sorted.
    If Not cboIntermediaryType.RowSource Is Nothing Then
        cboIntermediaryType.RowSource.Sort = "INTERMEDIARYPANELID DESC"
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateDate
' Description   : Checks that the from date occurs before the to date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateDate(ByRef txtFrom As MSGEditBox, ByRef txtTo As MSGEditBox, Optional ByVal sError As String = "") As Boolean
    
    On Error GoTo Failed

    Dim dTo As Date
    Dim dFrom As Date
    Dim bRet As Boolean
    Dim vDate As Variant
    Dim bNull As Boolean
        
    'Get the active from Date
    g_clsFormProcessing.HandleDate txtFrom, vDate, GET_CONTROL_VALUE
    
    If Not IsNull(vDate) Then
        dFrom = CDate(vDate)
    Else
        bNull = True
    End If
    
    'Get the active to date
    g_clsFormProcessing.HandleDate txtTo, vDate, GET_CONTROL_VALUE
    
    If Not bNull And Not IsNull(vDate) Then
        dTo = CDate(vDate)
    End If
    
    bRet = True
    
    If dTo > 0 Then
        If dFrom > dTo And Not bNull Then
            'There is an error
            If Len(sError) = 0 Then
                'No optional error message set
                sError = "Active from date must occur before the active to date"
            End If
            
            bRet = False
        End If
        
        If bNull Or Not bRet Then
            g_clsErrorHandling.DisplayError sError
            txtFrom.SetFocus
        End If
    End If
    
    ValidateDate = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleAddNode
' Description   : Public function to handle the addition of both header and type nodes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleAddNode(ByRef tv As MSComctlLib.TreeView, ByVal sTitle As String, Optional ByVal sKey As String = "", Optional ByVal bSetParent As Boolean = False, Optional ByVal sType As String, Optional ByVal sKeySuffix As String = "", Optional ByVal bIsSystemNode As Boolean = False, Optional ByVal bIsNew As Boolean = False)

    On Error GoTo Failed

    Dim sNewKey As String
    Dim bNodeExists As Boolean
    
    'Convert the enum to a literal.
    sType = GetIntermediaryTypeAsText(CLng(sType))
    
    'We now have the true type of the intermediay ie Admin LeadAgent etc
    'if not leadagent then we need to add a new child node to an existing node.
    If Len(sType) > 0 Or Not bIsNew And sType <> LEADAGENT Or Not bIsNew Then
        'We can check if a type node is to be added by checking if there is already an
        'existing node with a key as the new type and parent guid
        If m_sParentNodeKey <> sType Then
            sNewKey = m_sParentNodeKey & sType
        Else
            sNewKey = m_sParentNodeKey
        End If
        
        bNodeExists = DoesNodeExist(tv, sNewKey)
        
        If Not bNodeExists Then
            'First add the header
            If sNewKey <> LEADAGENT And sNewKey <> INDIVIDUAL Or sType <> INDIVIDUAL Then
                AddNode tv, sType, FormatKey(sNewKey), True, sType, sType, True, True
            End If
        Else
            m_sParentNodeKey = sNewKey
        End If
    End If
    
    'Add the detail Node
    AddNode tv, sTitle, sKey, bSetParent, sType, sKeySuffix, bIsSystemNode
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RemoveChildNodes
' Description   : Clears all child nodes before populating the children
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveChildNodes(ByRef tv As MSComctlLib.TreeView, ByRef tvNode As MSComctlLib.node)
    
    Dim nNode As Long
    Dim NodeChild As MSComctlLib.node
    
    On Error GoTo Failed
    
    For nNode = 1 To tvNode.children
        Set NodeChild = tvNode.Child
        
        tv.Nodes.Remove NodeChild.Key
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValueIDForIntType
' Description   : Given an Intermediary Validation Type the value ID is returned
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetValueIDForIntType(ByVal sType As String) As Integer
    
    Dim nValueID As Integer
    Dim clsCombo As ComboValidationTable
    
    On Error GoTo Failed
    
    Set clsCombo = New ComboValidationTable
    
    nValueID = clsCombo.GetSingleValueFromValidation("IntermediaryType", sType)
    GetValueIDForIntType = nValueID
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleIntermediaries
' Description   : Called via MainSuper to open the intermediary form in add-mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleIntermediaries(ByVal nIndex As Integer)

    Dim frmEdit As Form
    Dim sParentKey As String
    Dim clsTreeItem As TreeItem
    Dim uReturn As MSGReturnCode
    Dim uIntermediaryType As IntermediaryTypeEnum
    
    On Error GoTo Failed
        
    'Gets a reference to the current tree node.
    Set clsTreeItem = g_clsMainSupport.GetIntermediaryTreeItem
        
    'Map the menu index onto the type-enum.
    Select Case nIndex
        Case LV_INTERMEDIARY_ADMIN
            uIntermediaryType = AdminCentreType
        
        Case LV_INTERMEDIARY_INDIVIDUAL
            uIntermediaryType = IndividualType
        
        Case LV_INTERMEDIARY_COMPANY
            uIntermediaryType = CompanyType
        
        Case LV_INTERMEDIARY_LEADAGENT
            uIntermediaryType = LeadAgentType
            
        ' EP15 pct 16/03/2006 -
        Case LV_INTERMEDIARY_PACKAGER
            uIntermediaryType = PackagerType
            
        Case LV_INTERMEDIARY_BROKER
            uIntermediaryType = BrokerType
            
        ' EP15 pct END 16/03/2006
    End Select
    
    'Get a local reference to the intermediaries form.
    Set frmEdit = frmEditIntermediaries
    
    'Specify the intermediary type, we'll be adding.
    frmEdit.SetIntermediaryType uIntermediaryType
    
    'Obtain the parent intermediary's key.
    sParentKey = clsTreeItem.GetTag
        
    'Set it against the form.
    frmEdit.SetParentKey FormatKey(sParentKey)
    
    'Ensure the form is in Add mode.
    frmEdit.SetIsEdit False
        
    'Begin a transaction.
    g_clsDataAccess.BeginTrans
    
    'Show the main intermediary form.
    frmEdit.Show vbModal
        
    'Get the return code from the form.
    uReturn = frmEdit.GetReturnCode
    
    'frmEdit.Hide
    g_clsMainSupport.HandleReturn uReturn
        
    'Unload the form now.
    Unload frmEdit

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTreeItem
' Description   : Store a reference to the specified tree item.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTreeItem(ByRef clsTreeItem As TreeItem)
    Set m_clsTreeItem = clsTreeItem
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTreeItem
' Description   : Returns a reference to the stored treeitem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTreeItem() As TreeItem
    Set GetTreeItem = m_clsTreeItem
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleInteremediarySearch
' Description   : Display the search Screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HandleIntermediarySearch()

    Dim frmSearch As frmFindIntermediary
    
    On Error GoTo Failed
    
    'Create a new search screen.
    Set frmSearch = frmFindIntermediary
    
    'Show the screen modally.
    frmSearch.Show vbModal
    
    'Unload the form.
    Unload frmSearch
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FormatKey
' Description   : Filters the type out of the key if present
'                 when the user adds an intermediary from a type node (ie admin centre) the key
'                 contains Admin Centre.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FormatKey(ByVal sKey As String) As String

    Dim nFound As Long
    Dim vType As Variant
    Dim col As Collection
    
    On Error GoTo Failed
    
    Set col = New Collection
    col.Add LEADAGENT
    col.Add INDIVIDUAL
    col.Add INTERMEDIARY_COMPANY
    col.Add ADMIN_CENTRE
    col.Add INTERMEDIARIES
    col.Add PACKAGER  ' EP15 pct 16/03/2006
    col.Add BROKER  ' EP15 pct 16/03/2006
    
    For Each vType In col
        nFound = InStr(1, sKey, vType)
        
        If nFound > 0 Then
            sKey = Left(sKey, nFound - 1)
        End If
    Next

    FormatKey = sKey
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : UpdateNode
' Description   : Updates the node in the main tree view to reflect the changes made
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateNode(ByVal vIntermediaryGUID As Variant, ByVal vParentGUID As Variant, ByVal sNodeText As String, ByVal uIntermediaryType As IntermediaryTypeEnum, ByVal bIsEdit As Boolean)

    Dim sType As String
    Dim bCreate As Boolean
    Dim sNodeKey As String
    Dim node As MSComctlLib.node
    Dim sIntermediaryGUID As String
    Dim clsIntermediary As Intermediary
    
    On Error GoTo Failed
    
    'Convert the IntermediaryGUID into a string for comparison work within this routine.
    sIntermediaryGUID = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vIntermediaryGUID))
    
    'Create an Intermediary helper-class.
    Set clsIntermediary = New Intermediary
    
    'Ensure the current referenced node is cleared.
    frmMain.ClearCurrentTreeitem
    
    'Get a reference to the current tree-node.
    Set node = frmMain.tvwDB.SelectedItem
    
    'If the node has a parent, get its key.
    If Len(vParentGUID) > 0 Then
        sNodeKey = vParentGUID
    Else
        'Otherwise get the node's key.
        sNodeKey = node.Key
        
        'Has the user right clicked on the intermediary node?
        If sNodeKey = INTERMEDIARIES Then
' EP15 pct
            Select Case uIntermediaryType
                Case Is = IndividualType: sNodeKey = INDIVIDUAL
                Case Is = LeadAgentType: sNodeKey = LEADAGENT
                Case Is = PackagerType: sNodeKey = PACKAGER
                Case Is = BrokerType: sNodeKey = BROKER
            End Select
            

'            'Are we adding an individual or leadagent?
'            If uIntermediaryType = IndividualType Then
'                sNodeKey = INDIVIDUAL
'            Else
'                sNodeKey = LEADAGENT
'            End If

' EP15 pct End

        End If
    End If
    
    bCreate = False
            
    If Not node Is Nothing Then
        If node.Child Is Nothing Then
            If Not bIsEdit Then
                bCreate = True
            End If
        ElseIf Len(node.Child.Text) = 0 And bIsEdit Then
            bCreate = False
        Else
            bCreate = False
        End If
    
        If bCreate Or Not bIsEdit Then
            'Create new node.
            sType = CStr(uIntermediaryType)
            clsIntermediary.SetParentNodeKey sNodeKey
            clsIntermediary.HandleAddNode frmMain.tvwDB, sNodeText, sIntermediaryGUID, False, sType, , , True
            
        ElseIf bIsEdit And sNodeKey <> sIntermediaryGUID Then
            'Is the selected node intermediaries node?
            If node.Key <> INTERMEDIARIES Then
                'Amend existing node
                node.Text = sNodeText
            End If
        Else
            node.Text = sNodeText
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

