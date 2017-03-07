VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module: SecurityManager
' Description : Supervisor Security controller class. Provides the business
'               routines used to configure security access in Supervisor.
'
' The list of available security resources is currently built dynamically by
' supervisor (as the treeitems are added) and also, internally within this
' class (for items which do not correspond to treeitems).
'
' Some terms will need clarification: -
'
'   Resource - Represents an operation on an object which can be permissioned.
'              An example might be 'Edit, Mortgage Product' or 'Promote, User'.
'   Operation- An action someone can perform on a Resource. Security could be
'              set to this level if required, although currently only ObjectIDs
'              are used in the database (and not Operations).
'   Agent    - Represents an entity which can use a resource (user, group,
'              etc).
'
' Change history
' Prog      Date        Description
' STB       15/01/02    Created class.
' STB       22/01/02    SYS2957 Supervisor connection passed as reference.
' STB       04/02/02    SYS2957 Promotion of Users was flawed.
' STB       13/02/02    SYS4052 HasAccess returns TRUE if UserID is blank.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'This interface is defined in an external type-library (SecurityMgtInterface.tlb)
'to ensure client applications can't get access to the security management
'routines.
Implements ISecurityManager

'Permissions may be set upon any of these agents.
Public Enum AgentTypeEnum
    AgentTypeUser = 1   '1
    AgentTypeGroup      '2
End Enum

'Collection of all defined security resources (objects+operations) which can be
'granted or denied to agents.
Private m_clsSecurityResourceList As SecurityResourceList

'The connection used to access the database. This ensures we can use supervisor
'transactions.
Private m_clsConnection As ADODB.Connection


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Store the connection string used to access the database. This
'                 routine will return false if the security table doesn't exist
'                 in the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Initialise(ByRef clsConnection As ADODB.Connection) As Boolean
    
    Dim clsCommand As ADODB.Command
    
    On Error GoTo Err_Handler
    
    'Store the connection string to be used later.
    Set m_clsConnection = clsConnection
    
    'The parameter prefix is SQL-Server/Oracle dependant.
    ADOSetParameterPrefix clsConnection
    
    'Get a dummy command to execute (if it fails then we'll deduce that the
    'security table isn't present).
    Set clsCommand = ADOGetCommand(USP_GET_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString("")
    clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = Null

    'Test for existance of the SupervisorUserResource table.
    ADOExecuteCmd clsCommand, m_clsConnection
    
Err_Handler:
    If Err Then
        'Indicate failure to initialise.
        Initialise = False
        
        'Raise an error indicating a problem has occured in this component.
        Err.Raise 445, "Initialise", "Could not access the security objects in the database. This version of Supervisor will not work with the system you are trying to connect to. The reported error was: " & Err.Description
    Else
        Initialise = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Instantiate any object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    
    'Create a master-list of resources.
    Set m_clsSecurityResourceList = New SecurityResourceList
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Terminate
' Description   : Release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    
    'Release object references.
    Set m_clsConnection = Nothing
    Set m_clsSecurityResourceList = Nothing
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddSecurityResource
' Description   : Called by supervisor as items are added into the treeview.
'                 These items then become security resources available to be
' permissioned. This is currently assigned to a global list (within this
' component) so that the SecurityVerifier can reference and use it. Eventually
' this virtual table will be replaced by a physical table (and this method can
' then be removed).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_AddSecurityResource(ByVal lObjectID As Long, ByVal sObjectName As String, Optional ByVal sOperation As String) As Boolean

    Dim clsResource As SecurityResource

    'Create a security resource object to add into the master-list.
    Set clsResource = New SecurityResource

    'Set the object and operation properties.
    clsResource.ObjectID = lObjectID
    clsResource.ObjectName = sObjectName
    clsResource.Operation = sOperation

    'Add the resource to the master-list.
    m_clsSecurityResourceList.Add clsResource

    'Indicate success.
    ISecurityManager_AddSecurityResource = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllocatedResources
' Description   : Returns a SecurityResourceList for any SecurityResources
'                 assigned to the specified agent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_GetAllocatedResources(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum) As Security.SecurityResourceList

    Dim clsCommand As ADODB.Command
    Dim rsAgentResources As ADODB.Recordset
    Dim clsResourceList As SecurityResourceList

    'Get a command object and populate the parameters.
    Set clsCommand = ADOGetCommand(USP_GET_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)
    clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = Null

    'Get a recordset of the agent's SecurityResources.
    Set rsAgentResources = ADOExecuteCmd(clsCommand, m_clsConnection)

    'Populate a resource list from this recordset (and the master resource list).
    Set clsResourceList = New SecurityResourceList
    clsResourceList.LoadFromRecordset rsAgentResources, m_clsSecurityResourceList

    'Return the resource list to the caller.
    Set ISecurityManager_GetAllocatedResources = clsResourceList

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAvailableResources
' Description   : Return a reference to the list of SecurityResources
'                 available within Supervisor.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_GetAvailableResources() As Security.SecurityResourceList
    Set ISecurityManager_GetAvailableResources = m_clsSecurityResourceList
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PromoteAgent
' Description   : Copies the named agent's resource access to the target
'                 system.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_PromoteAgent(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum, clsTargetConnection As ADODB.Connection) As Boolean

    Dim clsCommand As ADODB.Command
    Dim rsAgentResources As ADODB.Recordset
    
    On Error GoTo Err_Handler
    
    'Remove all rights for this agent on the target system before setting new ones.
    Set clsCommand = ADOGetCommand(USP_DELETE_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)
    
    'Execute the command and remove permissions.
    ADOExecuteCmd clsCommand, clsTargetConnection
    
    'Get a command object and populate the parameters.
    Set clsCommand = ADOGetCommand(USP_GET_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)
    clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = Null

    'Get a recordset of the agent's SecurityResources.
    Set rsAgentResources = ADOExecuteCmd(clsCommand, m_clsConnection)
    
    'If we have resources for this agent, we'll set each one.
    If (rsAgentResources.EOF = False) And (rsAgentResources.BOF = False) Then
        'Move to the first SecurityResource record.
        rsAgentResources.MoveFirst

        'Get a command object and populate the parameters.
        Set clsCommand = ADOGetCommand(USP_SET_AGENT_RESOURCE)
        clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)
        clsCommand.Parameters(g_sParamPrefix & "ALLOW").Value = True
        
        'Iterate through each one and attempt to set it in the target system.
        Do While Not rsAgentResources.EOF
            clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = rsAgentResources.Fields("ITEMID").Value
            
            'Execute the command and set the permission.
            ADOExecuteCmd clsCommand, clsTargetConnection
        
            'Move onto the next resource.
            rsAgentResources.MoveNext
        Loop
    End If
    
Err_Handler:
    If Err Then
        ISecurityManager_PromoteAgent = False
        Err.Raise 445, "PromoteAgent", "Unable to promote Agent " & sAgentID & " to the target database. The reported error was: " & Err.Description
    Else
        ISecurityManager_PromoteAgent = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RemoveAllAccess
' Description   : Removes all defined permission records for the named agent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_RemoveAllAccess(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum) As Boolean

    Dim clsCommand As ADODB.Command

    'Get a command object and set the parameters.
    Set clsCommand = ADOGetCommand(USP_DELETE_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)

    'Execute the command and delete all the resources allocated to this agent.
    ADOExecuteCmd clsCommand, m_clsConnection

    'Indicate success.
    ISecurityManager_RemoveAllAccess = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAccess
' Description   : Grant or deny access to the specified agent/resource in the
'                 database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ISecurityManager_SetAccess(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum, ByVal bAllow As Boolean, ByVal lObjectID As Long, Optional ByVal sOperation As String = "") As Boolean

    Dim clsCommand As ADODB.Command

    'Get a command object and populate the parameters.
    Set clsCommand = ADOGetCommand(USP_SET_AGENT_RESOURCE)
    clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sAgentID)
    clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = lObjectID
    clsCommand.Parameters(g_sParamPrefix & "ALLOW").Value = bAllow

    'Execute the command and set the permission.
    ADOExecuteCmd clsCommand, m_clsConnection

    'Indicate success.
    ISecurityManager_SetAccess = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HasAccess
' Description   : Indicates if the specified USER has access to the named
'                 SecurityResource.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasAccess(ByVal sUserID As String, ByVal lObjectID As Long, Optional ByVal sOperation As String = "") As Boolean
    
    Dim clsCommand As ADODB.Command

    'If the specified user is blank, simply return true.
    If sUserID = "" Then
        HasAccess = True
    Else
        'Get a command object and populate the parameters.
        Set clsCommand = ADOGetCommand(USP_GET_AGENT_RESOURCE)
        clsCommand.Parameters(g_sParamPrefix & "USERID").Value = ADOFormatSQLString(sUserID)
        clsCommand.Parameters(g_sParamPrefix & "ITEMID").Value = lObjectID
    
        'Execute the command.
        ADOExecuteCmd clsCommand, m_clsConnection
    
        'Get the value from the output parameter and return whether the agent has
        'the access required.
        HasAccess = clsCommand.Parameters(g_sParamPrefix & "ALLOW").Value
    End If
    
End Function
