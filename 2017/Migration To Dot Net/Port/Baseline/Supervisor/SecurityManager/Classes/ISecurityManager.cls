VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISecurityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAvailableResources
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAvailableResources() As Security.SecurityResourceList
Attribute GetAvailableResources.VB_Description = "Return a list of all the resrouces in the virtual table."
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllocatedResources
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAllocatedResources(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum) As Security.SecurityResourceList
Attribute GetAllocatedResources.VB_Description = "Return a list of Resources granted to the specified user or group."
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAccess
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetAccess(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum, ByVal bAllow As Boolean, ByVal lObjectID As Long, Optional ByVal sOperation As String = "") As Boolean
Attribute SetAccess.VB_Description = "Grants access to the specified resource for the specified user or group."
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RemoveAllAccess
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RemoveAllAccess(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum) As Boolean
Attribute RemoveAllAccess.VB_Description = "Removes all access granted to the specified user or group."
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddSecurityResource
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddSecurityResource(ByVal lObjectID As Long, ByVal sObjectName As String, Optional ByVal sOperation As String) As Boolean
Attribute AddSecurityResource.VB_Description = "Add a securable resource into the virtual table."
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PromoteAgent
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PromoteAgent(ByVal sAgentID As String, ByVal uAgentType As Security.AgentTypeEnum, ByRef clsTargetConnection As ADODB.connection) As Boolean
Attribute PromoteAgent.VB_Description = "Copies all the Agent's security access to the target system."
End Function
