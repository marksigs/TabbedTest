VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module: SecurityResource
' Description : Place-holder class which represents an object and an
'               operation which can be denied or granted access to a group
'               of users. The composite of an object and an operation is known
'               as a 'Security Resource'.
'
'               A resource may have just an object name and no operation.
'
' Change history
' Prog      Date        Description
' STB       15/01/02    Created class.
' STB       22/01/02    SYS2957 Added the ObjectID property.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The unique ID given to this securable item.
Private m_lObjectID As Long

'The name of an item which can be permissioned.
Private m_sObjectName As String

'Optional name of an operation on the above item which may be being permissioned.
Private m_sOperation As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ObjectID
' Description   : Return the object ID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ObjectID() As Long
    ObjectID = m_lObjectID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ObjectName
' Description   : Return the object name.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ObjectName() As String
Attribute ObjectName.VB_Description = "The name of an item which a group may be permissioned on."
    ObjectName = m_sObjectName
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Operation
' Description   : Return the name of the operation on this object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Operation() As String
Attribute Operation.VB_Description = "An operation which can be performed on a named object for which access can be controlled."
    Operation = m_sOperation
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ObjectID
' Description   : Set the object ID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Let ObjectID(ByVal lObjectID As Long)
    m_lObjectID = lObjectID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ObjectName
' Description   : Set the object name.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Let ObjectName(ByVal sObjectName As String)
    m_sObjectName = sObjectName
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Operation
' Description   : Set the operation name.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Let Operation(ByVal sOperation As String)
    m_sOperation = sOperation
End Property

