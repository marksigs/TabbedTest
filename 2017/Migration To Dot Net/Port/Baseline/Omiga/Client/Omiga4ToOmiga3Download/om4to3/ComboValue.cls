VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "ComboValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      {ThisFileName}
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Used by RulesAssist to convert combovalueids into descriptions, OM3 values or key values
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'        Created
'------------------------------------------------------------------------------------------

Option Explicit

'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================
Private m_strGroupName As String
Private m_strValueID As String
Private m_strOM3ValueID As String
Private m_strValueType As String
Private m_strValueName As String
Private m_objSibling As ComboValue

Implements ObjectControl

Public Property Get SiblingCount() As Long
' header ----------------------------------------------------------------------------------
' description:
'    For a particular Groupname and ValueID or ValidationID there can be
'    multiple entries. This is the minority case. For these cases a chain
'    of siblings is built up.
'
' return: The number of siblings
'------------------------------------------------------------------------------------------
    If m_objSibling Is Nothing Then
        SiblingCount = 0
    Else
        SiblingCount = 1 + m_objSibling.SiblingCount
    End If
End Property

Public Property Get NextSibling() As Object
    Set NextSibling = m_objSibling
End Property

Public Property Get ValueIDKey() As String
    ValueIDKey = UCase$(m_strGroupName & "." & m_strValueID)
End Property

Public Property Get ValueTypeKey() As String
    ValueTypeKey = UCase$(m_strGroupName & "." & m_strValueType)
End Property

Public Property Get GroupName() As String
    GroupName = m_strGroupName
End Property

Public Property Get ValueID() As String
    ValueID = m_strValueID
End Property

Public Property Get OM3ValueID() As String
    OM3ValueID = m_strOM3ValueID
End Property

Public Property Get ValueType() As String
    ValueType = m_strValueType
End Property

Public Property Get ValueName() As String
    ValueName = m_strValueName
End Property


Public Sub Create(ByVal vstrGroupName As String, _
                    ByVal vstrValueID As String, _
                    ByVal vstrOM3ValueID As String, _
                    ByVal vstrValue As String, _
                    ByVal vstrValueName As String)
' header ----------------------------------------------------------------------------------
' description:
'    Creates this object. A Create is used to force a Read-only status on
'    its attributes.
'
'------------------------------------------------------------------------------------------


     m_strGroupName = vstrGroupName
     m_strValueID = vstrValueID
     m_strOM3ValueID = vstrOM3ValueID
     m_strValueType = vstrValue
     m_strValueName = vstrValueName
     Set m_objSibling = Nothing
    
End Sub

Public Sub AddSibling(ByVal vobjSibling As ComboValue)
' header ----------------------------------------------------------------------------------
' description:
'    Siblings are defined when the key used to store this item in the
'    collection is in conflict. In this case a chain of siblings is built
'    up.
'
' pass: vobjsibling        The sibling to add. May be Nothing. May recurse
'
' return:
'------------------------------------------------------------------------------------------

    If vobjSibling Is Nothing Then
        'Do nothing. This stops recursion
    Else
        'Move the current sibling (if any) down the chain
        vobjSibling.AddSibling m_objSibling
        'make the new sibling the favoured one
        Set m_objSibling = vobjSibling
    End If
End Sub

Private Sub ObjectControl_Activate()
'
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objSibling = Nothing
End Sub
