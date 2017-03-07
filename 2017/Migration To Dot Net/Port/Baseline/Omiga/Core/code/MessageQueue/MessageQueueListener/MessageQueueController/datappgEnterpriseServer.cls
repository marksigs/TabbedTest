VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datappgES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      datawpgEprise.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     10/01/01    Created
'------------------------------------------------------------------------------------------

Option Explicit

Public m_SnapInData As SnapInData
Public m_ScopeItem As ScopeItem

' propertypage support
Public m_ppgAccount As ppgESAccount
Public m_ppgPermit As ppgESPermit
Private m_strLastPageDone As String

' Account preferences
Public Enum AccountType
    Account_CurrentUser = 0
    Account_ConnectAsUser = 1
End Enum

Public m_eAccount As AccountType
Public m_strDomain As String
Public m_strUserName As String
Public m_strPassword As String

' Permissions
Public m_bAllowStartService As Boolean
Public m_bAllowStopService As Boolean
Public m_bAllowConnectComputer As Boolean
Public m_bAllowDisconnectComputer As Boolean
Public m_bAllowCreateQueue As Boolean
Public m_bAllowDeleteQueue As Boolean
Public m_bAllowStartQueue As Boolean
Public m_bAllowStallQueue As Boolean
Public m_bAllowDeleteQueueEvent As Boolean
Public m_bAllowCreateQueueEvent As Boolean
Public m_bAllowStartComponent As Boolean
Public m_bAllowStallComponent As Boolean

Public Sub WriteProperties(ByVal PropertyBag As SnapInLib.PropertyBag)

    ' Account
    PropertyBag.WriteProperty "m_eAccount", m_eAccount
    PropertyBag.WriteProperty "m_strDomain", m_strDomain
    PropertyBag.WriteProperty "m_strUserName", m_strUserName
    PropertyBag.WriteProperty "m_strPassword", m_strPassword

    ' Permissions
    PropertyBag.WriteProperty "m_bAllowStartService", m_bAllowStartService
    PropertyBag.WriteProperty "m_bAllowStopService", m_bAllowStopService
    PropertyBag.WriteProperty "m_bAllowConnectComputer", m_bAllowConnectComputer
    PropertyBag.WriteProperty "m_bAllowDisconnectComputer", m_bAllowDisconnectComputer
    PropertyBag.WriteProperty "m_bAllowCreateQueue", m_bAllowCreateQueue
    PropertyBag.WriteProperty "m_bAllowDeleteQueue", m_bAllowDeleteQueue
    PropertyBag.WriteProperty "m_bAllowStartQueue", m_bAllowStartQueue
    PropertyBag.WriteProperty "m_bAllowStallQueue", m_bAllowStallQueue
    PropertyBag.WriteProperty "m_bAllowDeleteQueueEvent", m_bAllowDeleteQueueEvent
    PropertyBag.WriteProperty "m_bAllowCreateQueueEvent", m_bAllowCreateQueueEvent
    PropertyBag.WriteProperty "m_bAllowStartComponent", m_bAllowStartComponent
    PropertyBag.WriteProperty "m_bAllowStallComponent", m_bAllowStallComponent

End Sub

Public Sub ReadProperties(ByVal PropertyBag As SnapInLib.PropertyBag)

    ' Account
    m_eAccount = PropertyBag.ReadProperty("m_eAccount", Account_CurrentUser)
    m_strDomain = PropertyBag.ReadProperty("m_strDomain", "")
    m_strUserName = PropertyBag.ReadProperty("m_strUserName", "")
    m_strPassword = PropertyBag.ReadProperty("m_strPassword", "")

    ' Permissions
    m_bAllowStartService = PropertyBag.ReadProperty("m_bAllowStartService", True)
    m_bAllowStopService = PropertyBag.ReadProperty("m_bAllowStopService", True)
    m_bAllowConnectComputer = PropertyBag.ReadProperty("m_bAllowConnectComputer", True)
    m_bAllowDisconnectComputer = PropertyBag.ReadProperty("m_bAllowDisconnectComputer", True)
    m_bAllowCreateQueue = PropertyBag.ReadProperty("m_bAllowCreateQueue", True)
    m_bAllowDeleteQueue = PropertyBag.ReadProperty("m_bAllowDeleteQueue", True)
    m_bAllowStartQueue = PropertyBag.ReadProperty("m_bAllowStartQueue", True)
    m_bAllowStallQueue = PropertyBag.ReadProperty("m_bAllowStallQueue", True)
    m_bAllowDeleteQueueEvent = PropertyBag.ReadProperty("m_bAllowDeleteQueueEvent", True)
    m_bAllowCreateQueueEvent = PropertyBag.ReadProperty("m_bAllowCreateQueueEvent", True)
    m_bAllowStartComponent = PropertyBag.ReadProperty("m_bAllowStartComponent", True)
    m_bAllowStallComponent = PropertyBag.ReadProperty("m_bAllowStallComponent", True)

End Sub

Public Sub Copy(datappgES As datappgES)
    
    ' m_SnapInData not copied
    ' m_ScopeItem not copied
    
    Set m_ppgAccount = Nothing
    Set m_ppgPermit = Nothing
    m_strLastPageDone = ""
    
    m_eAccount = datappgES.m_eAccount
    m_strDomain = datappgES.m_strDomain
    m_strUserName = datappgES.m_strUserName
    m_strPassword = datappgES.m_strPassword
    
    ' Permissions
    m_bAllowStartService = datappgES.m_bAllowStartService
    m_bAllowStopService = datappgES.m_bAllowStopService
    m_bAllowConnectComputer = datappgES.m_bAllowConnectComputer
    m_bAllowDisconnectComputer = datappgES.m_bAllowDisconnectComputer
    m_bAllowCreateQueue = datappgES.m_bAllowCreateQueue
    m_bAllowDeleteQueue = datappgES.m_bAllowDeleteQueue
    m_bAllowStartQueue = datappgES.m_bAllowStartQueue
    m_bAllowStallQueue = datappgES.m_bAllowStallQueue
    m_bAllowDeleteQueueEvent = datappgES.m_bAllowDeleteQueueEvent
    m_bAllowCreateQueueEvent = datappgES.m_bAllowCreateQueueEvent
    m_bAllowStartComponent = datappgES.m_bAllowStartComponent
    m_bAllowStallComponent = datappgES.m_bAllowStallComponent

End Sub

Public Sub PropertyPageDone(ByVal strLastPageDone As String)
    
    m_ScopeItem.PropertyChanged Me
    m_strLastPageDone = strLastPageDone

End Sub

Private Sub Class_Initialize()

    m_eAccount = Account_CurrentUser
    ' m_strDomain not initialised
    ' m_strUserName not initialised
    ' m_strPassword not initialised

    ' Permissions
    m_bAllowStartService = True
    m_bAllowStopService = True
    m_bAllowConnectComputer = True
    m_bAllowDisconnectComputer = True
    m_bAllowCreateQueue = True
    m_bAllowDeleteQueue = True
    m_bAllowStartQueue = True
    m_bAllowStallQueue = True
    m_bAllowDeleteQueueEvent = True
    m_bAllowCreateQueueEvent = True
    m_bAllowStartComponent = True
    m_bAllowStallComponent = True

End Sub
