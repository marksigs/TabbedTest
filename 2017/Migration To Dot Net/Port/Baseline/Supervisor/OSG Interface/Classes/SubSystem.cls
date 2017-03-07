VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : Subsystem
' Description   : A class used to store a users Optimus details which can be
'                 passed between forms and support classes in supervisor.
' History       :
'
' Prog  Date        Description
' STB   16-May-2002 SYS4609 Added class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Private m_sSubSystem As String
Private m_sGroupProfileName As String
Private m_sUserAuthorityLevel As String
Private m_sBackDatedReporting As String
Private m_sAccessOtherBranches As String
Private m_sAccessStaffAccounts As String
Private m_sMaximumRequestSchedule As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : SubSystem
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get SubSystem() As String
    SubSystem = m_sSubSystem
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : SubSystem
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let SubSystem(ByVal sSubSystem As String)
    m_sSubSystem = sSubSystem
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : GroupProfileName
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get GroupProfileName() As String
    GroupProfileName = m_sGroupProfileName
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : GroupProfileName
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let GroupProfileName(ByVal sGroupProfileName As String)
    m_sGroupProfileName = sGroupProfileName
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : UserAuthorityLevel
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get UserAuthorityLevel() As String
    UserAuthorityLevel = m_sUserAuthorityLevel
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : UserAuthorityLevel
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let UserAuthorityLevel(ByVal sUserAuthorityLevel As String)
    m_sUserAuthorityLevel = sUserAuthorityLevel
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : BackDatedReporting
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BackDatedReporting() As String
    BackDatedReporting = m_sBackDatedReporting
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : BackDatedReporting
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let BackDatedReporting(ByVal sBackDatedReporting As String)
    m_sBackDatedReporting = sBackDatedReporting
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : AccessOtherBranches
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get AccessOtherBranches() As String
    AccessOtherBranches = m_sAccessOtherBranches
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : AccessOtherBranches
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let AccessOtherBranches(ByVal sAccessOtherBranches As String)
    m_sAccessOtherBranches = sAccessOtherBranches
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : AccessStaffAccounts
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get AccessStaffAccounts() As String
    AccessStaffAccounts = m_sAccessStaffAccounts
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : AccessStaffAccounts
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let AccessStaffAccounts(ByVal sAccessStaffAccounts As String)
    m_sAccessStaffAccounts = sAccessStaffAccounts
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : MaximumRequestSchedule
' Description : Property get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get MaximumRequestSchedule() As String
    MaximumRequestSchedule = m_sMaximumRequestSchedule
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : MaximumRequestSchedule
' Description : Property let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let MaximumRequestSchedule(ByVal sMaximumRequestSchedule As String)
    m_sMaximumRequestSchedule = sMaximumRequestSchedule
End Property
