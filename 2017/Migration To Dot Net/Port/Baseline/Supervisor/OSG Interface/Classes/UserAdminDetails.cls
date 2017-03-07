VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserAdminDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : UserAdminDetails
' Description   : A class used to store a users Optimus details which can be
'                 passed between forms and support classes in supervisor.
' History       :
'
' Prog  Date        Description
' STB   16-May-2002 SYS4609 Added class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Properties held by this class.
Private m_sUserID As String
Private m_sUserName As String
Private m_sJobtitle As String
Private m_sDepartment As String
Private m_sBranchNumber As String
Private m_sInitialMenu As String
Private m_sTelephoneNumber As String
Private m_sExtensionNumber As String
Private m_sLanguagePreference As String

'A list of subsystem objects for this user.
Private m_colSubSystemList As Collection


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : SubSystemList
' Description : Return a reference to the list of sub-systems.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get SubSystemList() As Collection
    Set SubSystemList = m_colSubSystemList
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : UserID
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get UserID() As String
    UserID = m_sUserID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : UserID
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let UserID(ByVal sUserID As String)
    m_sUserID = sUserID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Username
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Username() As String
    Username = m_sUserName
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Username
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Username(ByVal sUsername As String)
    m_sUserName = sUsername
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : JobTitle
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get JobTitle() As String
    JobTitle = m_sJobtitle
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : JobTitle
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let JobTitle(ByVal sJobTitle As String)
    m_sJobtitle = sJobTitle
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Department
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Department() As String
    Department = m_sDepartment
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Department
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let Department(ByVal sDepartment As String)
    m_sDepartment = sDepartment
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : BranchNumber
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BranchNumber() As String
    BranchNumber = m_sBranchNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : BranchNumber
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let BranchNumber(ByVal sBranchNumber As String)
    m_sBranchNumber = sBranchNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : InitialMenu
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get InitialMenu() As String
    InitialMenu = m_sInitialMenu
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : InitialMenu
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let InitialMenu(ByVal sInitialMenu As String)
    m_sInitialMenu = sInitialMenu
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : TelephoneNumber
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get TelephoneNumber() As String
    TelephoneNumber = m_sTelephoneNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : TelephoneNumber
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let TelephoneNumber(ByVal sTelephoneNumber As String)
    m_sTelephoneNumber = sTelephoneNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : ExtensionNumber
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ExtensionNumber() As String
    ExtensionNumber = m_sExtensionNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : ExtensionNumber
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let ExtensionNumber(ByVal sExtensionNumber As String)
    m_sExtensionNumber = sExtensionNumber
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : LanguagePreference
' Description : Property Get.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LanguagePreference() As String
    LanguagePreference = m_sLanguagePreference
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : LangagePreference
' Description : Property Let.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let LanguagePreference(ByVal sLanguagePreference As String)
    m_sLanguagePreference = sLanguagePreference
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Class_Initialize
' Description : Create the sub-system collection class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_colSubSystemList = New Collection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : Class_Terminate
' Description : Release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Set m_colSubSystemList = Nothing
End Sub
