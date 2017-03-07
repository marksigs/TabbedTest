VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrgUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : OrgUser
' Description   : Class which handles the frmEditUser to OSG interface functions.
'
' Change history
' Prog      Date        Description
' CL        11/07/01    Added Class
' STB       28/11/2001  SYS2912 - Removed the frmEditUser screen and
'                       ported its functionality to this screen.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private m_clsOSGInterface As OSGInterface

Private Const USERID = 0
Private Const USERFORENAME = 4
Private Const USERSURNAME = 5


Public Sub CallOSG_SaveUserMortgageAdminDetails()

    On Error GoTo Failed
         
    Dim SaveUserSystemDetails As InputSaveUserMortgageAdminDetails
    Dim SaveUserSubSystemDetails As InputSaveSubSystemDetails
    
    Set m_clsOSGInterface = New OSGInterface
             
    'Get the USER information from the screen
     
    'Get the text controls
    SaveUserSystemDetails.sUserID = frmEditUser.txtUserDetails(USERID).Text
    SaveUserSystemDetails.sUserName = frmEditUser.txtUserDetails(USERFORENAME).Text & " " & frmEditUser.txtUserDetails(USERSURNAME).Text
    
    'Setup empty strings for all the other USERDETAILS
    SaveUserSystemDetails.sTelephoneNumber = ""
    SaveUserSystemDetails.sInitialMenu = ""
    SaveUserSystemDetails.sExtensionNumber = ""
    SaveUserSystemDetails.sDepartment = ""
    SaveUserSystemDetails.sJobtitle = ""
    SaveUserSystemDetails.sLanguagePreference = ""
    SaveUserSystemDetails.sBranchNumber = ""
           
    Set SaveUserSystemDetails.colSubSystems = New Collection
    
    'Setup empty strings for the SUBSYSTEM
    SaveUserSubSystemDetails.sSubSystem = ""
    SaveUserSubSystemDetails.sUserAuthorityLevel = ""
    SaveUserSubSystemDetails.sAccessOtherBranches = ""
    SaveUserSubSystemDetails.sAccessStaffAccounts = ""
    SaveUserSubSystemDetails.sMaximumRequestSchedule = ""
    SaveUserSubSystemDetails.sBackDatedReporting = ""
    SaveUserSubSystemDetails.sGroupProfileName = ""
    SaveUserSystemDetails.colSubSystems.Add SaveUserSubSystemDetails
    
    'Call middle tier method SaveMortgageAdminDetails
    m_clsOSGInterface.SaveUserMortgageAdminDetails SaveUserSystemDetails
         
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
     
End Sub
