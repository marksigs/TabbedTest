VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSGGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type InputGetUserMortgageAdminDetails
    sUserID As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
End Type

Public Type OutputGetUserMortgageAdminDetails
    sUserID As String
    sUserName As String
    sJobtitle As String
    sDepartment As String
    sBranchNumber As String
    sLanguagePreference As String
    sInitialMenu As String
    sTelephoneNumber As String
    sExtensionNumber As String
    colSubSystems As Collection
End Type

Public Type OutputSubSystemDetails
    sSubSystem As String
    sUserAuthorityLevel As String
    sGroupProfileName As String
    sAccessOtherBranches As String
    sAccessStaffAccounts As String
    sMaximumRequestschedule As String
    sBackDatedReporting As String
End Type

Public Type InputSaveUserMortgageAdminDetails
    sUserID As String
    sUserName As String
    sJobtitle As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
    sDepartment As String
    sBranchNumber As String
    sLanguagePreference As String
    sInitialMenu As String
    sTelephoneNumber As String
    sExtensionNumber As String
    colSubSystems As Collection
End Type

Public Type InputSaveSubSystemDetails
    sSubSystem As String
    sUserAuthorityLevel As String
    sGroupProfileName As String
    sAccessOtherBranches As String
    sAccessStaffAccounts As String
    sMaximumRequestschedule As String
    sBackDatedReporting As String
End Type

Public Type InputGetUserSubSystemDetails
    sUserID As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
End Type




