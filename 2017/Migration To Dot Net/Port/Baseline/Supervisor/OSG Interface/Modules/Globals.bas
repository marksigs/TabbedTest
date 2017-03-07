Attribute VB_Name = "Globals"
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
    sExtensionNumber As String
    sJobtitle As String
    sDepartment As String
    sBranchNumber As String
    sLanguagePreference As String
    sInitialMenu As String
    sTelephoneNumber As String
    sCountryCode As String
    sAreaCode As String
    colSubSystems As Collection
End Type

Public Type OutputSubSystemDetails
    sSubSystem As String
    sAuthorityLevel As String
End Type


Public Type InputSaveUserMortgageAdminDetails
    sUserID As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
    sUserName As String
    sJobtitle As String
    sDepartment As String
    sBranchNumber As String
    sLanguagePreference As String
    sInitialMenu As String
    sTelephoneNumber As String
    sCountryCode As String
    sAreaCode As String
    sExtensionNumber As String
End Type

Public Type InputGetUserSubSystemDetails
    sUserID As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
    sSubSystem As String
    sGroupProfileName As String
    sUserAuthorityLevel As String
    sAccessOtherBranches As String
    sAccessStaffAccounts As String
    sMaximumRequestLevel As String
    sBackDatedReporting As String
End Type

Public Type InputSaveUserSubSystemDetails
    sUserID As String
    sUnitID As String
    sMachineID As String
    sChannelID As String
    sSubSystem As String
    sGroupProfileName As String
    sUserAuthorityLevel As String
    sAccessOtherBranches As String
    sAccessStaffAccounts As String
    sMaximumRequestLevel As String
    sBackDatedReporting As String
End Type

