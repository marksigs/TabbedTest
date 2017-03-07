VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrganisationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IOrganisationDO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     21/11/00    CORE000011 - Added FindDistributionChannel, FindDepartmentList ,FindUnitList ,
'                   FindUserList Functions
'SAB    23/09/05    CORE191 - Added IsDateBankHoliday
'------------------------------------------------------------------------------------------
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'
'   AW     13/08/2002  BM029 - Added UpdateUnitChequeNo
'
Public Function ValidateUserMandateLevel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindCurrentUnitList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetOmigaUserDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindCurrentQualificationList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetCurrentCompetency(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetWorkingHoursForToday(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetCurrentPassword(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Sub CreatePassword(ByVal vxmlRequest As IXMLDOMElement)
End Sub
Public Function FindPasswordList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetCurrentOmigaUser(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetActionOwnerDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function IsDateBankHoliday(ByVal vstrDistributionChannelId As String, ByVal vdtmBankHoliday As Date) As Boolean
End Function
Public Function IsBankHoliday(ByVal vstrDistributionChannelId As String) As Boolean
End Function
Public Sub UpdateChangePasswordIndicator(ByVal vxmlRequest As IXMLDOMElement)
End Sub
Public Function FindUnitList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindDepartmentList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindDistributionChannelList() As IXMLDOMNode
End Function
Public Function FindUserList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUserNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUnitNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUnitSummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
                        
End Function
Public Sub UpdateUnitChequeNo(ByVal vxmlRequest As IXMLDOMElement)
End Sub
