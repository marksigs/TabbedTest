VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrganisationBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IOrganisationBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     21/11/00    CORE000011 - Added FindDistributionChannel, FindDepartmentList ,FindUnitList ,
'                   FindUserList Functions
'SAB    21/09/05    CORE191 - Added IsDateWorkingDay
'JD     02/02/2006  CORE237 - Added GetUserCompetency
'------------------------------------------------------------------------------------------
Option Explicit
Public Function ValidateUserMandateLevel(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUnitList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindDepartmentList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUserList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindDistributionChannelList() As IXMLDOMNode
End Function
Public Function FindUserNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUnitNameList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindUnitSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetActionOwnerDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function IsDateWorkingDay(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetUserCompetency(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
