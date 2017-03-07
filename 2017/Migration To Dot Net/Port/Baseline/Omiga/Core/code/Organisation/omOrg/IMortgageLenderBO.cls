VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgageLenderBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IMortgageLenderBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MV    24/11/00    Created
' PSC   12/01/01    SYS1753 Added GetMainLenderDirectoryGuid
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'Baseline specific History:
'
' Prog  Date        Description
' MV    20/11/03    Core6 - Added GetPanelLenderDetails
' AW    14/09/04    BBG1416 - Added GetLenderTelephoneDetails
' AW    14/09/04    BBG1436 - Added GetLenderNameDirectoryDetails
' SAB   20/09/2005  CORE155 - Added GetApplicationMortgageLender
'------------------------------------------------------------------------------------------
Option Explicit
Public Function FindLenderNames(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetPanelLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindLenderNameAndIncomeMultiples(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetMainLenderDirectoryGuid(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function FindAllowableIncomeFactors(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetLenderTelephoneDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetLenderNameDirectoryDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
Public Function GetApplicationMortgageLender(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function



