VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INewPropertyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      INewPropertyBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' BG    26/05/00    Added SaveNewPropertyAndAddress
' SR    02/06/00    Added 'SaveNewPropertyAndAddress'
' IVW   06/06/00    Added GetNewPropertyAndAddressDetails
' SR    13/06/00    New methods 'GetFullNewPropertyDetails' and 'SaveNewProperty'
' MC    13/07/00    New method SaveNewPropertyAddress
'------------------------------------------------------------------------------------------

Option Explicit
Public Enum INewPropertyBOMethod
    bomSaveNewPropertyAndAddress
    bomSaveNewPropertyAndDeposit
End Enum

'BMIDS History:
'
' Prog  Date        Description
' GD    04/07/02    BMIDS00165 - Added CalcExcessRentalIncome
' GHun  19/11/02    BMIDS00563 - Added GetNewPropertyAddress
' JD    22/06/04    BMIDS765 CC076 added GetRentalIncomeRates
' GHun  23/09/04    BMIDS887 Added SaveNewPropertyAndUpdateFactFind
'------------------------------------------------------------------------------------------

Public Function GetNewPropertyAndAddressDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewPropertyGeneral( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetValuationTypeAndLocation( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNewPropertyDescription( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetFullNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAndDeposit( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAndAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function Validate(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vbomMethod As INewPropertyBOMethod) As IXMLDOMNode
End Function

Public Function GetPropertyUnknownDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'GD BMIDS00165 - New Method
Public Function CalcExcessRentalIncome( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

'BMIDS00563
Public Function GetNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS00563 End
Public Function GetRentalIncomeRates(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

'BMIDS887 GHun
Public Function SaveNewPropertyAndUpdateFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'BMIDS887 End
