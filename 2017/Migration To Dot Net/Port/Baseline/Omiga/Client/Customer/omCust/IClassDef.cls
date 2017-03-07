VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomCustClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definitions interface for omCust
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     17/04/00    Added Enumeration for CustomerVersionClassDef
'MC     01/06/00    Added Tax Details Option to enum
'HMA    09/09/04    BMIDS864  Added LoadVerificationData
'----------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    28/06/2005  MAR5 - Added LoadCustomerAddressAndRoleData and LoadCustomerNameData
'HM     08/09/2005  MAR29   LoadWrapUpDetailsData was added
'HMA    20/10/2005  MAR23  Added LoadKYCData. Changed LoadCustomerData.
'------------------------------------------------------------------------------------------

Option Explicit
Public Enum omCustDataDefType
    omCustVersionDefault
    omCustVersionDOB
    omCustVersionTax
    omCustVersionNumber
End Enum

Public Function LoadAliasData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerTelephoneData() As FreeThreadedDOMDocument40
End Function

Public Function LoadWrapUpDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerAddressData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAliasPersonData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCurrentPropertyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadHomeInsuranceData() As FreeThreadedDOMDocument40
End Function

Public Function LoadTenancyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerData(Optional ByVal intCustDataDefType As _
                            omCustDataDefType = omCustVersionDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerVersionData(Optional ByVal intCustVersionDataDefType As _
                            omCustDataDefType = omCustVersionDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadDependantDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadDependantData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAreasOfInterestData() As FreeThreadedDOMDocument40

End Function

Public Function LoadDefinition(ByVal strTableName As String) As FreeThreadedDOMDocument40

End Function

Public Function LoadPropertyInsuranceData() As FreeThreadedDOMDocument40

End Function

Public Function LoadContactHistoryData() As FreeThreadedDOMDocument40

End Function

'BMIDS00005 CRWP2 BM065 GHun 30/05/2002
Public Function LoadApplicationContactData() As FreeThreadedDOMDocument40
End Function
'BMIDS00005 End

'BMIDS00026 IWP2 BM060 GHun 07/06/2002
Public Function LoadApplicationCustomerData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerPhoneData() As FreeThreadedDOMDocument40
End Function

Public Function LoadCustomerVerificationData() As FreeThreadedDOMDocument40
End Function
'BMIDS00026 End

'BMIDS00205 GHun 15/07/2002
Public Function LoadCRSCustomerAddressData() As FreeThreadedDOMDocument40
End Function
'BMIDS00205 End

'BMIDS00443 GHun 04/10/2002
Public Function LoadLatestCustomerVersionData() As FreeThreadedDOMDocument40
End Function
'BMIDS00443 End

'BMIDS758
Public Function LoadRemovedToECustomer() As FreeThreadedDOMDocument40
End Function

'BMIDS864
Public Function LoadVerificationData() As FreeThreadedDOMDocument40
End Function

'PSC 28/06/2005 MAR5 - Start
Public Function LoadCustomerAddressAndRoleData() As FreeThreadedDOMDocument40
End Function
Public Function LoadCustomerNameData() As FreeThreadedDOMDocument40
End Function
'PSC 28/06/2005 MAR5 - End

'MAR23
Public Function LoadKYCData() As FreeThreadedDOMDocument40
End Function
