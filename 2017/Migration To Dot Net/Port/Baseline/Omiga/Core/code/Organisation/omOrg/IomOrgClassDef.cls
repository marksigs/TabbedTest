VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOmOrgClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IomOrgClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Interfaces to ClassDef for omCm.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date            Description
' MCS    14/01/2000     Created for performance rework.
' MV     21/11/2000     CORE000011 : Added LoadDistributionChannelData ,LoadDepartmentData,LoadUnitData,
'                                    LoadUserData Functions
' DJP    18/12/2000     Added UserData enumeration type
' PSC    12/01/01       SYS1753 LoadMortgageLenderDirectoryData
' JR     04/09/01       Omiplus24 Added LoadContactTelephoneData
'-------------------------------------------------------------------------------------------
'BBG
'
'AW     21/04/04    WP06 - BBG561   Added LoadLenderNameAddressData, LoadLenderDirectoryData
'                                       LoadLenderAddressData, LoadLenderContactData
'-------------------------------------------------------------------------------------------
Public Enum MortgageLenderData
    
    LenderNameAndIncomeMultiples
    MortgageLenderDefault
End Enum
Public Enum DEPARTMENTDATA
    ddDefault
    ddFindCurrentUnitList
End Enum
Public Enum UserData
    udDefault
    udFindUserList
End Enum
Public Enum UNITSUMMARY
    usDefault
    usFindUnitSummary
End Enum
Public Enum CONTACTDETAILS
    cdDefault
    cdFindUnitSummary
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History :
'
'Prog   Date        AQR         Description
'AW     13/08/2002  BM029       Added LoadUnitChequeNo
'MV     04/09/2002  BMIDS00353  Added LoanMPMIGRATEBAND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadLenderNamesData() As FreeThreadedDOMDocument40
End Function
Public Function LoadMortgageLenderData(Optional ByVal vemcMortgageLenderData As _
                                    MortgageLenderData = MortgageLenderDefault) As FreeThreadedDOMDocument40
End Function
Public Function LoadActionOwnerData() As FreeThreadedDOMDocument40
End Function
Public Function LoadCompetencyData() As FreeThreadedDOMDocument40
End Function
Public Function LoadMortgageLenderParametersData() As FreeThreadedDOMDocument40
End Function
Public Function LoadLegalFeeData() As FreeThreadedDOMDocument40
End Function
Public Function LoadMigRateData() As FreeThreadedDOMDocument40
End Function
Public Function LoadMPMigRateData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUnitNameData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUnitData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUserNameData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUserData(Optional udUserData As UserData = udDefault) As FreeThreadedDOMDocument40
End Function
Public Function LoadDistributionChannelData() As FreeThreadedDOMDocument40
End Function
Public Function LoadDepartmentData( _
    Optional ByVal vDDDataView As DEPARTMENTDATA = ddDefault) As FreeThreadedDOMDocument40
End Function
Public Function LoadUserRoleData() As FreeThreadedDOMDocument40
End Function
Public Function LoadOrganisationUserData() As FreeThreadedDOMDocument40
End Function
Public Function LoadIntermediaryUserData() As FreeThreadedDOMDocument40
End Function
Public Function LoadIntermediaryIndividualData() As FreeThreadedDOMDocument40
End Function
Public Function LoadQualificationData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUserCompetencyData() As FreeThreadedDOMDocument40
End Function
Public Function LoadPasswordData() As FreeThreadedDOMDocument40
End Function
Public Function LoadWorkingHoursData() As FreeThreadedDOMDocument40
End Function
Public Function LoadMortgageLenderDirectoryData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUnitSummaryData(Optional ByVal vemcUNITSUMMARY As _
                                    UNITSUMMARY = usDefault) As FreeThreadedDOMDocument40
End Function
Public Function LoadContactDetailsData(Optional ByVal vemcCONTACTDETAILS As _
                                    CONTACTDETAILS = cdDefault) As FreeThreadedDOMDocument40
End Function
Public Function LoadContactTelephoneData() As FreeThreadedDOMDocument40
End Function
Public Function LoadAllowableIncomeFactorsData() As FreeThreadedDOMDocument40
End Function
Public Function LoadUnitChequeNo() As FreeThreadedDOMDocument40
End Function
Public Function LoadLenderContactData() As FreeThreadedDOMDocument40
End Function
Public Function LoadLenderAddressData() As FreeThreadedDOMDocument40
End Function
Public Function LoadLenderDirectoryData() As FreeThreadedDOMDocument40
End Function
Public Function LoadLenderNameAddressDirectoryData() As FreeThreadedDOMDocument40
End Function
