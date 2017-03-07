VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomMPClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      IomMPClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Interfaces to ClassDef for omCm.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS    12/01/2000    Created for performance rework.
' MC    06/09/2001  SYS2250 - BaseRate table & BaseRateBand amendments
' JLD   16/10/2001  SYS2250 - temporarily backed out optional parameter to LoadBaseRateBandData
' MC    06/12/2001  SYS2250 - Reinstate BaseRate table & BaseRateBand amendments
'------------------------------------------------------------------------------------------

Public Enum MortProdLanguageData
    QuotationSumaryView
    MortProdLanguageDefault
End Enum
Public Enum MortProdParametersData
    ValidateCompulsoryProducts
    MortProdParametersDefault
End Enum
Public Enum BaseRateOption
    opGetRateDifference
    opGetAbsoluteRate
End Enum


'------------------------------------------------------------------------------------------
'BMIDS History:
'
' DPF   15/07/2002  CMWP3   -   Have added a new method called LoadRedemptionFeeBandData()
' GD    10/09/02    BMIDS00313 - APWP2 - BM004 Added LoadConditionsData
' MDC   01/11/2002  BMIDS00654 Added LoadMortgageProductIncomeMultiplierSet
' JD    21/06/04    BMIDS765 CC076 added LoadRentalIncomeRateData
' HMA   13/12/04    BMIDS959 Remove MortgageProductBands
'------------------------------------------------------------------------------------------
Public Function LoadProductNamesData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductData() As FreeThreadedDOMDocument40
End Function

Public Function LoadReservedProductData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductSelectionData() As FreeThreadedDOMDocument40
End Function

Public Function LoadQuotationSummaryViewData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductParametersData(Optional ByVal vempMortProdParametersData As _
                                    MortProdParametersData = MortProdParametersDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductLanguageData() As FreeThreadedDOMDocument40
End Function

'BMIDS959  Remove LoadMortgageProductBandsData

Public Function LoadBaseRateBandData(Optional ByVal vRateOption As BaseRateOption = opGetRateDifference) As FreeThreadedDOMDocument40
End Function

Public Function LoadSpecialGroupData() As FreeThreadedDOMDocument40
End Function
Public Function LoadRentalIncomeRateData() As FreeThreadedDOMDocument40
End Function

Public Function LoadOtherFeeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadAdministrationFeeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadValuationFeeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductPurposeOfLoanData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageProductApplicationTypeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadInterestRateTypeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadIncentiveData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBaseRateData() As FreeThreadedDOMDocument40
End Function

Public Function LoadRedemptionFeeBandData() As FreeThreadedDOMDocument40
End Function

Public Function LoadConditionsData() As FreeThreadedDOMDocument40
End Function

'BMIDS00654 MDC 01/11/2002
Public Function LoadMortgageProductIncomeMultiplierSet() As FreeThreadedDOMDocument40
End Function
'BMIDS00654 MDC 01/11/2002 - End

