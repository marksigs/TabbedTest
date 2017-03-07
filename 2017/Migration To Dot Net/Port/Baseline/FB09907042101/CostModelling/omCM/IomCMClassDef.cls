VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomCMClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IomCMClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Interfaces to ClassDef for omCm.
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS   14/12/99    Created for performance rework.
' JLD   17/04/00    Added LoadBCProductData
' MS    04/07/00    SYS1048 BCSubQuoteNumber enumeration added to QuotationData
' SR    16/08/2004  BMIDS815 - Modified Enum LoanComponentData
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'ING History:
'
'Prog   Date        Description
'MV     14/09/2005  MARS35 - New Function LoadGetAllQuotationsForAppViewData()
'------------------------------------------------------------------------------------------
Option Explicit
Public Enum BCSubQuoteData
    
    BCSubQuoteView
    BCSubQuoteDefault
    
End Enum
Public Enum PPSubQuoteData
    
    PPSubQuoteView
    PPSubQuoteDefault
    
End Enum
Public Enum LifeSubQuoteData
    
    LifeSubQuoteView
    LifeSubQuoteDefault
    
End Enum
Public Enum MortgageSubQuoteData
    
    MortgageSubQuoteView
    MortgageSubQuoteDefault
    MortgageSubQuoteAmountRequested
'    MortgageKeyView
End Enum
Public Enum QuotationData
    
    QuotationView
    QuotationKeyView
    QuotationDefault
    BCSubQuoteNumber
    
End Enum
Public Enum MortgageIncentiveData
    
    MortgageIncentiveTotal
    MortgageIncentiveDefault
    MortgageIncentiveTotalReturn
    
End Enum

Public Enum LoanComponentData
    
    LoanComponentDefault
    LoanComponentView
    LoanComponentLoanAmount
    LoanComponentStoredQuoteView
    QuotationSummaryView
    LoanComponentPortingData  ' SR 16/08/2004 : BMIDS815
End Enum

Public Enum LifeBenefitData
    LifeBenefitDataDefault
    QuotationSummary
End Enum




Public Function LoadGetAllQuotationsForAppViewData() As FreeThreadedDOMDocument40

End Function

Public Function LoadLoanComponentRedemptionFeeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoanComponentData(Optional ByVal vemcLoanComponentData As _
                                    LoanComponentData = LoanComponentDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadQuotationData(Optional ByVal vemcQuotationData As _
                                    QuotationData = QuotationDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadPPSubQuoteData(Optional ByVal vemcPPSubQuoteData As _
                                    PPSubQuoteData = PPSubQuoteDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadBCSubQuoteData(Optional ByVal vemcBCSubQuoteData As _
                                    BCSubQuoteData = BCSubQuoteDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageSubQuoteData(Optional ByVal vemcMortgageSubQuoteData As _
                                    MortgageSubQuoteData = MortgageSubQuoteDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadLifeSubQuoteData(Optional ByVal vemcLifeSubQuoteData As _
                                    LifeSubQuoteData = LifeSubQuoteDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadPPDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLifeBenefitData(Optional ByVal velbdLifeBenefitData As _
                                    LifeBenefitData = LifeBenefitDataDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageIncentiveData(Optional ByVal vemcMortgageIncentive As _
                                    MortgageIncentiveData = MortgageIncentiveDefault) As FreeThreadedDOMDocument40
End Function

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be removed due to OOSS
'Public Function LoadLoanComponentResultsData() As FreeThreadedDOMDocument40
'End Function
'BMIDS736 End
'BMIDS782 End

Public Function LoadNonPanelMortgageProductData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageOneOffCostData() As FreeThreadedDOMDocument40
End Function

Public Function LoadPortedLenderData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBCDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBCNotesData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBCProductData() As FreeThreadedDOMDocument40
End Function

Public Function LoadBCValuablesOLimitData() As FreeThreadedDOMDocument40
End Function

Public Function LoadMortgageCalculatorDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadStoredQuoteFindListViewData() As FreeThreadedDOMDocument40
End Function

Public Function LoadQuotationSummaryViewData() As FreeThreadedDOMDocument40
End Function

Public Function LoadActiveLoanComponentsViewData() As FreeThreadedDOMDocument40
End Function

Public Function LoadQuotationLoanSummaryViewData() As FreeThreadedDOMDocument40
End Function

'BMIDS736 GHun 23/02/2004
Public Function LoadLoanComponentPaymentScheduleData() As FreeThreadedDOMDocument40
End Function

Public Function LoadLoanComponentBalanceScheduleData() As FreeThreadedDOMDocument40
End Function
'BMIDS736 End
