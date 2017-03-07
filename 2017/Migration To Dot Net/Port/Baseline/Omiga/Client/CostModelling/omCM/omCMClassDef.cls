VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omCMClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCMClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   XML definitions for Class/Tables for omCM
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS   14/12/99    Created for performance rework.
' JLD   08/02/2000  Altered LoadBCDetailsData - changed NUMBEROFCLAIMSLAST5YEARS
' AY    21/03/00    QuotationSummary structures altered
' JLD   28/03/00    Added LoanComponentView to LoadLoanComponentData
' JLD   17/04/00    Added LoadBCProductData
' JLD   05/05/00    SYS0690 - change name of MORTGAGEINCENTIVESEQUENCENUMBE
' JLD   12/05/00    Added TotalLoanComponentAmount to the LoanComponentView of LoadLoanComponentData
' JLD   18/05/00    SYS0725 - combo groupname for lifecover is LifeCoverType
' MS    04/07/00    SYS1048 - BCSubQuoteNumber enumeration added to QuotationData
' BG    21/08/00    SYS0964 added FREQUENCY Tag to LoadBCSubQuoteData ClassDef for use in CM010
' JLD   27/09/00    SYS1573 - changed datatype of STARTDATE on mortgageproduct table to dateTime from date
' PSC   05/03/01    SYS1879 - Add LoadQuotationLoanSummaryData
' SR    06/09/01    SYS2412 - modified 'LoadMortgageIncentiveData'
' PSC   17/01/02    SYS3801 - modified 'LoadMortgageIncentiveData'
' DM 25/03/02 SYS4313   - Added ADMINSUBLEDGERCODE to schema.
' SG    28/05/02    SYS4767 - MSMS to Core integration
'-------------------------------------------------------------------------------------------
'MARS History:
'
' Prog  Date        Description
' PSC   28/06/2005  MAR5 - Amended IomCMClassDef_LoadQuotationData()
' MV    14/09/2005  MAR35 - Added IomCMClassDef_LoadGetAllQuotationsForAppViewData()
' MV    14/09/2005  MAR35 - Amended IomCMClassDef_LoadGetAllQuotationsForAppViewData(),IomCMClassDef_LoadQuotationData()
' AM    14/09/2005  MAR29 - Ammended IomCMClassDef_LoadMortgageSubQuoteData
' MV    21/09/2005  MAR35 - Amended IomCMClassDef_LoadStoredQuoteFindListViewData
' HMA   16/09/2005  MAR28 - Amended LoadMortgageOneOffCostData
' JD    29/09/2005  MAR30   Amended LoadMortgageSubQuoteData to get type for CONFIRMEDCALCULATEDINCMULTIPLIERTYPE correct
' GHun  06/10/2005  MAR46   Changed LoadLoanComponentData for porting
' TW    28/10/2005  MAR211  Changed LoadQuotationData
' JD    17/02/2006  MAR1040 changed LoadQuotationLoanSummaryViewData to add msq income multiples
' JD    08/03/2006  MAR1061 Added purchasepriceorestimatedvalue to LoadMortgageSubquoteData
'                           Added purchasepriceorestimatedvalue to LoadQuotationLoanSummaryViewData
'                           Added AmountRequested,TotalLoanAmount,PurchasePrice,LTV to LoadQuotationSummaryViewData
'                           Added AmountRequested,PurchasePrice,LTV to LoadStoredQuoteFindListViewData
'-------------------------------------------------------------------------------------------
'Epsom History:
'
' Prog  Date        Description
' PE    26/07/2006  EP1038  Mars merge/MAR1893 Removed redundant LIFESUBQUOTE columns
' AW    28/07/2006  EP1050  Removed BenefitType from LoadQuotationLoanSummaryViewData()
'-------------------------------------------------------------------------------------------
'Epsom 2 History:
'
' Prog  Date        Description
' GHun  20/12/2006  EP2_56  Added OriginalLTV to LoadLoanComponentData
' AShaw 29/12/2006  EP2_56  Added MortgageLoanGUID to LoadLoanComponentData.
' MAH   03/01/2007  EP2_444 Added RepaymentVehicleMonthlyCost to LoadLoanComponentData.
' PE    18/01/2007  EP2_696 Removed duplicate MortgageLoanGUID
' PSC   12/03/2007  EP2_1611 Add MORTGAGELOANGUID to LoadLoanComponentData for porting
' SR    21/03/2007  EP2_1753 - modified LoadMortgageOneffCostData - Add REBATEAMOUNT
'-------------------------------------------------------------------------------------------

Option Explicit

Implements IomCMClassDef
Private m_objXmlAssist As XMLAssist
Private Sub Class_Initialize()
    Set m_objXmlAssist = New XMLAssist
End Sub

Private Function IomCMClassDef_LoadGetAllQuotationsForAppViewData() As FreeThreadedDOMDocument40
        
    Const strFunctionName As String = "IomCMClassDef_LoadGetAllQuotationsForAppViewData"
    
    Dim strXML As String
    
    strXML = "<TABLENAME>QUOTATION" _
                & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                & "<OTHERS>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                & "<OTHERS>LASTOFFERED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "</TABLENAME>"
                '& "<OTHERS>ACCEPTEDAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                '& "<OTHERS>ACCEPTEDFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                '& "<OTHERS>QUICKQUOTEAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                '& "<OTHERS>QUICKQUOTEFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
                
                
    Set IomCMClassDef_LoadGetAllQuotationsForAppViewData = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)


End Function
Private Sub Class_Terminate()
    Set m_objXmlAssist = Nothing
End Sub
Public Function LoadMortgageSubQuoteData(Optional ByVal vemcMortgageSubQuoteData As _
                                            MortgageSubQuoteData = PPSubQuoteDefault) As String
'call the internal method here
'FIX ME modify paymentprotectionDO to use
'Dom method not public interface string method
    Dim objDomDoc As FreeThreadedDOMDocument40

    Set objDomDoc = IomCMClassDef_LoadMortgageSubQuoteData()

    LoadMortgageSubQuoteData = objDomDoc.xml

End Function

Public Function LoadPPSubQuoteData(Optional ByVal vemcPPSubQuoteData As _
                                        PPSubQuoteData = PPSubQuoteDefault) As String
'call the internal method here
'FIX ME modify paymentprotectionDO to use
'Dom method not public interface string method
    Dim objDomDoc As FreeThreadedDOMDocument40
    
    Set objDomDoc = IomCMClassDef_LoadPPSubQuoteData()
    
    LoadPPSubQuoteData = objDomDoc.xml
    
End Function

Public Function LoadPPDetailsData(Optional ByVal vemcPPSubQuoteData As _
                                        PPSubQuoteData = PPSubQuoteDefault) As String
'call the internal method here
'FIX ME modify paymentprotectionDO to use
'Dom method not public interface string method
    Dim objDomDoc As FreeThreadedDOMDocument40
    
    Set objDomDoc = IomCMClassDef_LoadPPDetailsData()
    
    LoadPPDetailsData = objDomDoc.xml
    
End Function

Public Function LoadLifeSubQuoteData(Optional ByVal vemcLifeSubQuoteData _
                                        As LifeSubQuoteData = LifeSubQuoteDefault) As String
'call the internal method here
'FIX ME modify paymentprotectionDO to use
'Dom method not public interface string method
    Dim objDomDoc As FreeThreadedDOMDocument40
    
    Set objDomDoc = IomCMClassDef_LoadLifeSubQuoteData()
    
    LoadLifeSubQuoteData = objDomDoc.xml
    
End Function

Public Function LoadLifeBenefitData() As String
'call the internal method here
'FIX ME modify paymentprotectionDO to use
'Dom method not public interface string method
    Dim objDomDoc As FreeThreadedDOMDocument40
    
    Set objDomDoc = IomCMClassDef_LoadLifeBenefitData()
    
    LoadLifeBenefitData = objDomDoc.xml
    
End Function

Private Function IomCMClassDef_LoadActiveLoanComponentsViewData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadActiveLoanComponentsViewData"

    Dim strXML As String
    
    strXML = "<TABLENAME>FINDACTIVELOANCOMPONENTS" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>STARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _
    & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPITALANDINTERESTELEMENT<TYPE>dbdtLong</TYPE></OTHERS>"
    
    'SG 28/05/02 SYS4767
    strXML = strXML & "<OTHERS>NETCAPANDINTELEMENT<TYPE>dbdtLong</TYPE></OTHERS>"

    strXML = strXML & "<OTHERS>GROSSMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTONLYELEMENT<TYPE>dbdtInt</TYPE></OTHERS>"
    
    'SG 28/05/02 SYS4767
    strXML = strXML & "<OTHERS>NETINTONLYELEMENT<TYPE>dbdtLong</TYPE></OTHERS>"
        
    strXML = strXML & "<OTHERS>LOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MIRASAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MIRASELIGIBILITY<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>NETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>PORTEDLOAN<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>REPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>RepaymentType</COMBO></OTHERS>" _
    & "<OTHERS>TERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TERMINYEARS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TOTALLOANCOMPONENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "</TABLENAME>"

'SW 26/06/2006 EP771 Start
'    & "<OTHERS>REPAYMENTVEHICLE<TYPE>dbdtComboId</TYPE>" _
'    & "<COMBO>RepaymentVehicle</COMBO></OTHERS>" _
'    & "</TABLENAME>"
'SW 26/06/2006 EP771 End


    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadActiveLoanComponentsViewData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadBCProductData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadBCProductData"

    Dim strXML As String
    
    strXML = "<TABLENAME>BUILDINGANDCONTENTSPRODUCT" _
    & "<PRIMARYKEY>BCPRODUCTNUMBER<TYPE>dbdtLong</TYPE></PRIMARYKEY>" _
    & "<OTHERS>VALUABLESLIMIT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>BCSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>ENDDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>INCLUDESBUILDINGSCOVER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>INCLUDESCONTENTSCOVER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>VALUABLESITEMLIMIT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadBCProductData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadMortgageCalculatorDetailsData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadMortgageCalculatorDetailsData"

    Dim strXML As String
    
    strXML = "<TABLENAME>MORTGAGECALCULATORDETAILS" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>PURCHASEPRICE<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>TERMMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TERMYEARS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEPRODUCTSTARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadMortgageCalculatorDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCMClassDef_LoadQuotationData(Optional ByVal vemcQuotationData _
                            As QuotationData = QuotationDefault) As FreeThreadedDOMDocument40
'Description : Class Definition Template for Table : SubQuote
'
'Return  : String
' *****************************************************************************
    Const strFunctionName As String = "IomCMClassDef_LoadQuotationData"

    Dim strXML As String
    Dim strXmlCommon As String
    Dim strXmlView As String
    
    Dim strXmlBCSubQuoteNumber As String
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    ' MS    04/07/00    SYS1048 BCSubQuoteNumber enumeration added to QuotationData
    strXmlBCSubQuoteNumber = "<OTHERS>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    
    strXmlCommon = "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
                    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
                    & "<PRIMARYKEY>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" _
                    'DM     18/11/2002  BMIDS00935
                    'PSC 28/06/2005 MAR5
                    'TW 28/10/2005 MAR211
    'EP1038/MAR1893 Removed redundant LIFESUBQUOTE columns
    strXmlView = "<OTHERS>ACCEPTEDAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>ACCEPTEDFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUICKQUOTEAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>QUICKQUOTEFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>RECOMMENDEDAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>RECOMMENDEDFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
                    & "<OTHERS>LOANSANDLIABILITIES<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>MONTHLYMORTGAGEPAYMENTS<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGERELATEDINSURANCE<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>OTHEROUTGOINGS<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALMONTHLYINCOME<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALQUOTATIONCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & strXmlBCSubQuoteNumber _
                    & "<OTHERS>PPSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONCOMPLETE<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>BCSUBQUOTEREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>PPSUBQUOTEREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>KFIPRINTEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>LASTOFFERED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>KFIREQUIREDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
                    

    Select Case vemcQuotationData

        Case QuotationKeyView
    
            strXML = "<TABLENAME>BASICQUOTATIONDETAILS"
            strXML = strXML & strXmlCommon & "</TABLENAME>"

        Case QuotationView
    
            strXML = strXML & "<TABLENAME>QUOTATION"
            strXML = strXML & strXmlCommon
            strXML = strXML & strXmlView
            strXML = strXML & "</TABLENAME>"

        
        Case QuotationDefault
                    
            strXML = strXML & "<TABLENAME>QUOTATION"
            strXML = strXML & strXmlCommon
            strXML = strXML & strXmlView
            strXML = strXML & "<OTHERS>MSQAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>MSQAPPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>LSQAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>LSQAPPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>BCSQAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>BCSQAPPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>PPSQAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>PPSQAPPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "</TABLENAME>"
        ' MS    04/07/00    SYS1048 BCSubQuoteNumber enumeration added to QuotationData
         Case BCSubQuoteNumber
         
            strXML = strXML & "<TABLENAME>QUOTATION"
            strXML = strXML & strXmlBCSubQuoteNumber _
            & "</TABLENAME>"
                        
    Case Else
        
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum

    End Select
    
    Set IomCMClassDef_LoadQuotationData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)


    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

End Function

Private Function IomCMClassDef_LoadLifeSubQuoteData(Optional ByVal vemcLifeSubQuoteData _
                            As LifeSubQuoteData = LifeSubQuoteDefault) As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'Description : Class Definition Template for Table : SubQuote
'
'Return  : String
' *****************************************************************************
    Const strFunctionName As String = "IomCMClassDef_LoadLifeSubQuoteData"

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Dim strXML As String
    Dim strXmlCommon As String
    
    strXML = strXML & "<TABLENAME>LIFESUBQUOTE" _
            
    strXmlCommon = "<PRIMARYKEY>LIFESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>TOTALLIFEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
            
    Select Case vemcLifeSubQuoteData
    
        Case LifeSubQuoteView
        
            strXML = strXML & strXmlCommon
            strXML = strXML & "</TABLENAME>"
        
        Case LifeSubQuoteDefault
            
            strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDate</TYPE></OTHERS>" _
            & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & strXmlCommon & "</TABLENAME>"
            
        Case Else
            
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    
    End Select
                        
    Set IomCMClassDef_LoadLifeSubQuoteData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
'            & "<OTHERS>MSQAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'            & "<OTHERS>MSQAPPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"


End Function

Private Function IomCMClassDef_LoadLifeBenefitData(Optional ByVal velbdLifeBenefitData As _
                                                   LifeBenefitData = LifeBenefitDataDefault) As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'Description : Class Definition Template for Table : LifeBenefit
'
'Return  : String
' *****************************************************************************
    Const strFunctionName As String = "IomCMClassDef_LoadLifeBenefitData"

    Dim strXML As String
    
    Select Case velbdLifeBenefitData
    
        Case QuotationSummaryView
            strXML = strXML & "<TABLENAME>LIFEBENEFIT" _
            & "<PRIMARYKEY>LIFEBENEFITSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<OTHERS>CUSTOMERNUMBER1<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CUSTOMERNUMBER2<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>BENEFITTYPE<TYPE>dbdtComboId</TYPE>" _
            & "<COMBO>LifeCoverType</COMBO></OTHERS>" _
            & "<OTHERS>MONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "</TABLENAME>"
        
        Case LifeBenefitDataDefault
            strXML = strXML & "<TABLENAME>LIFEBENEFIT" _
            & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>LIFESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>LIFEBENEFITSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<OTHERS>CUSTOMERNUMBER1<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CUSTOMERVERSIONNUMBER1<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>CUSTOMERNUMBER2<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CUSTOMERVERSIONNUMBER2<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>BENEFITTYPE<TYPE>dbdtComboId</TYPE>" _
            & "<COMBO>LifeCoverType</COMBO></OTHERS>" _
            & "<OTHERS>IPT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>MONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "</TABLENAME>"
        Case Else
            
            Dim objErrAssist As New ErrAssist
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
            Set objErrAssist = Nothing
        
    End Select
    
    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadLifeBenefitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing

End Function


Private Function IomCMClassDef_LoadBCSubQuoteData(Optional ByVal vemcBCSubQuoteData _
                            As BCSubQuoteData = BCSubQuoteDefault) As FreeThreadedDOMDocument40

' header ----------------------------------------------------------------------------------
'
' description: Used in  QuotationDO.GetQuotationDetails
'                       BuildingsAndContentsSubQuoteDO.GetData
'
' pass: 'Optional value used in view defaults to full BC details
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadBCSubQuoteData"

    Dim strXML As String
    Dim strCommon As String
    
    Dim objErrAssist As New ErrAssist
    
    
    strXML = "<TABLENAME>BUILDINGSANDCONTENTSSUBQUOTE"
    strCommon = "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & "<OTHERS>BCPRODUCTNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                & "<OTHERS>BCREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                & "<OTHERS>INSURANCEPAYFREQUENCY<TYPE>dbdtComboId</TYPE>" _
                & "<COMBO>Frequency</COMBO></OTHERS>"
     'GD BMIDS00045 added BCREFERENCENUMBER
    'KRW BMIDS776 added INSURANCEPAYFREQUENCY
    
    '*=[MC]BMIDS763 - INSURANCE ADMIN FEE FIELD ADDED TO "BUILDINGSANDCONTENTSSUBQUOTE" dB TABLE
    strCommon = strCommon & "<OTHERS>INSURANCEADMINFEE<TYPE>dbdtDouble</TYPE></OTHERS>"
    '*=SECTION END
    
    Select Case vemcBCSubQuoteData
    
        Case BCSubQuoteView
                'BG 21/08/00 SYS0964 added FREQUENCY Tag
                strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
                & "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & "<OTHERS>FREQUENCY<TYPE>dbdtInt</TYPE></OTHERS></TABLENAME>"
            
        Case BCSubQuoteDefault
            
            strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
            & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & strCommon & "</TABLENAME>"
        Case Else
            
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    
    End Select


    
    Dim objXmlAssist As New XMLAssist

    Set IomCMClassDef_LoadBCSubQuoteData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

End Function



Private Function IomCMClassDef_LoadPPSubQuoteData(Optional ByVal vemcPPSubQuoteData _
                                As PPSubQuoteData = PPSubQuoteDefault) As FreeThreadedDOMDocument40

' header ----------------------------------------------------------------------------------
' description: used in  QuotationDO.GetQuotationDetails
'
'                       Full details used here
'                       PaymentProtectionSubQuoteDO.GetSubQuoteDetails
'                       PaymentProtectionSubQuoteDO.FindList
'                       PaymentProtectionSubQuoteDO.Create
'                       PaymentProtectionSubQuoteDO.Update
'                       PaymentProtectionSubQuoteDO.GetData
' pass:
' return:
'------------------------------------------------------------------------------------------
Dim XMLstr As String

    Const strFunctionName As String = "IomCMClassDef_LoadPPSubQuoteData"

    Dim strXML As String
    Dim strCommon As String
    Dim objErrAssist As New ErrAssist
    'GD BMIDS00045 ADDED PPREFERENCENUMBER
    strXML = "<TABLENAME>PAYMENTPROTECTIONSUBQUOTE"
    strCommon = "<PRIMARYKEY>PPSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>TOTALPPMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>PPREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"

    Select Case vemcPPSubQuoteData
    
        Case PPSubQuoteView
                
             strXML = strXML & strCommon & "</TABLENAME>"
        
        Case PPSubQuoteDefault
            
            strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
            & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
            & "<OTHERS>PPPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
            strXML = strXML & strCommon & "</TABLENAME>"

        Case Else
            
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    
    End Select



    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadPPSubQuoteData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

End Function


Private Function IomCMClassDef_LoadPPDetailsData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description: Used in  PaymentProtectionSubQuoteDO.GetSubQuoteDetails
'                       PaymentProtectionSubQuoteDO.CreatePaymentProtectionDetails
'                       PaymentProtectionSubQuoteDO.UpdatePaymentProtectionDetails
' pass:
' return:
'------------------------------------------------------------------------------------------
Dim XMLstr As String

    Const strFunctionName As String = "IomCMClassDef_LoadPPDetailsData"

    Dim strXML As String

    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>PAYMENTPROTECTIONDETAILS" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>PPSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>COVERTYPE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>PPCoverType</COMBO></OTHERS>" _
    & "<OTHERS>DEFERREDPERIOD<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>IPT<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGECOVERFORAPPLICANT1<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGECOVERFORAPPLICANT2<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>PREMIUM<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>SPLITAMOUNT<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "</TABLENAME>"
    
    Set IomCMClassDef_LoadPPDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function BCSubQuoteLoadData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
Dim XMLstr As String

    Const strFunctionName As String = "BCSubQuoteLoadData"

    Dim strXML As String
    Dim xmlomCMClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "BUILDINGSANDCONTENTSSUBQUOTE"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BCPRODUCTNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INSURANCEPAYFREQUENCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Frequency</COMBO></OTHERS>"
    'KW Added INSURANCEPAYFREQUENCY BMIDS776
    

    strXML = strXML & "</TABLENAME>"

    Set xmlomCMClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set BCSubQuoteLoadData = xmlomCMClassDefDoc
    
    Set xmlomCMClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadFullQuotationData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description: Used in QuotationDO.GetQuotationDetails
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCMClassDef_LoadFullQuotationData"

    Dim strXML As String

    'this is all that is needed for the select
    'but we need more data back when we use getxmlfromrecordset
    'EP1038/MAR1893 Removed redundant LIFESUBQUOTE columns
    strXML = "<TABLENAME>GETQUOTATIONDETAILS" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>ACCEPTEDAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>ACCEPTEDFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>QUICKQUOTEAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>QUICKQUOTEFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>RECOMMENDEDAPPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>RECOMMENDEDFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "<OTHERS>LOANSANDLIABILITIES<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MONTHLYMORTGAGEPAYMENTS<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGERELATEDINSURANCE<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>OTHEROUTGOINGS<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TOTALMONTHLYINCOME<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>TOTALQUOTATIONCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>PPSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>QUOTATIONCOMPLETE<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>BCSUBQUOTEREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>PPSUBQUOTEREQUIRED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    Set IomCMClassDef_LoadFullQuotationData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing

End Function


Private Function IomCMClassDef_LoadMortgageSubQuoteData(Optional ByVal vemcMortgageSubQuoteData _
                            As MortgageSubQuoteData = MortgageSubQuoteDefault) As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description: used in  QuotationDO.GetQuotationDetails
'
' pass:
' return:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "IomCMClassDef_LoadMortgageSubQuoteData"

    Dim strXML As String
    Dim strXmlCommon As String
    Dim strKey As String
    Dim strOption As String
    Dim strAmountRequested As String
        
        strAmountRequested = "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtLong</TYPE></OTHERS>"
        
        'DPF 08/10/2002 - CPWP1 (BM037) added MONTHLYCOSTLESSDRAWDOWN
        'AM 14/09/2005 - MARS29 added CONFIRMEDMAXBORROWING,CONFIRMEDCALCULATEDINCMULTIPLE,
        '                               CONFIRMEDCALCULATEDINCMULTIPLIERTYPE
        'JD MAR30 CONFIRMEDCALCULATEDINCMULTIPLIERTYPE type changed to dbdtString
        strXmlCommon = "<OTHERS>DEPOSIT<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>LTV<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>NONPANELLENDERSELECTD<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALCHARGES<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALGROSSMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALINITIALCHARGES<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALNETCHARGES<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALNETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>UNUSEDMIRASAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>NONPANELLENDERNAME<TYPE>dbdtString</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALACCRUEDINTEREST<TYPE>dbdtLong</TYPE></OTHERS>" _
                        & "<OTHERS>MANUALINCENTIVEAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>MANUALINCENTIVEUSERID<TYPE>dbdtString</TYPE></OTHERS>" _
                        & "<OTHERS>MANUALINCENTIVEDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" _
                        & "<OTHERS>DRAWDOWN<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>MONTHLYCOSTLESSDRAWDOWN<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>CONFIRMEDMAXBORROWING<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>CONFIRMEDCALCULATEDINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>CONFIRMEDCALCULATEDINCMULTIPLIERTYPE<TYPE>dbdtString</TYPE></OTHERS>"
                    
        'BMIDS736 GHun
        strXmlCommon = strXmlCommon _
                        & "<OTHERS>AMOUNTPERUNITBORROWED<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>"
        'BMIDS736 End
        
        'BMIDS766 GHun
        strXmlCommon = strXmlCommon _
                        & "<OTHERS>TOTALAMOUNTPAYABLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
                        & "<OTHERS>TOTALMORTGAGEPAYMENTS<TYPE>dbdtDouble</TYPE></OTHERS>"
        'BMIDS766 End
        
        'MAR1061
        strXmlCommon = strXmlCommon _
                        & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>" _
    
        strKey = "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
                & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

        
    Select Case vemcMortgageSubQuoteData
    
        Case MortgageSubQuoteView
    
                'DM     18/11/2002  BMIDS00935
                strXML = "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" _
                & "<OTHERS>TOTALNETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & "<OTHERS>TOTALGROSSMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"

        Case MortgageSubQuoteDefault
    
                'BMIDS00654 MDC 31/10/2002 - Added fields INCOMEMULTIPLIERTYPE, INCOMEMULTIPLE & MAXIMUMBORROWINGAMOUNT
'                strXML = strKey _
                & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
                & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
                & strAmountRequested & strXmlCommon
                'DM     18/11/2002  BMIDS00935
                strXML = strKey _
                & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" _
                & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>" _
                & "<OTHERS>INCOMEMULTIPLIERTYPE<TYPE>dbdtString</TYPE></OTHERS>" _
                & "<OTHERS>INCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & "<OTHERS>MAXIMUMBORROWINGAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                & strAmountRequested & strXmlCommon
                'BMIDS00654 MDC 31/10/2002 - End
        
        Case MortgageSubQuoteAmountRequested
                'DM     18/11/2002  BMIDS00935
                strXML = strKey _
                & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" _
                & strAmountRequested

'        Case MortgageKeyView
'                strXml = strKey _
'                & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _

        Case Else
                   
            Dim objErrAssist As ErrAssist
            Set objErrAssist = New ErrAssist
        
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum

    End Select

        strXML = "<TABLENAME>MORTGAGESUBQUOTE" & strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadMortgageSubQuoteData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

End Function

Private Function IomCMClassDef_LoadLoanComponentData(Optional ByVal vemcLoanComponentData _
                            As LoanComponentData = LoanComponentDefault) As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description: Used in MortgageSubQuoteDO.FindLoanComponentList
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadLoanComponentData"

    Dim strXML As String
    Dim strLoanComponentView As String
    Dim strLoanComponentFindList As String
    Dim strCommon As String
    Dim strLoanAmount As String
    Dim strPrimaryKey As String
    Dim strPortingData As String 'SR 16/08/2004:BMIDS815
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist

    strLoanAmount = "<OTHERS>LOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"

    strPrimaryKey = "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    'DM     18/11/2002  BMIDS00935
    
    ' DM 25/03/02 SYS4313
    strLoanComponentFindList = strPrimaryKey _
    & strLoanAmount _
    & "<OTHERS>PORTEDLOAN<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>TOTALLOANCOMPONENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _
    & "<OTHERS>ADMINSUBLEDGERCODE<TYPE>dbdtString</TYPE></OTHERS>"
    
    'SG 28/05/02 SYS4767
    strLoanComponentFindList = strLoanComponentFindList & _
    "<OTHERS>ADMINSUBLEDGERCODECI<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>ADMINSUBLEDGERCODEIO<TYPE>dbdtString</TYPE></OTHERS>"
    
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>REPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>RepaymentType</COMBO></OTHERS>"
    
'SW 20/06/2006 EP771 Start
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>REPAYMENTVEHICLE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>RepaymentVehicle</COMBO></OTHERS>"
'SW 20/06/2006 EP771 End
'MAH   03/01/2007  EP2_444 start
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>REPAYMENTVEHICLEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
'MAH   03/01/2007  EP2_444 End
    'SG 28/05/02 SYS4767
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>PRODUCTCODESEARCHIND<TYPE>dbdtInt</TYPE></OTHERS>"
    
    ' PSC 03/07/2002 BMIDS00062 - Start
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>FINALRATE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>FINALRATEAPR<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>FINALRATEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
    ' PSC 03/07/2002 BMIDS00062 - End
    
    'SR 10/08/2004 : BMIDS815
    strLoanComponentFindList = strLoanComponentFindList _
    & "<OTHERS>PRODUCTSTEP<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>REMAININGSTEPDURATION<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>PORTEDPRODUCTSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    'SR 10/08/2004 : BMIDS815 - End
    
    strLoanComponentView = strPrimaryKey _
    & strLoanAmount _
    & "<OTHERS>PORTEDLOAN<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _
    & "<OTHERS>TOTALLOANCOMPONENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>REPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>RepaymentType</COMBO></OTHERS>"
    
'SW 20/06/2006 EP771 Start
'    strLoanComponentView = strLoanComponentView _
'    & "<OTHERS>REPAYMENTVEHICLE<TYPE>dbdtComboId</TYPE>" _
'    & "<COMBO>RepaymentVehicle</COMBO></OTHERS>"
'SW 20/06/2006 EP771 End
    
    
    'SG 28/05/02 SYS4767
    'CMWP3 - DPF 11/7/02 - Have added 2 new fields (SUBPURPOSEOFLOAN & MANUALPORTEDLOANIND)
    'CPWP1 - DPF 23/10/02 - have added MonthlyCostLessDrawdown
    strLoanComponentView = strLoanComponentView _
    & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>STARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPITALANDINTERESTELEMENT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>GROSSMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTONLYELEMENT<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>MIRASAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MIRASELIGIBILITY<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>NETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>SUBPURPOSEOFLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>MANUALPORTEDLOANIND<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>TERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TERMINYEARS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>RESOLVEDRATE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MONTHLYCOSTLESSDRAWDOWN<TYPE>dbdtDouble</TYPE></OTHERS>"

    strCommon = "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>STARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPITALANDINTERESTELEMENT<TYPE>dbdtLong</TYPE></OTHERS>"
    
    'SG 28/05/02 SYS4767
    strCommon = strCommon _
    & "<OTHERS>NETCAPANDINTELEMENT<TYPE>dbdtLong</TYPE></OTHERS>" _

    strCommon = strCommon _
    & "<OTHERS>GROSSMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTONLYELEMENT<TYPE>dbdtInt</TYPE></OTHERS>"

    'SG 28/05/02 SYS4767
    strCommon = strCommon _
    & "<OTHERS>NETINTONLYELEMENT<TYPE>dbdtInt</TYPE></OTHERS>"
    
    'DPF 08/10/2002 CPWP1 - BM037
    strCommon = strCommon _
    & "<OTHERS>MONTHLYCOSTLESSDRAWDOWN<TYPE>dbdtDouble</TYPE></OTHERS>"

    'DPF 22/07/02 - added two new fields: SUBPURPOSEOFLOAN & MANUALPORTEDLOANIND
    strCommon = strCommon _
    & "<OTHERS>MIRASAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>MIRASELIGIBILITY<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>NETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>SUBPURPOSEOFLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>MANUALPORTEDLOANIND<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>TERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TERMINYEARS<TYPE>dbdtInt</TYPE></OTHERS>"

    'SG 28/05/02 SYS4767
    strCommon = strCommon _
    & "<OTHERS>RESOLVEDRATE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MANUALADJUSTMENTPERCENT<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MANUALADJUSTUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    
    'BMIDS736 GHun
    strCommon = strCommon _
    & "<OTHERS>AMOUNTPERUNITBORROWED<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPINTMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTONLYMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MINMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MAXMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>UNROUNDEDAPR<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>ACCRUEDINTEREST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPINTACCRUEDINTEREST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTONLYACCRUEDINTEREST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INCREASEDMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INCREASEDMONTHLYCOSTDIFFERENCE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>TOTALAMOUNTPAYABLE<TYPE>dbdtDouble</TYPE></OTHERS>"
    'BMIDS736 End
    
    strCommon = strCommon & "<OTHERS>ORIGINALLTV<TYPE>dbdtInt</TYPE></OTHERS>"  'EP2_56 GHun
    
    'SR 16/08/2004:BMIDS815
    strPortingData = strPrimaryKey & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>PRODUCTSTEP<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>REMAININGSTEPDURATION<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>MANUALPORTEDLOANIND<TYPE>dbdtInt</TYPE></OTHERS>"
    'SR 16/08/2004:BMIDS815 - End
    
    'MAR46 GHun
    ' PSC 12/03/2007 EP2_1611
    strPortingData = strPortingData & "<OTHERS>PORTEDPRODUCTSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
                    & "<OTHERS>PRODUCTSWITCHRETAINPRODUCTIND<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>CURRENTRATEEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _

                    
    strLoanComponentFindList = strLoanComponentFindList & "<OTHERS>PRODUCTSWITCHRETAINPRODUCTIND<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>CURRENTRATEEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    'MAR46 End

    'EP2_56 add MORTGAGELOANGUID
    'EP2_696 Removed duplicate MortgageLoanGUID
    'strLoanComponentFindList = strLoanComponentFindList & "<OTHERS>MORTGAGELOANGUID<TYPE>dbdtGuid</TYPE></OTHERS>"


    Select Case vemcLoanComponentData
    
        Case LoanComponentDefault
        
            strXML = strLoanComponentFindList & strCommon
        
        Case LoanComponentView
        
            strXML = strLoanComponentView & strCommon
        
'        Case LoanComponentFindList
            
'            strXML = strLoanComponentFindList
            
         Case LoanComponentLoanAmount
        
             strXML = strLoanAmount
            
'        Case LoanComponentPrimaryKey
        
'            strXML = strPrimaryKey
            
        Case LoanComponentStoredQuoteView
        
            'PSC 28/06/2005 MAR5
            strXML = "<OTHERS>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>REPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>" _
                     & "<COMBO>RepaymentType</COMBO></OTHERS>" _
                     & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>FINALRATEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>LOANAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>CAPITALANDINTERESTELEMENT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>INTERESTONLYELEMENT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
                     
            'SG 28/05/02 SYS4767
            strXML = strXML & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>"
        
            'DPF 30/07/02 BMIDS00224
            strXML = strXML & "<OTHERS>FINALRATEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
        
            'SW 20/06/2006 EP771 Start
'SW            strXML = strXML & "<OTHERS>REPAYMENTVEHICLE<TYPE>dbdtComboId</TYPE>" _
                            & "<COMBO>RepaymentVehicle</COMBO></OTHERS>"
            'SW 20/06/2006 EP771 End
                
        Case QuotationSummaryView
        'DPF 15/11/2002 - BMIDS00516 - add two fields to view (c&i & i/o amounts)
        
            strXML = "<OTHERS>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>PORTEDLOAN<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>REPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>" _
                     & "<COMBO>RepaymentType</COMBO></OTHERS>" _
                     & "<OTHERS>TOTALLOANCOMPONENTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                     & "<OTHERS>TERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>TERMINYEARS<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>WITHDRAWNDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
                     & "<OTHERS>MORTGAGEPRODUCTDESCRIPTION<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>INTERESTRATE<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>MONTHLYREPAYMENT<TYPE>dbdtInt</TYPE></OTHERS>" _
                     & "<OTHERS>NETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>CAPITALANDINTERESTELEMENT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>INTERESTONLYELEMENT<TYPE>dbdtDouble</TYPE></OTHERS>"

            'SG 28/05/02 SYS4767
            strXML = strXML & "<OTHERS>APR<TYPE>dbdtDouble</TYPE></OTHERS>" _
                     & "<OTHERS>RESOLVEDRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
            
            'DPF 30/07/02 BMIDS00224
            strXML = strXML & "<OTHERS>FINALRATEMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
            
            'INR 08/08/03 BMIDS624
            strXML = strXML & "<OTHERS>MANUALADJUSTMENTPERCENT<TYPE>dbdtDouble</TYPE></OTHERS>"
            
            'JD BMIDS765 CC076
            strXML = strXML & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>" _
                            & "<OTHERS>STARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
        
'            'SW 20/06/2006 EP771 Start
'            strXML = strXML & "<OTHERS>REPAYMENTVEHICLE<TYPE>dbdtComboId</TYPE>" _
'                            & "<COMBO>RepaymentVehicle</COMBO></OTHERS>"
'            'SW 20/06/2006 EP771 End
        
       Case LoanComponentPortingData
            strXML = strPortingData
                     
       Case Else
        
            Set objErrAssist = New ErrAssist
                
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum

    End Select
    
    ' SR 09/01/2003 : BM0209 - Do not lock table when queried (as this is application specific data)
    'strXML = "<TABLENAME>LOANCOMPONENT" & strXML & "</TABLENAME>"
    strXML = "<TABLENAME>LOANCOMPONENT<SQLNOLOCK>TRUE</SQLNOLOCK> " & strXML & "</TABLENAME>"
    ' SR 09/01/2003 : BM0209 - End
    
    Set objXmlAssist = New XMLAssist
    
    Set IomCMClassDef_LoadLoanComponentData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadLoanComponentRedemptionFeeData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCMClassDef_LoadLoanComponentRedemptionFeeData"

    Dim strXML As String
    
    strXML = "<TABLENAME>LOANCOMPONENTREDEMPTIONFEE" _
    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>REDEMPTIONFEESTEPNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>REDEMPTIONFEEPERIOD<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>REDEMPTIONFEEPERIODENDDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>REDEMPTIONFEEAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadLoanComponentRedemptionFeeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCMClassDef_LoadMortgageIncentiveData(Optional ByVal vemcMortgageIncentiveData _
                            As MortgageIncentiveData = MortgageIncentiveDefault) As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadMortgageIncentiveData"

    Dim strXML As String
    Dim strCommon As String
    Dim strTable As String
    
        strTable = "<TABLENAME>MORTGAGEINCENTIVE"
        strCommon = strTable & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
        & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
        & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    Select Case vemcMortgageIncentiveData
    
    Case MortgageIncentiveTotalReturn
            
            'TOTALINCENTIVEAMOUNT is NOT a field on this table
            'it is defined as the summed return for the sql
            strXML = strTable & "<OTHERS>TOTALINCENTIVEAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
            
    Case MortgageIncentiveTotal
            
            strXML = strCommon
    
    Case MortgageIncentiveDefault
    
        'DPF 23/10/2002 - CPWP1 - added INCENTIVEGUID field to Class Definition
    
        strXML = strCommon & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
        & "<PRIMARYKEY>MORTGAGEINCENTIVESEQUENCENUM<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
        & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>" _
        & "<OTHERS>INCENTIVEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
        & "<OTHERS>TYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
        & "<OTHERS>INCENTIVEBENEFITTYPE<TYPE>dbdtComboId</TYPE>" _
        & "<COMBO>IncentiveBenefitType</COMBO></OTHERS>" _
        & "<OTHERS>COMPLETIONINDICATOR<TYPE>dbdtComboId</TYPE>" _
        & "<COMBO>CompletionIndicator</COMBO></OTHERS>" _
        & "<OTHERS>INCENTIVEGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        
        
    Case Else
                
        Dim objErrAssist As New ErrAssist
            
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum

    
    End Select
    
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadMortgageIncentiveData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
End Function

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be totally removed due to OOSS
'Private Function IomCMClassDef_LoadLoanComponentResultsData() As FreeThreadedDOMDocument40
'' header ----------------------------------------------------------------------------------
''
'' description:
'' pass:
'' return:
''------------------------------------------------------------------------------------------
'    Dim strFunctionName As String
'    strFunctionName = "IomCMClassDef_LoadLoanComponentResultsData"
'
'    Dim strXML As String
'
'    strXML = "<TABLENAME>LOANCOMPONENTRESULTS" _
'    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
'    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
'    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
'    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
'    & "<OTHERS>ACCRUEDINTEREST<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>FINALPAYMENT<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT2<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT3<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT4<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT5<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT6<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>GROSSPAYMENT7<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT2<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT3<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT4<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT5<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT6<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>NETPAYMENT7<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>TOTALAMOUNTPAYABLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>TOTALMORTGAGEPAYMENTS<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "<OTHERS>UNROUNDEDAPR<TYPE>dbdtDouble</TYPE></OTHERS>" _
'    & "</TABLENAME>"
'
'    Dim objXmlAssist As New XMLAssist
'
'    Set IomCMClassDef_LoadLoanComponentResultsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
'
'    Set objXmlAssist = Nothing
'End Function
'BMIDS736 End
'BMIDS782 End

Private Function IomCMClassDef_LoadMortgageOneOffCostData() As FreeThreadedDOMDocument40

' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadMortgageOneOffCostData"

    Dim strXML As String
    'GD  BM0256 17/02/2003 Added SQLNOLOCK
    'DRC  BMIDS767 18/05/2004 Added AdHocInd
    'HMA  MAR28  Add Refund Amount
    'SR   EP2_1753 Add Rebate Amount
    strXML = "<TABLENAME>MORTGAGEONEOFFCOST" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGEONEOFFCOSTTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>ONEOFFCOSTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>" _
    & "<OTHERS>ADDTOLOAN<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>ADHOCIND<TYPE>dbdtBoolean</TYPE></OTHERS>" _
    & "<OTHERS>AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>IPT<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>REFUNDAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "<OTHERS>REBATEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadMortgageOneOffCostData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadNonPanelMortgageProductData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadNonPanelMortgageProductData"

    Dim strXML As String
    
    strXML = "<TABLENAME>NONPANELMORTGAGEPRODUCT" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>INTERESTPERIOD1<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTRATE1<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTRATE2<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEPRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>MORTGAGEPRODUCTTYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadNonPanelMortgageProductData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadPortedLenderData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomCMClassDef_LoadPortedLenderData"

    Dim strXML As String
    
    strXML = "<TABLENAME>MORTGAGELENDERDIRECTORY" _
    & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomCMClassDef_LoadPortedLenderData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadBCDetailsData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCMClassDef_LoadBCDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>BUILDINGSANDCONTENTSDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCIDENTALBUILDCOVERREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCIDENTALCONTENTCOVERREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALARMFITTED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ANNUALBUILDINGSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ANNUALCONTENTSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BELONGINGSFROMHOMEREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSEXCESSOPTION<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>BuildingExcessAmount</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSEXCESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSEXCESSREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGSSUMINSURED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTENTSEXCESSOPTION<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>ContentsExcessAmount</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>CONTENTSEXCESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTENTSEXCESSREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTENTSSUMINSURED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COVERTYPE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>BCCoverType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>FREQUENCY<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>BCPaymentFreq</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>IPT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYBUILDINGSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYCONTENTSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFCLAIMSLAST5YEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFBEDROOMS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ANNUALBCPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROPERTYAFFECTED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROPERTYPOSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STATUS<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>BCSubQuoteStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALVALUABLESEXCEEDINGLIMIT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALVALUABLESEXCEEDEDIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BCRETURNTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"


    Set IomCMClassDef_LoadBCDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadBCNotesData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCMClassDef_LoadBCNotesData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>BUILDINGSANDCONTENTSNOTES"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>NOTES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCMClassDef_LoadBCNotesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadBCValuablesOLimitData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCMClassDef_LoadBCValuablesOLimitData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>VALUABLESOVERLIMIT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>VALUABLESLIMITSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INSURABLEVALUE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCMClassDef_LoadBCValuablesOLimitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadQuotationLoanSummaryViewData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCMClassDef_LoadQuotationLoanSummaryViewData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    'JD MAR1040 added incmultiples
    'JD MAR1061 added purchasepriceorestimatedvalue
    strXML = strXML & "<TABLENAME>QUOTATIONLOANSUMMARY" _
                    & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>LOANSANDLIABILITIES<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>MONTHLYMORTGAGEPAYMENTS<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGERELATEDINSURANCE<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>OTHEROUTGOINGS<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALMONTHLYINCOME<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>LTV<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>INCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>CONFIRMEDCALCULATEDINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>PURPOSEOFLOAN<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>PurposeOfLoan</COMBO></OTHERS>" _
                    & "<OTHERS>TERMINYEARS<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>TERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>ONEOFFCOSTTYPE<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>SUBPURPOSEOFLOAN<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>SubPurposeOfLoan</COMBO></OTHERS>" _
                    & "</TABLENAME>"

    Set IomCMClassDef_LoadQuotationLoanSummaryViewData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadStoredQuoteFindListViewData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomCMClassDef_LoadStoredQuoteFindListViewData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    'PSC 28/06/2005 MAR5
    'EP1038/MAR1893 Removed redundant LIFECOVER
    strXML = strXML & "<TABLENAME>STOREDQUOTATION" _
                    & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALQUOTATIONCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALNETMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALPPMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>BCCOVERTYPE<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>BCCoverType</COMBO></OTHERS>" _
                    & "<OTHERS>ANNUALBUILDINGSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>ANNUALCONTENTSPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>PPCOVERTYPE<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>PPCoverType</COMBO></OTHERS>" _
                    & "<OTHERS>LASTOFFERED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>KFIREQUIREDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>LTV<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "</TABLENAME>"

    Set IomCMClassDef_LoadStoredQuoteFindListViewData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCMClassDef_LoadQuotationSummaryViewData() As FreeThreadedDOMDocument40

    
    Const strFunctionName As String = "IomCMClassDef_LoadQuotationSummaryViewData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
'GD     15/08/02    BMIDS00320
    strXML = strXML & "<TABLENAME>QUOTATIONSUMMARY" _
                    & "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>QUOTATIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>PURCHASEPRICEORESTIMATEDVALUE<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>LTV<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALPPMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>PPCOVERTYPE<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>PPCoverType</COMBO></OTHERS>" _
                    & "<OTHERS>PPREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGECOVERFORAPPLICANT1<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>MORTGAGECOVERFORAPPLICANT2<TYPE>dbdtLong</TYPE></OTHERS>" _
                    & "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
                    & "<OTHERS>BCCOVERTYPE<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>BCCoverType</COMBO></OTHERS>" _
                    & "<OTHERS>BCREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
                    & "<OTHERS>BCFREQUENCY<TYPE>dbdtComboId</TYPE>" _
                    & "<COMBO>BCPaymentFreq</COMBO></OTHERS>" _
                    & "<OTHERS>ACCIDENTALBUILDCOVERREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "<OTHERS>ACCIDENTALCONTENTCOVERREQUIRED<TYPE>dbdtInt</TYPE></OTHERS>" _
                    & "</TABLENAME>"

    Set IomCMClassDef_LoadQuotationSummaryViewData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

'BMIDS736 GHun
Private Function IomCMClassDef_LoadLoanComponentBalanceScheduleData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IomCMClassDef_LoadLoanComponentBalanceScheduleData"
    Dim strXML            As String
    Dim objXmlAssist      As XMLAssist
    
    'LOANCOMPONENTBALANCESCHEDULEID is part of the primary key but is a generated value (identity column)
    
    strXML = "<TABLENAME>LOANCOMPONENTBALANCESCHEDULE" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>LOANCOMPONENTBALANCESCHEDULEID<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>SCHEDULETYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>STARTDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>BALANCE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTONLYBALANCE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPINTBALANCE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "</TABLENAME>"
    
    Set objXmlAssist = New XMLAssist
    Set IomCMClassDef_LoadLoanComponentBalanceScheduleData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
'BMIDS736 End

'BMIDS736 GHun
Private Function IomCMClassDef_LoadLoanComponentPaymentScheduleData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IomCMClassDef_LoadLoanComponentPaymentScheduleData"
    Dim strXML            As String
    Dim objXmlAssist      As XMLAssist
    
    ' LOANCOMPONENTPAYMENTSCHEDULEID is part of the primary key but is a generated value (identity column)
    
    strXML = "<TABLENAME>LOANCOMPONENTPAYMENTSCHEDULE" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>MORTGAGESUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>LOANCOMPONENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>LOANCOMPONENTPAYMENTSCHEDULEID<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>INTERESTRATETYPESEQUENCENUMBER<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>STARTDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>MONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INTONLYMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>CAPINTMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MINMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>MAXMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INCREASEDMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "<OTHERS>INCREASEDMONTHLYCOSTDIFFERENCE<TYPE>dbdtDouble</TYPE></OTHERS>" _
    & "</TABLENAME>"
    
    Set objXmlAssist = New XMLAssist
    Set IomCMClassDef_LoadLoanComponentPaymentScheduleData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
'BMIDS736 End


