VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omMPClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omMPClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   XML definitions for Class/Tables for omMP
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' PSC   14/12/99    Created
' MCS   12/01/2000  amended
' JLD   01/03/2000  Added format of 2dp to FIRSTMONTHLYINTERESTRATE
' AY    21/03/00    QuotationSummary structures altered
' JLD   23/05/00    Added TypicalAPR to output of LoadMortgageProductSelectionData
' JLD   30/05/00    Added FIRSTMONTHSREPAYMENTBYDAILY, FIRSTMONTHLYSVRREPAYBYDAILY to output of LoadMortgageProductSelectionData
' JLD   27/09/00    SYS1573 - changed datatype of STARTDATE on mortgageproduct table to dateTime from date
' SR    18/06/01    modified 'IomMPClassDef_LoadIncentiveData' - two new columns
'                   'IncentiveBenefitType' and 'NotionalAmount'
' MDC   06/09/2001  SYS2250 Add BaseRate table
' MC    06/09/2001  SYS2250 - Add LoadBaseRateData and amend LoadBaseRateBandData
' GF    15/10/2001  SYS2250 - Changed DomDocument reference to remove XML version
' JLD   16/10/2001  SYS2250 - temporarily removed optional parameter to LoadBaseRateBandData
' DB    28/05/2002  SYS4767 - MSMS to Core Integration
'-------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
'
' GD    06/06/02    BMIDS00034 Amended LoadMortgageProductSelectionData
' MDC   14/06/2002  BMIDS00074 Amended LoadInterestRateTypeData
' MV    21/06/2002  BMIDS00077 Core Upgrade - Amended IomMPClassDef_LoadMortgageProductData
' DPF   15/07/2002  CMWP3   -  Have added a new method called LoadRedemptionFeeBandData()
' MV    03/09/2002  BMIDS00353 Amended IomMPClassDef_LoadMortgageProductData() with MPMIGSTARTLTV and MPMIGRATESET
' GD    10/09/02    BMIDS00313 - APWP2 - BM004 Added LoadConditionsData
' PSC   18/09/2002  SYS5077    Applied core change to LoadMortgageProductSelectionData
' MDC   01/11/2002  BMIDS00654 Added LoadMortgageProductIncomeMultiplierSet
' KRW   29/04/2004  BMIDS737 Removed RedemptionFeeStartDate
' JD    21/06/04    BMIDS765 CC076 added LoadRentalIncomeRateData
' HMA   13/12/04    BMIDS959   Remove MortgageProductBands
'-------------------------------------------------------------------------------------------
'Mars History:
'
' PE    10/02/2006  MAR1190 - Added fields LoadBaseRateBandData SQL
' GHun  26/06/2006  MAR1878 Changed LoadMortgageProductSelectionData to remove redundant columns
'-------------------------------------------------------------------------------------------
' Epsom2 History:
'
' IK    02/03/2007  EP2_1283 LoadMortgageProductSelectionData to return FREELEGALFEES
'-------------------------------------------------------------------------------------------

Option Explicit

Implements IomMPClassDef

Private Function IomMPClassDef_LoadBaseRateData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomMPClassDef_LoadBaseRateData"
    
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "BASERATE"
    strXML = strXML & "<PRIMARYKEY>RATEID<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BASERATESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BASEINTERESTRATE<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>RATETYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BASERATEAPPLIEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IomMPClassDef_LoadBaseRateData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadLenderNamesData() As FreeThreadedDOMDocument40
     
    Const strFunctionName As String = "IomMPClassDef_LoadLenderNamesData"
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGELENDER"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDCOSTSTOLOAN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDIPTTOMIGPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWADMINFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWARRGEMTFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWDEEDSRELFEEADD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWMIGFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWPORTINGFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWREINSPTFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWSEALINGFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWTTFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWVALNFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MLARRANGEMENTFEEAMOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MLARRANGEMENTFEEPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MLARRANGEMENTFEEPERCENTMAX<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DEEDSRELEASEFEE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IGNOREMIGPREMIUMAMOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERENDDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERSTARTDATE<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXNOLOANS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIGSTARTLTV<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MIGTHRESHOLD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMMIGPREMIUM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SEALINGFEE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TTFEE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SINGLEINCOMEMULTIPLE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HIGHESTEARNINCMULTIPLE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDEARNINCMULTIPLE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>JOINTINCOMEMULTIPLE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWLEGALFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ALLOWOTHERFEEADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RATESET<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NONGUARANTEEDINCOMEMULT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GUARANTEEDINCOMEMULT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomMPClassDef_LoadProductNamesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function IomMPClassDef_LoadAdministrationFeeData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomMPClassDef_LoadAdministrationFeeData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "ADMINISTRATIONFEE"
    strXML = strXML & "<PRIMARYKEY>ADMINISTRATIONFEESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>ADMINISTRATIONFEESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TypeOfMortgage</COMBO></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>LOCATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyLocation</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFAPPLICATIONTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOCATIONTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadAdministrationFeeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IomMPClassDef_LoadRedemptionFeeBandData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomMPClassDef_LoadRedemptionFeeBandData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "REDEMPTIONFEEBAND"
    strXML = strXML & "<PRIMARYKEY>REDEMPTIONFEESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
   ' strXML = strXML & "<PRIMARYKEY>REDEMPTIONFEESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    'BMIDS737 Removed above line KRW 29/04/2004
    strXML = strXML & "<PRIMARYKEY>REDEMPTIONFEESTEPNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>PERIOD<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PERIODENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FEEPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FEEMONTHSINTEREST<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LASTCHANGEDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadRedemptionFeeBandData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IomMPClassDef_LoadProductNamesData() As FreeThreadedDOMDocument40
     
    Const strFunctionName As String = "IomMPClassDef_LoadProductNamesData"
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCT"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME> "

    Set IomMPClassDef_LoadProductNamesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function IomMPClassDef_LoadBaseRateBandData(Optional ByVal vRateOption As BaseRateOption = opGetRateDifference) As FreeThreadedDOMDocument40
    
    'The RateOption parameter allows this classdef to be used in two situations:
    '   1) Retrieving data from BaseRateBand as is getting the RateDifference (default option)
    '   2) Retrieving an absolute rate calculated as BaseRate.BaseInterestRate + BaseRateBand.RateDifference
    
    Const strFunctionName As String = "IomMPClassDef_LoadBaseRateBandData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "BASERATEBAND"
    strXML = strXML & "<PRIMARYKEY>BASERATESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BASERATEBANDSTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMLOANAMOUNT<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMLTV<TYPE>dbdtDouble</TYPE><FORMATMASK>0.000</FORMATMASK></PRIMARYKEY>"
    'SYS2250 MDC 06/09/2001. Add RateDifference field
    If vRateOption = opGetRateDifference Then
        strXML = strXML & "<OTHERS>RATEDIFFERENCE<TYPE>dbdtDouble</TYPE><FORMATMASK>0.000</FORMATMASK></OTHERS>"
    Else
        strXML = strXML & "<OTHERS>RATE<TYPE>dbdtDouble</TYPE><FORMATMASK>0.000</FORMATMASK></OTHERS>"
    End If
    'SYS2250 End
    strXML = strXML & "<OTHERS>RATEDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BASERATESTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BASERATESETDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadBaseRateBandData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadIncentiveData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'INCENTIVEGUID                 NOT NULL RAW(16)
'DESCRIPTION                            VARCHAR2(50)
'AMOUNT                                 Number
'PERCENTAGE                             Number
'PERCENTAGEMAXIMUM                      Number
' *****************************************************************************
    Const strFunctionName As String = "IomMPClassDef_LoadIncentiveData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INCENTIVE"
    strXML = strXML & "<PRIMARYKEY>INCENTIVEGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PERCENTAGE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PERCENTAGEMAXIMUM<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INCENTIVEBENEFITTYPE<TYPE>dbdtComboId</TYPE><COMBO>IncentiveBenefitType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>NOTIONALAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadIncentiveData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing



End Function

Private Function IomMPClassDef_LoadInterestRateTypeData() As FreeThreadedDOMDocument40

'------------------------------------------------------------------------------------------
'MORTGAGEPRODUCTCODE            NOT NULL    VARCHAR2(6)
'STARTDATE                      NOT NULL    DATE
'INTERESTRATETYPESEQUENCENUMBER NOT NULL    NUMBER(5)
'INTERESTRATEENDDATE                        Date
'INTERESTRATEPERIOD                         Number(5)
'RATETYPE                                   VARCHAR2(1)
'RATE                                       NUMBER
'CEILINGRATE                                Number
'FLOOREDRATE                                Number
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomMPClassDef_LoadInterestRateTypeData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INTERESTRATETYPE"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>INTERESTRATETYPESEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INTERESTRATEENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERESTRATEPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RATETYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CEILINGRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLOOREDRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    'BMIDS00074 MDC 14/06/2002 - Add BaserateSet
    strXML = strXML & "<OTHERS>BASERATESET<TYPE>dbdtInt</TYPE></OTHERS>"
    'BMIDS00074 MDC 14/06/2002 - End
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadInterestRateTypeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadMortgageProductApplicationTypeData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'MORTGAGEPRODUCTCODE            NOT NULL VARCHAR2(6)
'STARTDATE                      NOT NULL DATE
'APPLICATIONTYPE                         NUMBER(5)
' *****************************************************************************
    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductApplicationTypeData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCTAPPLICATIONTYPE"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>APPLICATIONTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TypeOfMortgage</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadMortgageProductApplicationTypeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

' BMIDS959  Remove IomMPClassDef_LoadMortgageProductBandsData

Private Function IomMPClassDef_LoadMortgageProductParametersData(Optional ByVal vempMortProdParametersData As _
                            MortProdParametersData = MortProdParametersDefault) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductParametersData"

    Dim strXML As String
    Dim strValidateCompulsoryProducts As String
    
    Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>MORTGAGEPRODUCTPARAMETERS"
    strValidateCompulsoryProducts = "<OTHERS>COMPULSORYBC<TYPE>dbdtBoolean</TYPE></OTHERS>" _
                    & "<OTHERS>COMPULSORYPP<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    
    Select Case vempMortProdParametersData
    
    Case ValidateCompulsoryProducts
    
            strXML = strXML & strValidateCompulsoryProducts
        
    Case MortProdParametersDefault

    
        strXML = strXML _
                & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
                & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>" _
                & "<OTHERS>COMPULSORYBCTERM<TYPE>dbdtInt</TYPE></OTHERS>" _
                & "<OTHERS>COMPULSORYPPTERM<TYPE>dbdtInt</TYPE></OTHERS>" _
                & "<OTHERS>DIFFERENTIALUPPERLIMIT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & "<OTHERS>DIFFERENTIALLOWERLIMIT<TYPE>dbdtDouble</TYPE></OTHERS>" _
                & strValidateCompulsoryProducts
    Case Else
            
        Dim objErrAssist As New ErrAssist
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
        Set objErrAssist = Nothing
        
    End Select

    strXML = strXML & "</TABLENAME>"
    
    Set IomMPClassDef_LoadMortgageProductParametersData = _
                                objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadMortgageProductPurposeOfLoanData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'PURPOSEOFLOANSEQUENCENUMBER    NOT NULL NUMBER(5)
'MORTGAGEPRODUCTCODE            NOT NULL VARCHAR2(6)
'STARTDATE                      NOT NULL DATE
' *****************************************************************************
    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductPurposeOfLoanData"

    Dim strXML As String
    
    strXML = strXML + "<TABLENAME>"
    strXML = strXML + "MORTGAGEPRODUCTPURPOSEOFLOAN"
    strXML = strXML + "<PRIMARYKEY>PURPOSEOFLOANSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML + "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML + "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML + "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadMortgageProductPurposeOfLoanData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing



End Function

Private Function IomMPClassDef_LoadOtherFeeData() As FreeThreadedDOMDocument40

'------------------------------------------------------------------------------------------
'MORTGAGEPRODUCTCODE            NOT NULL VARCHAR2(6)
'STARTDATE                      NOT NULL DATE
'NAME                           NOT NULL VARCHAR2(20)
'OTHERFEESTARTDATE              NOT NULL DATE
'OTHERFEEAMOUNT                          NUMBER
' *****************************************************************************
    Const strFunctionName As String = "IomMPClassDef_LoadOtherFeeData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "OTHERFEE"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>NAME<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>OTHERFEESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHER>OTHERFEEAMOUNT<TYPE>dbdtDouble</TYPE></OTHER>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadOtherFeeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadQuotationSummaryViewData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomMPClassDef_LoadQuotationSummaryViewData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCT"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RATETYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IomMPClassDef_LoadQuotationSummaryViewData = _
                                objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomMPClassDef_LoadMortgageProductSelectionData() As FreeThreadedDOMDocument40
'GD BMIDS00034
    
    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductSelectionData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist    'MAR1878 GHun
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCT"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PRODUCTTEXTDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERSNAME<TYPE>dbdtString</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>FIRSTMONTHSREPAYMENT<TYPE>dbdtLong</TYPE></OTHERS>" 'MAR1878 GHun
    'strXML = strXML & "<OTHERS>FIRSTMONTHSREPAYMENTBYDAILY<TYPE>dbdtLong</TYPE></OTHERS>"  'MAR1878 GHun
    strXML = strXML & "<OTHERS>FIRSTMONTHLYINTERESTRATE<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"
    strXML = strXML & "<OTHERS>DISCOUNTEDPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPE<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BMIDS00034
    strXML = strXML & "<OTHERS>DISCOUNTAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    
    'strXML = strXML & "<OTHERS>FIRSTMONTHLYSVRREPAYMENT<TYPE>dbdtLong</TYPE></OTHERS>" 'MAR1878 GHun
    'strXML = strXML & "<OTHERS>FIRSTMONTHLYSVRREPAYBYDAILY<TYPE>dbdtLong</TYPE></OTHERS>"  'MAR1878 GHun
    strXML = strXML & "<OTHERS>FEE<TYPE>dbdtLong</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>PENALTIES<TYPE>dbdtBoolean</TYPE></OTHERS>" 'MAR1878 GHun
    strXML = strXML & "<OTHERS>FLEXIBLEMORTGAGEPRODUCT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NONPANELLENDEROPTION<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERESTRATEENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERESTRATEPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    
    ' EP2_1283
    strXML = strXML & "<OTHERS>FREELEGALFEES<TYPE>dbdtBoolean</TYPE></OTHERS>"
    
    'strXML = strXML & "<OTHERS>TYPICALAPR<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"    'MAR1878 GHun
    
    'PSC 18/09/2002 SYS5077
    strXML = strXML & "<OTHERS>ORGANISATIONID<TYPE>dbdtGuid</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"

    Set IomMPClassDef_LoadMortgageProductSelectionData = _
                                objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomMPClassDef_LoadMortgageProductData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductData"

    Dim strXML As String

'
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCT"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ORGANISATIONID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MPARRANGEMENTFEEAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MPARRANGEMENTFEEPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MPARRANGEMENTFEEPERCENTMAX<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MPARRANGEMENTFEEPERCENTMIN<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AVAILABLEFORMORTGAGECALCULATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AVAILABLEFORSELFEMPLOYED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AVAILABLEQUICKQUOTE<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXISTCUSTOMERMORTGAGEPRODUCT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXAPPLICANTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMTERM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINAPPLICANTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMTERM<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PORTINGFEE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PREFERENTIALRATEPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAFFMORTGAGEPRODUCT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>WITHDRAWNDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    'BMIDS00074 MDC 14/06/2002 - Move BaseRateSet to InterestRateType
    'strXML = strXML & "<OTHERS>BASERATESET<TYPE>dbdtInt</TYPE></OTHERS>"
    'BMIDS00074 MDC 14/06/2002 - End
    strXML = strXML & "<OTHERS>VALUATIONFEESET<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADMINISTRATIONFEESET<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLEXIBLEMORTGAGEPRODUCT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISPLAYORDER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NONPANELLENDEROPTION<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYOFORIGIN<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>IMPAIREDCREDIT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CASHBACKPRODUCT<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MORTGAGEPRODUCTPORTABLEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'CMWP3 - DPF 17/7/02 - added field for redemption fee set
    strXML = strXML & "<OTHERS>REDEMPTIONFEESET<TYPE>dbdtInt</TYPE></OTHERS>"
    'MV - 20/06/2002 - BMIDS00077 - Core Upgrade
    'Start
    'DB SYS4767 - MSMS Integration
    strXML = strXML & "<OTHERS>MANUALDECREASELIMITPERCENT<TYPE>dbdtDouble</TYPE></OTHERS>"
    'DB End
    'End

    'MV - 03/09/2002   - BMIDS00353
    strXML = strXML & "<OTHERS>MPMIGSTARTLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MPMIGRATESET<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RENTALINCOMERATESET<TYPE>dbdtInt</TYPE></OTHERS>" 'JD BMIDS765 CC076
    
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadMortgageProductData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadRentalIncomeRateData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IomMPClassDef_LoadRentalIncomeRateData"
    
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "RENTALINCOMERATESETBAND"
    strXML = strXML & "<PRIMARYKEY>RENTALINCOMERATESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXLOANAMOUNT<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXLTV<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>RENTALINCOMEINTERESTRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set IomMPClassDef_LoadRentalIncomeRateData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadReservedProductData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
'
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomMPClassDef_LoadReservedProductData"

    Dim strXML As String
    
    strXML = "<TABLENAME>RESERVEDPRODUCT" _
    & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadReservedProductData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IomMPClassDef_LoadMortgageProductLanguageData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'MORTGAGEPRODUCTCODE            NOT NULL    VARCHAR2(6)
'STARTDATE                      NOT NULL    DATE
'LANGUAGE                       NOT NULL    NUMBER(5)
'PRODUCTNAME                                VARCHAR2(80)
'PRODUCTTEXTDETAILS                         VARCHAR2(500)

'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductLanguageData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGEPRODUCTLANGUAGE"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>LANGUAGE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PRODUCTTEXTDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadMortgageProductLanguageData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadSpecialGroupData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomMPClassDef_LoadSpecialGroupData"

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "SPECIALGROUP"
    strXML = strXML & "<PRIMARYKEY>GROUPTYPE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadSpecialGroupData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing



End Function

Private Function IomMPClassDef_LoadValuationFeeData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "IomMPClassDef_LoadValuationFeeData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "VALUATIONFEE"
    strXML = strXML & "<PRIMARYKEY>VALUATIONFEESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>VALUATIONFEESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TYPEOFVALUATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ValuationType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMVALUE<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>LOCATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PropertyLocation</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TYPEOFVALUATIONTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOCATIONTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MINIMUMFEEVALUE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXIMUMFEEVALUE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FEEPERCENTAGE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadValuationFeeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomMPClassDef_LoadConditionsData() As FreeThreadedDOMDocument40

    
    Const strFunctionName As String = "IomMPClassDef_LoadConditionsData"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CONDITIONS"
    strXML = strXML & "<PRIMARYKEY>CONDITIONREFERENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CONDITIONNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONDITIONDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "<OTHERS>CONDITIONTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ConditionType</COMBO></OTHERS>"
    
    
    strXML = strXML & "<OTHERS>EDITABLEIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FREEFORMATIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONDITIONIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RULEREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DELETEFLAG<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadConditionsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

'BMIDS00654 MDC 01/11/2002
Private Function IomMPClassDef_LoadMortgageProductIncomeMultiplierSet() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomMPClassDef_LoadMortgageProductIncomeMultiplierSet"
    
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "GETINCOMEMULTIPLIERSETFORPRODUCT"
    strXML = strXML & "<PRIMARYKEY>MORTGAGEPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>STARTDATE<TYPE>dbdtDateTime</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INCOMEMULTIPLIERCODE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomMPClassDef_LoadMortgageProductIncomeMultiplierSet = _
                                objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
'BMIDS00654 MDC 01/11/2002 - End
