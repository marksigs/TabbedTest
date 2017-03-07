VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omCEClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for omiga4 Class Definition
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/12/99    Created
'MC     02/06/00    Added LoadIncomeSummaryDetailsData method
'JLD    09/06/00    Added otherincomedetails to LoadUnearnedIncomeData
'JLD    14/06/00    Added NatureOfBusiness to LoadSelfEmployedDetailsData
'MC     13/07/00    SYS1006 Move SHARESOWNEDINDICATOR to Employment table
'MS     14/07/00    SYS0968 MAINSTATUS Tag added for search
'DRC    23/03/01    SYS1932 LoadEmploymentAccountantGUID added
'------------------------------------------------------------------------------------------

Option Explicit

Implements IomCEClassDef



'BMIDS-----------------------------------------------------------------------------
'
'AW 17/10/2002  BMIDS00653   - BM089  Override LoadIncomeSummaryData, LoadIncomeSummaryDetailsData
'                                      Added NETALLOWABLEANNUALINCOME
'DM 15/11/2002  BMIDS00935  Added NOLOCK hint ot the following methods IomCEClassDef_LoadContractDetailsData,
'                           IomCEClassDef_LoadSelfEmployedDetailsData, IomCEClassDef_LoadEarnedIncomeData,
'                           IomCEClassDef_LoadEmploymentData, IomCEClassDef_LoadNetProfitData
'                           IomCEClassDef_LoadUnearnedIncomeData, IomCEClassDef_LoadEmploymentAccountantGUID
'AW 06/12/2002  BM0148      Amended ALLOWABLEMONTHLYINCOME type

'EPSOM---------------------------------------------------------------------------------
'MAH 16/11/2006 E2CR35      Added NATUREOFBUSINESS and COMPANYSTATUSTYPE
'MAH 17/11/2006 E2CR35      Added SELFASSESS to IncomeSummary
'MAH 20/11/2006 E2CR35      Added NUMBEROFRENEWALS to Contract  Details


Private Function IomCEClassDef_LoadAccountantData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadAccountantData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>ACCOUNTANT"
    strXML = strXML & "<PRIMARYKEY>ACCOUNTANTGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTANCYFIRMNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>QUALIFICATIONS<TYPE>dbdtInt</TYPE>"
    strXML = strXML & "<COMBO>AccountantQualifications</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>YEARSACTINGFORCUSTOMER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadAccountantData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadContractDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadContractDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CONTRACTDETAILS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935 Added NOLOCK hint
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>LENGTHOFCONTRACTMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENGTHOFCONTRACTYEARS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LIKELYTOBERENEWED<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFMONTHSEMPLOYED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFYEARSEMPLOYED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTRACTDETAILSSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NUMBEROFRENEWALS<TYPE>dbdtInt</TYPE></OTHERS>" 'MAH 20/11/2006 E2CR35
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadContractDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadContractDetailsProfitData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadContractDetailsProfitData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CONTRACTDETAILSPROFIT"
    strXML = strXML & IomCEClassDef_LoadContractDetailsData.xml
    strXML = strXML & IomCEClassDef_LoadNetProfitData.xml
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadContractDetailsProfitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadDataForEmployment() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadDataForEmployment"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    'MS     14/07/00    SYS0968 MAINSTATUS Tag added for search
    
    strXML = strXML & "<TABLENAME>EMPLOYMENT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>EMPLOYMENTSTATUS"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>EmploymentStatus</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>MAINSTATUS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadDataForEmployment = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadEarnedIncomeData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadEarnedIncomeData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "EARNEDINCOME"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935 BEGIN
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>EARNEDINCOMESEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935 END
    strXML = strXML & "<OTHERS>EARNEDINCOMEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EARNEDINCOMETYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>IncomeType</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTFREQUENCYTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>EmploymentPaymentFreq</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadEarnedIncomeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadEmployedDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadEmployedDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>EMPLOYEDDETAILS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>NOTICEPROBLEMINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYROLLNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PERCENTSHARESHELD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROBATIONARYINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>P60SEENINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>WAGESLIPSSEENINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMPLOYMENTRELATIONSHIPIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>SHARESOWNEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadEmployedDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadEmploymentAccountantGUID() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadEmploymentAccountantGUID"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "EMPLOYMENT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTANTGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadEmploymentAccountantGUID = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadEmployedDetailsIncomeData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadEmployedDetailsIncomeData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>EMPLOYEDDETAILSINCOME"
    strXML = strXML & IomCEClassDef_LoadEmploymentData(omCEEmployment).xml
    strXML = strXML & IomCEClassDef_LoadEmployedDetailsData.xml
    strXML = strXML & IomCEClassDef_LoadEarnedIncomeData.xml
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadEmployedDetailsIncomeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadEmploymentData(Optional ByVal intCEDataDefType As _
                            omCEDataDefType = omCEDefault) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadEmploymentData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "EMPLOYMENT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    If intCEDataDefType = omCEDefault Then
        strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DATELEFTORCEASEDTRADING<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DATESTARTEDORESTABLISHED<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EMPLOYMENTSTATUS"
        strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>EmploymentStatus</COMBO>"
        strXML = strXML & "</OTHERS>"
        strXML = strXML & "<OTHERS>EMPLOYMENTTYPE"
        strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>EmploymentType</COMBO>"
        strXML = strXML & "</OTHERS>"
        'MAH 16/11/2006 E2CR35 Start
        strXML = strXML & "<OTHERS>INDUSTRYTYPE" 'A compromise on account of the Combo IndustryType being HiJacked for something else
        strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>NatureOfBusiness</COMBO>"
        strXML = strXML & "</OTHERS>"
        'strXML = strXML & "<OTHERS>INDUSTRYTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        'MAH 16/11/2006 E2CR35 End
        strXML = strXML & "<OTHERS>MAINSTATUS<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OCCUPATIONTYPE<TYPE>dbdtComboId</TYPE><COMBO>OccupationType</COMBO></OTHERS>"
        strXML = strXML & "<OTHERS>JOBTITLE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OTHEREMPLOYMENTSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NETMONTHLYINCOME<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ACCOUNTANTGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    End If
    strXML = strXML & "<OTHERS>SHARESOWNEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadEmploymentData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadEmploymentDirectoryData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IomCEClassDef_LoadEmploymentDirectoryData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef

    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>EMPLOYMENTDIRECTORY"
    strXML = strXML & IomCEClassDef_LoadEmploymentData.xml
    strXML = strXML & objTPClassDef.LoadDirectoryData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IomCEClassDef_LoadEmploymentDirectoryData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function

Private Function IomCEClassDef_LoadEmploymentThirdPartyData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IomCEClassDef_LoadEmploymentThirdPartyData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objTPClassDef As IThirdPartyClassDef

    Set objTPClassDef = New ThirdPartyClassDef

    strXML = "<TABLENAME>EMPLOYMENTTHIRDPARTY"
    strXML = strXML & IomCEClassDef_LoadEmploymentData.xml
    strXML = strXML & objTPClassDef.LoadThirdPartyData.xml
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IomCEClassDef_LoadEmploymentThirdPartyData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    Set objTPClassDef = Nothing
End Function


Private Function IomCEClassDef_LoadIncomeSummaryData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadIncomeSummaryData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INCOMESUMMARY"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ALLOWABLEANNUALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    'AW 17/10/2002  BMIDS00653   - BM089
    strXML = strXML & "<OTHERS>NETALLOWABLEANNUALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    'AW 06/12/2002  BM0148
    strXML = strXML & "<OTHERS>ALLOWABLEMONTHLYINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NONUKTAXDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXOFFICE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXEDOUTSIDEUKINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UKTAXPAYERINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HIGHESTTAXRATE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>HighestTaxRate</COMBO>"
    strXML = strXML & "</OTHERS>"
    '~ MF MAR30 Added new fields UNDERWRITEROVERRIDEINCLUDEOTHERINC
    '~              and UNDERWRITEROTHERINCONEPERCENTAGE
    'MAH 17/11/2006 E2CR35 Start
    strXML = strXML & "<OTHERS>SELFASSESS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'MAH 17/11/2006 E2CR35 End
    strXML = strXML & "<OTHERS>UNDERWRITEROTHERINCOMEPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNDERWRITEROVERRIDEINCLUDEOTHERINC<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'JD MAR30 add new fields
    strXML = strXML & "<OTHERS>TOTALREGULAROUTGOINGSANNUALOUTGOINGS<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALLOANSLIABILITYANNUALOUTGOINGS<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMORTGAGEANNUALOUTGOINGS<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALLOANSLIABILITYOUTSTANDINGBALANCE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMORTGAGEOUTSTANDINGBALANCE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GROSSCONFIRMEDALLOWABLEINCOME<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NETCONFIRMEDALLOWABLEINCOME<TYPE>dbdtDouble</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadIncomeSummaryData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadIncomeSummaryDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadIncomeSummaryDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INCOMESUMMARYDETAILS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ALLOWABLEANNUALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    'AW 17/10/2002  BMIDS00653   - BM089
    strXML = strXML & "<OTHERS>NETALLOWABLEANNUALINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    'AW 06/12/2002  BM0148
    strXML = strXML & "<OTHERS>ALLOWABLEMONTHLYINCOME<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NONUKTAXDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXOFFICE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TAXEDOUTSIDEUKINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UKTAXPAYERINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HIGHESTTAXRATE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>HighestTaxRate</COMBO>"
    strXML = strXML & "</OTHERS>"
    'MAH 17/11/2006 E2CR35 Start
    strXML = strXML & "<OTHERS>SELFASSESS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'MAH 17/11/2006 E2CR35 End

    strXML = strXML & "<OTHERS>NATIONALINSURANCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    '~ MF MAR30 Added new fields UNDERWRITEROVERRIDEINCLUDEOTHERINC
    '~              and UNDERWRITEROTHERINCOMEPERCENTAGE
    strXML = strXML & "<OTHERS>UNDERWRITEROTHERINCOMEPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNDERWRITEROVERRIDEINCLUDEOTHERINC<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadIncomeSummaryDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function



Private Function IomCEClassDef_LoadNetProfitData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadNetProfitData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "NETPROFIT"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>EVIDENCEOFACCOUNTS<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR1<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR1AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR2<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR2AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR3<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>YEAR3AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadNetProfitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IomCEClassDef_LoadSelfEmployedDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadSelfEmployedDetailsData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>SELFEMPLOYEDDETAILS"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>EMPLOYMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DATEFINANCIALINTERESTHELD<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERBUSINESSCONNECTIONS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PERCENTSHARESHELD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REGISTRATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NATUREOFBUSINESS<TYPE>dbdtString</TYPE></OTHERS>"
    ' MAH 16/11/2006 E2CR35 Start
    strXML = strXML & "<OTHERS>COMPANYSTATUSTYPE"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>CompanyStatusType</COMBO>"
    strXML = strXML & "</OTHERS>"
    ' MAH 16/11/2006 E2CR35 Start
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadSelfEmployedDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadSelfEmployedDetailsProfitData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadSelfEmployedDetailsProfitData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>SELFEMPLOYEDDETAILSPROFIT"
    strXML = strXML & IomCEClassDef_LoadSelfEmployedDetailsData.xml
    strXML = strXML & IomCEClassDef_LoadNetProfitData.xml
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadSelfEmployedDetailsProfitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function IomCEClassDef_LoadUnearnedIncomeData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomCEClassDef_LoadUnearnedIncomeData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>UNEARNEDINCOME"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM 15/11/2002  BMIDS00935
    strXML = strXML & "<PRIMARYKEY>UNEARNEDINCOMESEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>UNEARNEDINCOMEAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UNEARNEDINCOMETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>UnEarnedIncomeDescription</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTFREQUENCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE><COMBO>UnEarnedIncomePaymentFreq</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>OTHERINCOMEDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomCEClassDef_LoadUnearnedIncomeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


