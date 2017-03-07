VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreditCheckClassDef"
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
'PSC    28/02/00    Remove Dim As New
'MDC    18/05/00    Added LoadDemoOptionsData method
'SR     08/12/00    included mehods for Full Bureau Download
'------------------------------------------------------------------------------------------

Option Explicit

Implements ICreditCheckClassDef


'------------------------------------------------------------------------------------------
'BMIds History:
'
'Prog   Date        Description
'MDC    07/10/2002  BMIDS00561 Add new full bureau fields
'GD     30/06/2003  BM0592  Include FBJOINTACCOUNT in LoadFullBureauPublicInfoData
'GD     28/07/2003  BMIDS610  Include FBCOURTPLAINTIFF in LoadFullBureauPublicInfoData
'INR    24/05/2004  BMIDS744  Added to LoadApplicationCreditCheckData and
'                   LoadFBResultsData for TPD changes.
'JD     09/08/2004  BMIDS832 add FBMATCHEDTO to LoadFullBureauPublicInfoData
'------------------------------------------------------------------------------------------

Private Function ICreditCheckClassDef_LoadApplicationCreditCheckData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadApplicationCreditCheckData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "APPLICATIONCREDITCHECK"
    
    Select Case intWhatToGet
    Case ccCurrent
        strXML = strXML & "<OTHERS>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
        'BMIDS744
        strXML = strXML & "<OTHERS>TPDOUTCOMECODE<TYPE>dbdtString</TYPE></OTHERS>"
    Case Else
        strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CREDITCHECKREFERENCENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SUCCESSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    End Select
    
    strXML = strXML & "</TABLENAME>"

    Set ICreditCheckClassDef_LoadApplicationCreditCheckData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function ICreditCheckClassDef_LoadCreditCheckCustomerMiscData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadCreditCheckCustomerMiscData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CREDITCHECKCUSTOMERMISC"
    
    Select Case intWhatToGet
    Case ccCurrent
        strXML = strXML & "<OTHERS>NAMESCONFIRMED<TYPE>dbdtString</TYPE></OTHERS>"
    Case Else
        strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>REFERENCESSCURRENT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>REFERENCESSPREVIOUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>REFERENCESPCURRENT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>REFERENCESPPREVIOUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ERYEARSSSCURRENT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ERYEARSSSPREVIOUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ERYEARSSPCURRENT<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ERYEARSSPPREVIOUS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CURRENTADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PREVIOUSADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NAMESCONFIRMED<TYPE>dbdtString</TYPE></OTHERS>"
    End Select
    
    strXML = strXML & "</TABLENAME>"

    Set ICreditCheckClassDef_LoadCreditCheckCustomerMiscData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function ICreditCheckClassDef_LoadCreditCheckDetailsData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadCreditCheckDetailsData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CREDITCHECKDETAILS"
    
    Select Case intWhatToGet
    Case ccCurrent
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>PERSONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALCCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUECCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOUTSTANDINGCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUECAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAISINCMORT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSMORTGAGE6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MORTGAGEACCOUNTTIME<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEOC<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VOTERSROLLPAF<TYPE>dbdtInt</TYPE></OTHERS>"

    Case ccCAIS
        strXML = strXML & "<OTHERS>TOTALACTIVECAIS3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSACTIVECAIS4M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS4M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSACTIVECAIS12M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS12M<TYPE>dbdtInt</TYPE></OTHERS>"
    
    Case Else
        strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
        strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>PERSONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALCCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUECCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUECAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSETTLEDCAIS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSETTLEDCAIS12MONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALDELINQUENTACCOUNTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEDELINQUENTACCOUNTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTDELINQUENTACCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALACTIVECAIS3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSACTIVECAIS4M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS4M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSACTIVECAIS12M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAIS12M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSALLACTIVECAIS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTCURRENTSTATUSACTIVECAIS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALACTIVECAISACCOUNTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAISEXCMORT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEACTIVECAISINCMORT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALNUMBERACTIVECAISSTATUS1<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALNUMBERACTIVECAISSTATUS3<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOCACCOUNTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEOC<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALDELINQUENTOCACCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOCCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSETTLEDOCCAIS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALMOCAIS89<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSETTLEDMOCAIS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSETTLEDMOCAIS12M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALDELINQUENTMOCAISACCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESMOCAIS3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESMOCAIS6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CIFASDETECTED<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUENDCAIS6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUENDCAISALL<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALACTIVENDCAIS6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOCSEARCHES<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOCCAISACCOUNTS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RATIOBALANCE3M12M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSNDH3MLT12M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSNDH3MGT12M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSNDH6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSNDHACTIVE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSNDHNONACTIVE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALNDGOODCAIS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>BANKRUPTCYDETECTED<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CMLDETECTED<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESFH3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESRETAIL3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESBANK3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESCC3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHES6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESOC3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALOUTSTANDINGCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALVALUEOUTSTANDINGCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>AGEOFLASTOUTSTANDINGCCJ<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VOTERSROLLPAF<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESNONMO3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESNONMO6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESMO3M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESMO6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MARKETBASEDACTIVITYIND<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSFCS12M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSTV12M<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CAISSPECIALINSTRUCTIONIND<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GAINCURRENTADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GAINPREVIOUSADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FAMILYADDRESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SINGLESADDRESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MIXEDADDRESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>UNSATISFIEDCCJADDRESSPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CCJPERELECTORPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>DIRECTORPERELECTORPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>UNSATISFIEDCCJGT500PERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>BUILDINGSOCSEARCHPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ELECTRICSEARCHPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATYPESEARCHPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PERSONALLOANSEARCHPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>UTILITIESSEARCHPERCENT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCOREID_2_1<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCOREID_2_2<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCOREID_2_3<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NDDATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MORTGAGEACCOUNTTIME<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MORTGAGEACCOUNTBALANCE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SFREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SFYEARSONVOTERSROLL<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WORSTSTATUSMORTGAGE6M<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALINACTIVECAIS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALNDSEARCHES12M<TYPE>dbdtInt</TYPE></OTHERS>"
    End Select
    
    strXML = strXML & "</TABLENAME>"

    Set ICreditCheckClassDef_LoadCreditCheckDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function ICreditCheckClassDef_LoadCreditCheckMiscellaneousData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadCreditCheckMiscellaneousData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CREDITCHECKMISCELLANEOUS"
    
    Select Case intWhatToGet
    Case ccCurrent
        strXML = strXML & "<OTHERS>SCORECARDID1<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORE1<TYPE>dbdtInt</TYPE></OTHERS>"

    Case Else
        strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
        strXML = strXML & "<OTHERS>ADDRESSTYPE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>WEALTHSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CHILDSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>STABILITYSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>RURABILITYSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>EDUCATIONSCORE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CIFASDETECTED<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VOTERSROLLNOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CCJNOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CIFASNOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CAISNOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OTHERNOC<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SATISFIEDCCJDETECTED<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VOTERSROLLDOBCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>VOTERSROLLDOB<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CAISDOBCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CAISDOB<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SEARCHDOBCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SEARCHDOB<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ALLTYPESDOBCOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ALLTYPESDOB<TYPE>dbdtDate</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MOSAIC93CODE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>MOSAICCODE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PROFESSIONALS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PUBLICTRANSPORT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>ROOMSPERPERSON<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CARSPERHOUSEHOLD<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>GREATERTHANFOURCHILDREN<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OWNEROCCUPIERS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>OVERCROWDING<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FAMILIES<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PSEUDOFAMILIES<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBEROF17TO21YEAROLDS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NUMBERADDRESSESSINCE1981<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>LOWVALUECCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>HIGHVALUECCJS<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHES<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FLATSPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FAMILIESPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CCJSPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>TOTALSEARCHESPERCENTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>POPULATIONINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SUBPOPULATIONINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORECARDPHASE<TYPE>dbdtString</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORECARDID1<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORE1<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORECARDID2<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORE2<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORECARDID3<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>SCORE3<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>NATIONALINSURANCEMOSAICCODE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>FINANCIALMOSAICCODE<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>CURRENTADDRESSNAMECOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
        strXML = strXML & "<OTHERS>PREVIOUSADDRESSNAMECOUNT<TYPE>dbdtInt</TYPE></OTHERS>"
    
    End Select
    
    strXML = strXML & "</TABLENAME>"

    Set ICreditCheckClassDef_LoadCreditCheckMiscellaneousData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


Private Function ICreditCheckClassDef_LoadDemoOptionsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadDemoOptionsData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CREDITCHECKDEMO"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DEMOMODEIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set ICreditCheckClassDef_LoadDemoOptionsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function ICreditCheckClassDef_LoadFullBureauCAISData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "ICreditCheckClassDef_LoadFullBureauCAISData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUCAIS"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBCAISSEQUENCE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISMONTHLYPAYMENT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISCREDITLIMIT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISACCOUNTSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISACCOUNTINSTATUS1OR2<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISACCOUNTINSTATUS3<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSTATUSCOUNTMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSETTLEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISLASTUPDATEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISACCOUNTTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISCOMPANYTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSPECIALINSTRFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISCURRENTDEFAULTBAL<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISREPAYPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSOURCECODEACCNO<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISWORSTSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISCMLADDRESSTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISOWNDATAFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISDEFAULTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS00561 MDC 07/10/2002
    strXML = strXML & "<OTHERS>FBJOINTACCOUNTFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBDATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBBUREAUREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBDEFAULTDATE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREVTERMS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBENDDATE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSEX<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFCSCUGIND<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBBUCKETDATABLOCKS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSTREETMATCH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBHOUSEMATCH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSURNAMEMATCH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFORENAMEMATCH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBBANKFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBMATCHEDTO<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBMATCHTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNODATABLOCKS<TYPE>dbdtInt</TYPE></OTHERS>"
    'BMIDS00561 MDC 07/10/2002 - End
    'BMIDS744
    strXML = strXML & "<OTHERS>FBBUREAUREFCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauCAISData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function ICreditCheckClassDef_LoadFBResultsData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFBResultsData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FBRESULTS"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>FBADDRESSINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNAMEINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBINITIALS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFLAT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSTREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBDISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBTOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCOUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPOSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "<OTHERS>FBCAISSEQUENCE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISACCOUNTSTATUS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSPECIALINSTRFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISOWNDATAFLAG<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAISCMLADDRESSTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSOWNAPPLICATIONIND<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONINFOTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS744
    strXML = strXML & "<OTHERS>FBBUREAUREFCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFBResultsData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
    
End Function

Private Function ICreditCheckClassDef_LoadFullBureauAssociationsData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauAssociationsData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUASSOCIATIONS"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBASSOCIATIONSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBASSOCIATIONINFODATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIAITIONINFOSOURCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONCOTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONINFOTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBASSOCIATIONINITIALS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS744
    strXML = strXML & "<OTHERS>FBBUREAUREFCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauAssociationsData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function ICreditCheckClassDef_LoadFullBureauCAPSData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauCAPSData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUCAPS"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBCAPSSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBCAPSAPPLICATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSAPPLICATIONTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSDATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSTIMEATADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSJOINTAPPLICATIONIND<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSCOMPANYTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSTERMINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCAPSOWNAPPLICATIONIND<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauCAPSData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
    
End Function

Private Function ICreditCheckClassDef_LoadFullBureauCIFASData() As FreeThreadedDOMDocument40
 
    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauCIFASData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUCIFAS"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBCIFASSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBCIFASSUPPLYCONAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCIFASFRAUDCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCIFASSUPPLYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCIFASEXPIRYDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCIFASREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauCIFASData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function ICreditCheckClassDef_LoadFullBureauCorrectionLinesData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauCorrectionLinesData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUCORRECTIONLINES"
    strXML = strXML & "<PRIMARYKEY>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBNOTICEOFCORRECTIONLINENO<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONDETAILS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set ICreditCheckClassDef_LoadFullBureauCorrectionLinesData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function ICreditCheckClassDef_LoadFullBureauDataHeaderData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauDataHeaderData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUDATAHEADER"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBADDRESSINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNAMEINDICATOR<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBINITIALS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBFLAT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBSTREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBDISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBTOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBCOUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPOSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauDataHeaderData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function


Private Function ICreditCheckClassDef_LoadFullBureauPublicInfoData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauPublicInfoData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUPUBLICINFO"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBPUBLICINFOSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBPUBLICINFODATATYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPUBLICINFOAMOUNTINPOUNDS<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPUBLICINFOAMOUNTINPENCE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPUBLICINFOREGISTRATIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBPUBLICINFOSATISFIEDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BM0592 START
    strXML = strXML & "<OTHERS>FBJOINTACCOUNT<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BM0592 END
    'GD BMIDS610 START
    strXML = strXML & "<OTHERS>FBCOURTPLAINTIFF<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BMIDS610 END
    'BMIDS744
    strXML = strXML & "<OTHERS>FBBUREAUREFCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
    'JD BMIDS832 add FBMATCHEDTO
    strXML = strXML & "<OTHERS>FBMATCHEDTO<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauPublicInfoData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function


Private Function ICreditCheckClassDef_LoadFullBureauStandardHeaderData() As FreeThreadedDOMDocument40

Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauStandardHeaderData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUSTANDARDHEADER"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FULLBUREAUDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FULLBUREAUUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauStandardHeaderData = _
                objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
End Function

Private Function ICreditCheckClassDef_LoadFullBureauVotersRollData() As FreeThreadedDOMDocument40

    Const cstrFunctionName As String = "ICreditCheckClassDef_LoadFullBureauVotersRollData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "FULLBUREAUVOTERSROLL"
    strXML = strXML & "<PRIMARYKEY>CREDITCHECKGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBBLOCKID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBHEADERSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>FBVOTERSROLLSEQUENCE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    
    strXML = strXML & "<OTHERS>FBVOTERSROLLAPPLICANTMATCH<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBVOTERSROLLMATCHTYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATE1<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATELEFT1<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATE2<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATELEFT2<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATE3<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBREGISTRATIONDATELEFT3<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FBNOTICEOFCORRECTIONREF<TYPE>dbdtString</TYPE></OTHERS>"
    'BMIDS744
    strXML = strXML & "<OTHERS>FBBUREAUREFCATEGORY<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"
        
    Set ICreditCheckClassDef_LoadFullBureauVotersRollData = objXmlAssist.load(strXML, TypeName(Me), cstrFunctionName)
    
    Set objXmlAssist = Nothing
    
    
End Function

