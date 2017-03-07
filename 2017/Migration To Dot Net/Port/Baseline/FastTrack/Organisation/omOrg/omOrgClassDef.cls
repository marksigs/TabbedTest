VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omOrgClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omOrgClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   XML definitions for Class/Tables
'
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS   14/01/2000  Xeneration version branched so modify omiga4 version to use single
'                   interface method
' JLD   22/05/00    Added new fields to LoadMortgageLenderData
' SR    27/06/00    Modified LoadMortgageLenderData
' MV    21/11/00    CORE000011 : Added LoadDistributionChannelData , LoadDepartmentData ,LoadUnitData,
'                   LoadUserList Funcitons
' DJP   19/12/00    Added field to LoadUserData
' MV    21/12/00    SYS1729 :Added a Condition to Check ChannelName in the LoadDepartmentData.
' PSC   12/01/01    SYS1753 LoadMortgageLenderDirectoryData
' DRC   05/04/01    SYS1925 Added LoadActionOwnerData
' JR    04/09/01    Omiplus24, modified LoadContactDetails and added LoadContactTelephoneData
' DB    28/05/02    SYS4767 - MSMS to Core Integration
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        AQR         Description
'MV     21/06/2002  BMIDS00077  Core Upgrade - Amended IOmOrgClassDef_LoadMortgageLenderData
'AW     13/08/2002  BM029       Override IOmOrgClassDef_LoadUnitData
'                               Added IOmOrgClassDef_LoadUnitChequeNo
'MV     04/09/2002  BMIDS00353  Added IOmOrgClassDef_LoanMPMIGRATEBAND
'GD     21/05/2003  BM0370      Change Type of DXID
'SAB    22.09.2005  CORE188     Added WORKGROUPUSER and CREDITCHECKACCESS to IOmOrgClassDef_LoadUserData and
'                               ALLOWOMIGALOGON and ALLOWEXITFROMWRAPUP to IOmOrgClassDef_LoadUnitData
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        AQR     Description
'SR     25/02/2004  BBG77   Modified 'LoadMortgageLenderData' New columns in MortgageLender table
'GHun   14/03/2004  BBG134  CORE49 Amended LoanMortgageLenderParametersData to add extra fields
'AW     21/04/04    WP06 - BBG561   Added LoadLenderNameAddressData, LoadLenderDirectoryData
'                                       LoadLenderAddressData, LoadLenderContactData
'TK     20/08/2004  BBG1192 Amend IomOrgClassDef_LoadLegalFeeData()
'TK     10/09/2004  E2EM00001250 Amend IOmOrgClassDef_LoadMortgageLenderData
'------------------------------------------------------------------------------------------
Option Explicit
Implements IOmOrgClassDef

Private Function IOmOrgClassDef_LoadContactDetailsData(Optional ByVal vemcCONTACTDETAILS As _
                                    CONTACTDETAILS = cdDefault) As FreeThreadedDOMDocument40
 Dim strFunctionName As String
    strFunctionName = "LoadContactDetailsData"
    Dim strXML As String
    Dim strTable As String
    strXML = "<TABLENAME>CONTACTDETAILS" ' We always need this data
    Select Case vemcCONTACTDETAILS
        Case cdDefault ' This is the default - include everything
            strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _
            & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>CONTACTTYPE<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>" _
            & IOmOrgClassDef_LoadContactTelephoneData.xml
            'JR - Omiplus24, removed as part of telephone number change
            '& "<OTHERS>FAXNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            '& "<OTHERS>TELEPHONEEXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            '& "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        Case cdFindUnitSummary
            strXML = strXML & IOmOrgClassDef_LoadContactTelephoneData.xml
            'JR - Omiplus24, removed as part of telephone number change
            'strXML = strXML & "<OTHERS>TELEPHONEEXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" _
            '& "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
        Case Else
                    
            Dim objErrAssist As New ErrAssist
                
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    End Select
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadContactDetailsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadContactTelephoneData() As FreeThreadedDOMDocument40
    
    Dim strFunctionName As String
    strFunctionName = "IContactClassDef_LoadContactTelephoneData"
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    strXML = strXML & "<TABLENAME>CONTACTTELEPHONEDETAILS"
    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TELEPHONESEQNUM<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactTelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TELENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Set IOmOrgClassDef_LoadContactTelephoneData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        
    Set objXmlAssist = Nothing
End Function
Private Function IomOrgClassDef_LoadLegalFeeData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'FEESET                         NOT NULL    NUMBER(5)
'LEGALFEESTARTDATE              NOT NULL    DATE
'FEETYPE                        NOT NULL    NUMBER(5)
'MAXIMUMVALUE                   NOT NULL    NUMBER(10)
'TYPEOFAPPLICATION              NOT NULL    NUMBER(5)
'AMOUNT                                     Number(10)
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "LoadLegalFeeData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "LEGALFEE"
    strXML = strXML & "<PRIMARYKEY>FEESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>LEGALFEESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    'TK 20/8/2004 BBG1192 - strXML = strXML & "<PRIMARYKEY>FEETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<PRIMARYKEY>SPECIALGROUP<TYPE>dbdtComboId</TYPE>"
    'TK 20/08/2004 BBG1192 - strXML = strXML & "<COMBO>LegalFeeType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<COMBO>SpecialGroup</COMBO></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMVALUE<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TYPEOFAPPLICATION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>TypeOfMortgage</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>AMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
     Dim objXmlAssist As New XMLAssist
    Set IomOrgClassDef_LoadLegalFeeData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function

Private Function IOmOrgClassDef_LoadLenderAddressData() As MSXML2.IXMLDOMDocument2
 Dim strFunctionName As String
    strFunctionName = "LoadLenderAddressData"
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "ADDRESS"
    strXML = strXML & "<PRIMARYKEY>ADDRESSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadLenderAddressData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function

Private Function IOmOrgClassDef_LoadLenderDirectoryData() As MSXML2.IXMLDOMDocument2
 Dim strFunctionName As String
    strFunctionName = "LoadLenderDirectoryData"
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGELENDERDIRECTORY"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & IOmOrgClassDef_LoadLenderNameAddressDirectoryData.xml
    strXML = strXML & "</TABLENAME>"
    
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadLenderDirectoryData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function

Private Function IOmOrgClassDef_LoadLenderContactData() As MSXML2.IXMLDOMDocument2

 Dim strFunctionName As String
    strFunctionName = "LoadLenderContactData"
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CONTACTDETAILS"
    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadLenderContactData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing

End Function

Private Function IOmOrgClassDef_LoadLenderNameAddressDirectoryData() As MSXML2.IXMLDOMDocument2
 Dim strFunctionName As String
    strFunctionName = "LoadLenderNameAddressDirectoryData"
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "NAMEANDADDRESSDIRECTORY"
    strXML = strXML & "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COMPANYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & IOmOrgClassDef_LoadLenderContactData.xml
    strXML = strXML & IOmOrgClassDef_LoadLenderAddressData.xml
    strXML = strXML & "</TABLENAME>"
    
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadLenderNameAddressDirectoryData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function

Private Function IOmOrgClassDef_LoadLenderNamesData() As FreeThreadedDOMDocument40
     
    Dim strFunctionName As String
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    strFunctionName = "IOmOrgClassDef_LoadLenderNamesData"
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGELENDER"
'   strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERNAME<TYPE>dbdtString</TYPE></OTHERS>"
    'DB SYS4767 - MSMS Integration
    strXML = strXML & "<OTHERS>ORGANISATIONID<TYPE>dbdtGuid</TYPE></OTHERS>"
    'DB End
    strXML = strXML & "</TABLENAME> "
    Set IOmOrgClassDef_LoadLenderNamesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IomOrgClassDef_LoadMigRateData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'RATESET                        NOT NULL NUMBER(5)
'MIGRATESTARTDATE               NOT NULL DATE
'MAXIMUMLTV                     NOT NULL NUMBER(5)
'MAXIMUMLOANAMOUNT              NOT NULL NUMBER(10)
'INDEMNITYRATE                           NUMBER
' *****************************************************************************
    Dim strFunctionName As String
    strFunctionName = "LoadMigRateData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MIGRATE"
    strXML = strXML & "<PRIMARYKEY>RATESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MIGRATESTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMLTV<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>MAXIMUMLOANAMOUNT<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INDEMNITYRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
     Dim objXmlAssist As New XMLAssist
    Set IomOrgClassDef_LoadMigRateData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadMortgageLenderData(Optional ByVal vemcMortgageLenderData As _
                            MortgageLenderData = MortgageLenderDefault) As FreeThreadedDOMDocument40
    Dim strFunctionName As String
    strFunctionName = "LoadMortgageLenderData"
    Dim strXML As String
    Dim strTable As String
    strXML = "<TABLENAME>MORTGAGELENDER" ' We always need this data
    Select Case vemcMortgageLenderData
        Case LenderNameAndIncomeMultiples ' it does exactly what it says on the tin
            'This is used in a find so it will not raise an error if its missing the primary
            'key value
            strXML = strXML & "<PRIMARYKEY>LENDERNAME<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>SINGLEINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>HIGHESTEARNINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>SECONDEARNINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>JOINTINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>NONGUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>GUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>"
        Case MortgageLenderDefault ' This is the default - include everything
            strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" _
            & "<OTHERS>ADDCOSTSTOLOAN<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>ADDIPTTOMIGPREMIUM<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWADMINFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWARRGEMTFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWDEEDSRELFEEADD<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWMIGFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWPORTINGFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWREINSPTFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWSEALINGFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWTTFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWVALNFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>MLARRANGEMENTFEEAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>MLARRANGEMENTFEEPERCENT<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>MLARRANGEMENTFEEPERCENTMAX<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>MLARRANGEMENTFEEPERCENTMIN<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>DEEDSRELEASEFEE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>IGNOREMIGPREMIUMAMOUNT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>LENDERENDDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
            & "<OTHERS>LENDERNAME<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>LENDERSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>" _
            & "<OTHERS>MAXNOLOANS<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>MIGSTARTLTV<TYPE>dbdtDouble</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>MIGTHRESHOLD<TYPE>dbdtLong</TYPE></OTHERS>" _
            & "<OTHERS>MINIMUMMIGPREMIUM<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>SEALINGFEE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>TTFEE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>SINGLEINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>HIGHESTEARNINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>SECONDEARNINCMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>JOINTINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>NONGUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>GUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWLEGALFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWOTHERFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>RATESET<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>NONGUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>GUARANTEEDINCOMEMULT<TYPE>dbdtDouble</TYPE></OTHERS>" _
            & "<OTHERS>ALLOWREVALUATIONFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>"
            'MV - 20/06/2002 - BMIDS00077 - Core Upgrade
            'Start
            'DB SYS4767 - MSMS Integration
            'SG 19/03/02 SYS4202
            strXML = strXML & "<OTHERS>BACSSORTCODE<TYPE>dbdtDouble</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>BACSACCOUNTNUMBER<TYPE>dbdtDouble</TYPE></OTHERS>"
            'SG 19/03/02 SYS4202
            'STB: MSMS0025 - CHAPS/Hexagon fields.
            strXML = strXML & "<OTHERS>BARCLAYSCHAPSSORTCODE<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>BARCLAYSCHAPSACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>BARCLAYSCHAPSORDERINGINST<TYPE>dbdtString</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>HEXAGONCHAPSACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
            'STB: MSMS0025 - End.
            'DB End
            'End
            
            'SR 25/02/2004 : BBG77 - New columns in MortgageLender table
            strXML = strXML & "<OTHERS>NONLENDINSADMINFEE<TYPE>dbdtDouble</TYPE></OTHERS>"
            strXML = strXML & "<OTHERS>ALLOWNONLENDINSADMINFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>"
            'SR 25/02/2004 : BBG77 - End
            'TK 10/09/2004 E2EM00001250
            strXML = strXML & "<OTHERS>ALLOWPRECOMPSWITCHFEEADDED<TYPE>dbdtBoolean</TYPE></OTHERS>"
        Case Else
                    
            Dim objErrAssist As New ErrAssist
                
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    End Select
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadMortgageLenderData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadActionOwnerData() As FreeThreadedDOMDocument40
    Dim strFunctionName As String
    strFunctionName = "LoadActionOwnerData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>ACTIONOWNER"
    strXML = strXML & "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TASKOWNERTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TASKOWNERTYPETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadActionOwnerData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadCompetencyData() As FreeThreadedDOMDocument40
    Dim strFunctionName As String
    strFunctionName = "LoadCompetencyData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>COMPETENCY"
    strXML = strXML & "<PRIMARYKEY>COMPETENCYTYPE<TYPE>dbdtComboId</TYPE>" & _
                        "<COMBO>CompetencyType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COMPETENCYACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMPETENCYACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FUNDSRELEASEMANDATE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOANAMOUNTMANDATE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LTVMANDATE<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"
    strXML = strXML & "<OTHERS>COMPETENCYDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    'DB SYS4767 - MSMS Integration
    strXML = strXML & "<OTHERS>USERINCREASELIMITPERCENT<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"
    strXML = strXML & "<OTHERS>USERDECREASELIMITPERCENT<TYPE>dbdtDouble</TYPE><FORMATMASK>0.00</FORMATMASK></OTHERS>"
    'DB End
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadCompetencyData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadDepartmentData(Optional ByVal vDDDataView As DEPARTMENTDATA = 0&) As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadDepartmentData"
    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
        
    strXML = "<TABLENAME>DEPARTMENT" _
            & "<PRIMARYKEY>DEPARTMENTID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                    "<OTHERS>CHANNELID<TYPE>dbdtString</TYPE></OTHERS>" & _
                    "<OTHERS>DEPARTMENTACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" & _
                    "<OTHERS>DEPARTMENTACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" & _
                    "<OTHERS>DEPARTMENTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    If vDDDataView = ddFindCurrentUnitList Then
        strXML = strXML & "<OTHERS>CHANNELNAME<TYPE>dbdtString</TYPE></OTHERS>"
    End If
    strXML = strXML & "</TABLENAME>"
    Set IOmOrgClassDef_LoadDepartmentData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadDistributionChannelData() As FreeThreadedDOMDocument40
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim strXML As String
    Const strFunctionName = "IOmOrgClassDef_LoadDistributionChannelData"
    strXML = "<TABLENAME>DISTRIBUTIONCHANNEL" & _
                             "<PRIMARYKEY>CHANNELID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
                             "<OTHERS>CHANNELNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>APPLICATIONNUMBERPREFIX<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>COUNTRYNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" & _
                             "</TABLENAME>"
    Set IOmOrgClassDef_LoadDistributionChannelData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadMortgageLenderDirectoryData() As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "LoadMortgageLenderDirectoryData"
    Const cstrXML As String = _
    "<TABLENAME>" & _
    "MORTGAGELENDERDIRECTORY" & _
    "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
    "<PRIMARYKEY>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
    "<OTHERS>MAINMORTGAGELENDERIND<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
    "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadMortgageLenderDirectoryData = objXmlAssist.load(cstrXML, TypeName(Me), cstrFunctionName)
    Set objXmlAssist = Nothing
       
End Function
Private Function IOmOrgClassDef_LoadMortgageLenderParametersData() As FreeThreadedDOMDocument40
    
    Const strFunctionName As String = "LoadMortgageLenderParametersData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGELENDERPARAMETERS"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCINTERESTIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCINTERESTPAYABLEIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTPAYMENTDUEDATEIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTINGSTARTMONTH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERESTCHARGEDIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCINTERESTROUNDINGFACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCINTERESTROUNDINGDIRECTION<TYPE>dbdtString</TYPE></OTHERS>"
    
    'BBG134 GHun No longer required
    'strXML = strXML & "<OTHERS>DEVONRULINGIND<TYPE>dbdtInt</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>INTERESTCHANGEIND<TYPE>dbdtInt</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>PARTANDPARTROUNDING<TYPE>dbdtInt</TYPE></OTHERS>"
    'BBG134 End
    
    strXML = strXML & "<OTHERS>MONTHLYPAYMENTIND<TYPE>dbdtInt</TYPE></OTHERS>"
    
    'BBG134 GHun No longer required
    'strXML = strXML & "<OTHERS>NETGROSSPROFILEIND<TYPE>dbdtString</TYPE></OTHERS>"
    'BBG134 End
    
    strXML = strXML & "<OTHERS>PAYMENTROUNDINGFACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTROUNDINGDIRECTION<TYPE>dbdtString</TYPE></OTHERS>"
    
    'BBG134 GHun No longer required
    'strXML = strXML & "<OTHERS>INTERESTCALCROUNDINGFACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>INTERESTCALCROUNDINGDIRECTION<TYPE>dbdtString</TYPE></OTHERS>"
    'BBG134 End
    
    strXML = strXML & "<OTHERS>SHIFTINMONTHS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>POSITIONINMONTH<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SHIFTINDAYS<TYPE>dbdtInt</TYPE></OTHERS>"
    
    'BBG134 GHun No longer required
    'strXML = strXML & "<OTHERS>INCLUDEACCINTEREST<TYPE>dbdtInt</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>INCLUDEINACCINTEREST<TYPE>dbdtInt</TYPE></OTHERS>"
    'BBG134 End
    
    'BBG134 GHun New fields required by Alpha+
    strXML = strXML & "<OTHERS>COOLINGOFFPERIOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SPECIFIEDPAYMENTDAY<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DAYSINYEAR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCRUEDDAYSINCLUDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCRUEDDAYSADDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APRMONTH<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>WHENACCINTROUNDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>WHENPAYMENTROUNDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>WHENBALANCEROUNDED<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BALANCEROUNDINGFACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BALANCEROUNDINGDIRECTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CHARGEROUNDINGFACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CHARGEROUNDINGDIRECTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>OUTPUTPAYMENTSCHEDULE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>OutputPaymentSchedule</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>OUTPUTBALANCESCHEDULE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>OutputBalanceSchedule</COMBO></OTHERS>"
    'BBG134 End
    strXML = strXML & "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IOmOrgClassDef_LoadMortgageLenderParametersData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadMPMigRateData() As FreeThreadedDOMDocument40
    
    Dim strFunctionName As String
    strFunctionName = "IOmOrgClassDef_LoadMPMigRateData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MPMIGRATEBAND"
    strXML = strXML & "<PRIMARYKEY>MPMIGRATESET<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>EFFECTIVEDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOWERLTV<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>HIGHERLTV<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>LOWERLOANAMOUNT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<PRIMARYKEY>HIGHERLOANAMOUNT<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>MINIMUMPREMIUM<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INDEMNITYRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
     Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadMPMigRateData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUnitData() As FreeThreadedDOMDocument40
    '   AW  13/08/02    BM029   -   Added UNHIGHCHEQUENUMBER
    'Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    Const strFunctionName = "IOmOrgClassDef_LoadUnitData"
    Const strXML As String = "<TABLENAME>UNIT" & _
                             "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
                             "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>" & _
                             "<OTHERS>DEPARTMENTID<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>UNITACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" & _
                             "<OTHERS>UNITACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" & _
                             "<OTHERS>DXID<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>ORGANISATIONINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                             "<OTHERS>PROCESSINGUNITINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                             "<OTHERS>UNITNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>UNHIGHCHEQUENUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>ALLOWOMIGALOGON<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                             "<OTHERS>ALLOWOMIGAEXITFROMWRAP<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
                             "</TABLENAME>"
                            ''GD BM0370 Changed Type of DXID
    Set IOmOrgClassDef_LoadUnitData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Public Function IOmOrgClassDef_LoadUnitSummaryData(Optional ByVal vemcUNITSUMMARY As _
                                    UNITSUMMARY = usDefault) As FreeThreadedDOMDocument40
    Dim strFunctionName As String
    strFunctionName = "LoadUnitSummaryData"
    Dim strXML As String
    Dim strTable As String
    strXML = "<TABLENAME>UNIT" ' We always need this data
    Select Case vemcUNITSUMMARY
        Case usDefault ' This is the default - include everything
            strXML = strXML & "<PRIMARYKEY>ADDRESSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" _
            & "<PRIMARYKEY>DEPARTMENTID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
            & "<OTHERS>UNITACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" _
            & "<OTHERS>UNITACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" _
            & "<OTHERS>DXID<TYPE>dbdtInt</TYPE></OTHERS>" _
            & "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>" _
            & "<OTHERS>ORGANISATIONINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>PROCESSINGUNITINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" _
            & "<OTHERS>UNITNAME<TYPE>dbdtString</TYPE></OTHERS>"
        Case usFindUnitSummary
            strXML = strXML & "<OTHERS>UNITNAME<TYPE>dbdtString</TYPE></OTHERS>"
        Case Else
                    
            Dim objErrAssist As New ErrAssist
                
            objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingLoadDataEnum
    End Select
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadUnitSummaryData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUserNameData() As FreeThreadedDOMDocument40
    
    Dim objXmlAssist As New XMLAssist
    Const strFunctionName = "IOmOrgClassDef_LoadUserNameData"
    Const strXML As String = "<TABLENAME>USER" & _
                             "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
                             "<OTHERS>USERFORENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>USERSURNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "<OTHERS>USERTITLE<TYPE>dbdtComboId</TYPE>" & _
                             "<COMBO>Title</COMBO></OTHERS>" & _
                             "</TABLENAME>"
    Set IOmOrgClassDef_LoadUserNameData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUnitNameData() As FreeThreadedDOMDocument40
    
    Dim objXmlAssist As New XMLAssist
    Const strFunctionName = "IOmOrgClassDef_LoadUserNameData"
    Const strXML As String = "<TABLENAME>UNIT" & _
                             "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
                             "<OTHERS>UNITNAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "</TABLENAME>"
    Set IOmOrgClassDef_LoadUnitNameData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUserData(Optional ddUserData As UserData = udDefault) As FreeThreadedDOMDocument40
    
    Dim objXmlAssist As New XMLAssist
    Const strFunctionName = "IOmOrgClassDef_LoadUserData"
    Dim strXML As String
    strXML = "<TABLENAME>OMIGAUSER" & _
            "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
            "<OTHERS>ACCESSTYPE<TYPE>dbdtComboId</TYPE>" & _
            "<COMBO>UserAccessType</COMBO></OTHERS>" & _
            "<OTHERS>OMIGAUSERACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" & _
            "<OTHERS>OMIGAUSERACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" & _
            "<OTHERS>NOTES<TYPE>dbdtString</TYPE></OTHERS>" & _
            "<OTHERS>CHANGEPASSWORDINDICATOR<TYPE>dbdtInt</TYPE></OTHERS>" & _
            "<OTHERS>WORKINGHOURTYPE<TYPE>dbdtInt</TYPE></OTHERS>" & _
            "<OTHERS>CREDITCHECKACCESS<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
            "<OTHERS>WORKGROUPUSER<TYPE>dbdtBoolean</TYPE></OTHERS>"
            If ddUserData = udFindUserList Then
                strXML = strXML & "<OTHERS>USERFORENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
                "<OTHERS>USERSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
            End If
            strXML = strXML & "</TABLENAME>"
    Set IOmOrgClassDef_LoadUserData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUserRoleData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
'RF 11/11/99 Removed CONTACTDETAILSGUID to match database change.
'--------------------------------------------------------------------------------
       
    Const strFunctionName As String = "LoadUserRoleData"
    Const strXML As String = "<TABLENAME>USERROLE" _
    & "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<OTHERS>USERROLEACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>USERROLEACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>ROLE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>UserRole</COMBO></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadUserRoleData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadOrganisationUserData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadOrganisationUserData"
    Const strXML As String = "<TABLENAME>ORGANISATIONUSER" _
    & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>EMERGENCYPASSWORD<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>MOTHERMAIDENNAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>USERFORENAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>USERINITIALS<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>USERSURNAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>USERTITLE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>Title</COMBO></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadOrganisationUserData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadIntermediaryUserData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadIntermediaryUserData"
    Const strXML As String = "<TABLENAME>INTERMEDIARYUSER" _
    & "<PRIMARYKEY>INTERMEDIARYUSERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<OTHERS>ORGANISATIONUSERID<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadIntermediaryUserData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadIntermediaryIndividualData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadIntermediaryIndividualData"
    Const strXML As String = "<TABLENAME>INTERMEDIARYINDIVIDUAL" _
    & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" _
    & "<OTHERS>FORENAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>Title</COMBO></OTHERS>" _
    & "<OTHERS>USERROLEUSERID<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>USERROLEUNITID<TYPE>dbdtString</TYPE></OTHERS>" _
    & "<OTHERS>KEYDEVELOPMENTAREA<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "</TABLENAME>"
      
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadIntermediaryIndividualData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadQualificationData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadIntermediaryIndividualData"
    Const strXML As String = "<TABLENAME>QUALIFICATION" _
    & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>QUALIFICATIONTYPE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>UserQualification</COMBO></PRIMARYKEY>" _
    & "<OTHERS>QUALIFICATIONACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "<OTHERS>QUALIFICATIONACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadQualificationData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUserCompetencyData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadUserCompetencyData"
    Const strXML As String = "<TABLENAME>USERCOMPETENCY" _
    & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>COMPETENCYTYPE<TYPE>dbdtComboId</TYPE>" _
    & "<COMBO>CompetencyType</COMBO></PRIMARYKEY>" _
    & "<OTHERS>USERCOMPETENCYACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadUserCompetencyData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadPasswordData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadPasswordData"
    Const strXML As String = "<TABLENAME>PASSWORD" _
    & "<PRIMARYKEY>USERID<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>PASSWORDVALUE<TYPE>dbdtString</TYPE></PRIMARYKEY>" _
    & "<OTHERS>PASSWORDCREATIONDATE<TYPE>dbdtDateTime</TYPE></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadPasswordData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
        
End Function
Private Function IOmOrgClassDef_LoadWorkingHoursData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "LoadWorkingHoursData"
    Const strXML As String = "<TABLENAME>WORKINGHOURS" _
    & "<PRIMARYKEY>WORKINGHOURTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<PRIMARYKEY>DAYOFTHEWEEK<TYPE>dbdtInt</TYPE></PRIMARYKEY>" _
    & "<OTHERS>FROMHOUR<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>FROMMINUTE<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TOHOUR<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "<OTHERS>TOMINUTE<TYPE>dbdtInt</TYPE></OTHERS>" _
    & "</TABLENAME>"
    Dim objXmlAssist As New XMLAssist
    Set IOmOrgClassDef_LoadWorkingHoursData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadAllowableIncomeFactorsData() As FreeThreadedDOMDocument40
    
    Dim strFunctionName As String
    strFunctionName = "LoadAllowableIncomeFactorsData"
    Dim strXML As String
    Dim strTable As String
    strXML = "<TABLENAME>ALLOWABLEINCOMEFACTORS"
    strXML = strXML & "<PRIMARYKEY>ORGANISATIONID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>INCOMEGROUP<TYPE>dbdtComboId</TYPE>" & _
                      "<COMBO>IncomeGroupType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>FACTOR<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
        
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IOmOrgClassDef_LoadAllowableIncomeFactorsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IOmOrgClassDef_LoadUnitChequeNo() As FreeThreadedDOMDocument40
'   AW     13/08/2002  BM029
    Dim objXmlAssist As New XMLAssist
    Const strFunctionName = "IOmOrgClassDef_LoadUnitChequeNo"
    Const strXML As String = "<TABLENAME>UNIT" & _
                             "<PRIMARYKEY>UNITID<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
                             "<OTHERS>UNHIGHCHEQUENUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
                             "</TABLENAME>"
    Set IOmOrgClassDef_LoadUnitChequeNo = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
