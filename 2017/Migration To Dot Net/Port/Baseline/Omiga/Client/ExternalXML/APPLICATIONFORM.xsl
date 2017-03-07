<?xml version="1.0" encoding="UTF-8"?>
<!-- edited with XMLSpy v2005 rel. 3 U (http://www.altova.com) by ﻿ITS (Marlborough Stirling plc) -->

<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	DS		20/02/2007   EP2_1165	Added an NEWLOAN.DIRECTFINANCIALBENEFITIND Details to APPLICATIONFORM
	OS		08/03/2007	 EP2_972		Incorporated combovalues for various fields (Country of residence), corrected XSL statements
	OS		19/03/2007	 EP2_972		Made changes according to updated specs
	OS		27/03/2007	 EP2_972		Made changes according to updated specs
	OS		13/04/2007	 EP2_972		Corrected display of non monetary numbers, fixed arrears.
	OS		16/04/2007	 EP2_2407	Fixed "Fee payment",  "Adviser's Declaration" and "Adviser's Details" sections
	OS		20/04/2007	 EP2_2502	Fixed various issues, incorporated spec changes
	PE		21/04/2007	EP2_2502	Fixed more various issues, incorporated spec changes
	PE		21/04/2007	EP2_2502	Fixed more various issues, incorporated spec changes
	===============================================================================================================-->


<!-- XSLT for Application Form hard copy -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<xsl:key name="elements" match="*" use="concat(name(), '::', .)"/>
	<msxsl:script language="JavaScript" implements-prefix="msg"><![CDATA[

	function HandleYesNo(decisionFlag)
	{
			if(decisionFlag == '1')
				return "Yes"
			else if (decisionFlag == '0')
				return"No"
			else
				return decisionFlag
	}
	function GetCustomerType(customerOrder, customerRoleType)
	{
		var customerType = 0
		if(customerOrder == '1' && customerRoleType == '1')
			customerType = 1
		else if(customerOrder == '2' && customerRoleType == '1')
			customerType = 2
		if(customerRoleType == '2')
				customerType = 3
		return customerType
	}				
	function GetHeading(customerOrder, customerRoleType)
	{
	var numberArray = new Array(
														"",
														"First",
														"Second",
														"Third",
														"Fourth",																																										
														"Fifth",
														"Sixth",
														"Seventh",
														"Eighth",
														"Ninth"															
														)
	var applicantTypeArray = new Array(	
																"",
																"Applicant",
																"Guarantor"
																)
	return numberArray[customerOrder] + " " + applicantTypeArray[customerRoleType]
	}
	
	function HandleAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, Country)
        {
		var strAddress = "";
			
        if (FlatNumber != '')
		{
			strAddress = strAddress  + 'Flat ' + FlatNumber + ', ';
		}
        if (HouseName != '')
		{
			strAddress = strAddress + HouseName + ', ' ;
		}
        if (HouseNumber != '')
		{
			strAddress = strAddress + HouseNumber + ' ';
		}
        if (Street != '')
		{
			strAddress = strAddress + Street + ', ';;
		}
        if (District != '')
		{
			strAddress = strAddress + District + ', ';			
		}
        if (Town != '')
		{
			strAddress = strAddress + Town + ', ';
		}
        if (County != '')
		{
			strAddress = strAddress + County + ', ';
		}
		if (Country != '')
		{
			strAddress = strAddress +  Country;
		}		
		return (strAddress);
        }   

	function HandlePhoneNumber(CountryCode, AreaCode, TeleohoneNumber, ExtensionNumber)
	{
		var strPhone= "";
			
        if (CountryCode != '')
		{
			strPhone = strPhone   + CountryCode + ' '	
		}
		 if (AreaCode != '')
		{
			strPhone = strPhone   + AreaCode + ' '	
		}
		 if (TeleohoneNumber != '')
		{
			strPhone = strPhone   + TeleohoneNumber + ' '	
		}
		 if (ExtensionNumber != '')
		{
			strPhone = strPhone   +  'ext: ' + ExtensionNumber	
		}
		
		return strPhone
	}
    ]]></msxsl:script>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="RESPONSE" select="/RESPONSE/APPLICATIONFORM"/>
			<xsl:element name="APPLICATIONFORM">
				<xsl:element name="QUESTIONS">
					<xsl:attribute name="SUBMITTYPE"><xsl:value-of select="string('Yes')"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/@APPLICATIONNUMBER),'}')"/></xsl:attribute>
				</xsl:element>
				<!-- Adviser details -->
				<xsl:element name="ADVISERDETAILS">
					<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']">
						<xsl:variable name="FlatNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@FLATNUMBER"/>
						<xsl:variable name="HouseName" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@BUILDINGORHOUSENAME"/>
						<xsl:variable name="HouseNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@BUILDINGORHOUSENUMBER"/>
						<xsl:variable name="District" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@DISTRICT"/>
						<xsl:variable name="Street" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@STREET"/>
						<xsl:variable name="Town" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@TOWN"/>
						<xsl:variable name="County" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@COUNTY"/>
						<xsl:variable name="Country" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@COUNTRY_TEXT"/>
						<xsl:variable name="Address" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
						<xsl:attribute name="BROKERCASEREF"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER),'}')"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
						<xsl:attribute name="POSTCODE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@POSTCODE"/></xsl:attribute>
						<xsl:attribute name="FULLNAME"><xsl:value-of select="concat(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/ORGANISATIONUSER/@USERFORENAME), ' ', string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/ORGANISATIONUSER/@USERSURNAME))"/></xsl:attribute>
						<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER[@INTRODUCERTYPE='1']/@EMAILADDRESS)"/></xsl:attribute>
						<xsl:choose>
							<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM">
								<!-- process AR Firm here -->
								<xsl:attribute name="FSAFIRMREF"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@FSAARFIRMREF),'}')"/></xsl:attribute>
								<xsl:attribute name="COMPANYNAME"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@ARFIRMNAME)"/></xsl:attribute>
								<xsl:variable name="CountryCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@TELEPHONECOUNTRYCODE"/>
								<xsl:variable name="AreaCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@TELEPHONEAREACODE"/>
								<xsl:variable name="TelephoneNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@TELEPHONENUMBER"/>
								<xsl:attribute name="OFFICETELEPHONE"><xsl:value-of select="concat('{',msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($TelephoneNumber), ''),'}')"/></xsl:attribute>
								<xsl:variable name="FaxCountryCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@FAXCOUNTRYCODE"/>
								<xsl:variable name="FaxAreaCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@FAXAREACODE"/>
								<xsl:variable name="FaxTelephoneNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/@FAXNUMBER"/>
								<xsl:attribute name="FAXNUMBER"><xsl:value-of select="concat('{',msg:HandlePhoneNumber(string($FaxCountryCode), string($FaxAreaCode), string($FaxTelephoneNumber), ''),'}')"/></xsl:attribute>
								<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/ARFIRM/FIRMPERMISSIONS/ACTIVITYFSA">
									<xsl:element name="FRMPERMISSIONSET">
										<xsl:attribute name="FRMPERMISSION"><xsl:value-of select="./@ACTIVITYDESCRIPTION"></xsl:value-of></xsl:attribute>
									</xsl:element>
								</xsl:for-each>
							</xsl:when>
							<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']">
								<!-- process DA Firm here -->
								<xsl:attribute name="FSAFIRMREF"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@FSAREF),'}')"/></xsl:attribute>
								<xsl:attribute name="COMPANYNAME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@PRINCIPALFIRMNAME"/></xsl:attribute>
								<xsl:variable name="CountryCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@TELEPHONECOUNTRYCODE"/>
								<xsl:variable name="AreaCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@TELEPHONEAREACODE"/>
								<xsl:variable name="TelephoneNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@TELEPHONENUMBER"/>
								<xsl:attribute name="OFFICETELEPHONE"><xsl:value-of select="concat('{',msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($TelephoneNumber), ''),'}')"/></xsl:attribute>
								<xsl:variable name="FaxCountryCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@FAXCOUNTRYCODE"/>
								<xsl:variable name="FaxAreaCode" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@FAXAREACODE"/>
								<xsl:variable name="FaxTelephoneNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/@FAXNUMBER"/>
								<xsl:attribute name="FAXNUMBER"><xsl:value-of select="concat('{',msg:HandlePhoneNumber(string($FaxCountryCode), string($FaxAreaCode), string($FaxTelephoneNumber), ''),'}')"/></xsl:attribute>
								<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='0']/FIRMPERMISSIONS/ACTIVITYFSA">
									<xsl:element name="FRMPERMISSIONSET">
										<xsl:attribute name="FRMPERMISSION"><xsl:value-of select="./@ACTIVITYDESCRIPTION"></xsl:value-of></xsl:attribute>
									</xsl:element>
								</xsl:for-each>
							</xsl:when>
						</xsl:choose>
					</xsl:if>
				</xsl:element>
				<!-- adviser details ends here-->
				<!-- Application Route-->
				<xsl:element name="APPLICATIONROUTE">
					<xsl:choose>
						<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/INTRODUCER">
							<xsl:choose>
								<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='1']">
									<xsl:attribute name="PACKAGINGTYPE"><xsl:value-of select="string('Packaged >')"/></xsl:attribute>
									<!-- populate principal firm record as introducer-->
									<xsl:choose>
										<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/MORTGAGECLUBNETWORKASSOCIATION">
											<xsl:attribute name="CLUBORNETWORK"><xsl:value-of select="concat('with ',$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/MORTGAGECLUBNETWORKASSOCIATION/@MORTGAGECLUBNETWORKASSOCNAME,' >')"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="CLUBORNETWORK"><xsl:value-of select="string(' no association selected >')"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM[@PACKAGERINDICATOR='1']/@PRINCIPALFIRMNAME)"/></xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<!-- Unpackaged -->
									<xsl:attribute name="PACKAGINGTYPE"><xsl:value-of select="string('Unpackaged >')"/></xsl:attribute>
									<xsl:choose>
										<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/MORTGAGECLUBNETWORKASSOCIATION">
											<!-- with Mortgage Club -->
											<xsl:attribute name="CLUBORNETWORK"><xsl:value-of select="concat('with ',$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/MORTGAGECLUBNETWORKASSOCIATION/@MORTGAGECLUBNETWORKASSOCNAME,' >')"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="CLUBORNETWORK"><xsl:value-of select="string('- Direct to DB')"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="PACKAGINGTYPE"><xsl:value-of select="string('- Direct to DB - No third party')"/></xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:element>
				<!-- Application route ends here-->
				<xsl:element name="SOURCEOFINTRODUCTION">
					<!-- Process Packager details-->
					<xsl:attribute name="PRINCIPALFSANUMBER"><xsl:value-of select="concat('{', $RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM/@FSAREF, '}')"/></xsl:attribute>
					<xsl:attribute name="PACKAGERCASEREF"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER),'}')"/></xsl:attribute>
					<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCER/PRINCIPALFIRM">
						<xsl:element name="SOURCEOFINTRODUCTIONMULTIPLE">
							<xsl:variable name="AddressLine1" select="@ADDRESSLINE1"/>
							<xsl:variable name="AddressLine2" select="@ADDRESSLINE2"/>
							<xsl:variable name="AddressLine3" select="@ADDRESSLINE3"/>
							<xsl:variable name="AddressLine4" select="@ADDRESSLINE4"/>
							<xsl:variable name="AddressLine5" select="@ADDRESSLINE5"/>
							<xsl:variable name="AddressLine6" select="@ADDRESSLINE6"/>
							<xsl:variable name="Address" select="msg:HandleAddress('', string($AddressLine1), '', string($AddressLine2), string($AddressLine3), string($AddressLine4), string($AddressLine5), string($AddressLine6))"/>
							<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="@PRINCIPALFIRMNAME"/></xsl:attribute>
							<xsl:attribute name="ADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
							<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
							<xsl:attribute name="OFFICETELEPHONE"><xsl:value-of select="concat(string(@TELEPHONECOUNTRYCODE), ' ' , string(@TELEPHONEAREACODE), ' ' , string(@TELEPHONENUMBER))"/></xsl:attribute>
							<xsl:attribute name="FAXNUMBER"><xsl:value-of select="concat(string(@FAXCOUNTRYCODE), ' ' , string(@FAXAREACODE), ' ' , string(@FAXNUMBER))"/></xsl:attribute>
						</xsl:element>							
					</xsl:for-each>
				</xsl:element>
				<!-- Adviser declaration section -->
				<xsl:element name="ADVISERDECLARATION">
					<xsl:attribute name="REGULATED"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@REGULATIONINDICATOR_TEXT"/></xsl:attribute>
					<xsl:attribute name="LEVELOFADVICE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@LEVELOFADVICE_TEXT"/></xsl:attribute>
					<xsl:variable name="vTotalProcurationFee">
						<xsl:choose>
							<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TOTALPROCURATIONFEE and $RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TOTALPROCURATIONFEE != '0'">
								<xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TOTALPROCURATIONFEE"></xsl:value-of>
							</xsl:when>
							<xsl:otherwise><xsl:value-of select="'None'"></xsl:value-of></xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<xsl:attribute name="TOTALPROCURATIONFEE">
						<xsl:value-of select="$vTotalProcurationFee"></xsl:value-of>
					</xsl:attribute>
					<xsl:attribute name="TOTALPROCURATIONFEESIGN">
						<xsl:if test="$vTotalProcurationFee != 'None'">
							<xsl:value-of select="'£'"></xsl:value-of>							
						</xsl:if> 
					</xsl:attribute>
					<xsl:attribute name="KFIRECEIVEDBYALLAPPS"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@KFIRECEIVEDBYALLAPPS))"/></xsl:attribute>
				</xsl:element>
				<!-- Adviser declaration section ends -->
				<!-- Feelist section -->
				<xsl:element name="FEELIST">
					<xsl:for-each select=".//QUOTATION/MORTGAGESUBQUOTE/MORTGAGEONEOFFCOST[not(contains(@MORTGAGEONEOFFCOSTTYPE_TYPES,'TID'))]">
						<xsl:if test="./@AMOUNT and ./@AMOUNT != '0'">
							<xsl:element name="FEETYPE">
								<xsl:variable name="OneOffCostType" select="./@MORTGAGEONEOFFCOSTTYPE"/>
								<xsl:variable name="ADDTOLOAN" select="./@ADDTOLOAN"/>
								<xsl:attribute name="MORTGAGEONEOFFCOSTTYPE"><xsl:value-of select="./@MORTGAGEONEOFFCOSTTYPE_TEXT"></xsl:value-of></xsl:attribute>
								<xsl:attribute name="AMOUNT"><xsl:value-of select="./@AMOUNT"/></xsl:attribute>
								<xsl:attribute name="ADDTOLOAN"><xsl:value-of select="msg:HandleYesNo(string(./@ADDTOLOAN))"/></xsl:attribute>
								<xsl:attribute name="WHENPAYABLE">
									<xsl:choose>
										<xsl:when test="contains(@MORTGAGEONEOFFCOSTTYPE_TYPES,'END')">
											<xsl:text>Redemption</xsl:text>
										</xsl:when>
										<xsl:when test="contains(@MORTGAGEONEOFFCOSTTYPE_TYPES,'COMP')">
											<xsl:text>Completion</xsl:text>
										</xsl:when>
										<xsl:when test="contains(@MORTGAGEONEOFFCOSTTYPE_TYPES,'PAID')">
											<xsl:text>Already Paid</xsl:text>
										</xsl:when>
										<xsl:otherwise>
											<xsl:if test="@ADDTOLOAN=1">
												<xsl:text>Added to Loan</xsl:text>
											</xsl:if>
											<xsl:if test="not(@ADDTOLOAN) or @ADDTOLOAN!=1">
												<xsl:text>Paid on Application</xsl:text>
											</xsl:if>											
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<xsl:attribute name="REFUNDABLE">
									<xsl:choose>
										<xsl:when test="./@REFUNDAMOUNT">
											<xsl:value-of select="'Yes'"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="'No'"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
							</xsl:element>
						</xsl:if>
					</xsl:for-each>
				</xsl:element>
				<!-- Feelist section ends -->
				<!-- Product Details  section -->
				<xsl:element name="PRODUCTDETAILS">
					<xsl:attribute name="APPLICATIONINCOMESTATUS"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@APPLICATIONINCOMESTATUS_TEXT)"/></xsl:attribute>
					<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TEXT)"/></xsl:attribute>
					<xsl:attribute name="NATUREOFLOAN"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@NATUREOFLOAN_TEXT)"/></xsl:attribute>
					<xsl:attribute name="PRODUCTSCHEME"><xsl:value-of select="string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@PRODUCTSCHEME_TEXT)"/></xsl:attribute>
					<xsl:apply-templates select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT"/>					
				</xsl:element>
				<!-- Product Details  section ends -->
				<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE">
					<xsl:sort select="./@CUSTOMERORDER"/>
					<xsl:variable name="Heading" select="msg:GetHeading(string(./@CUSTOMERORDER),string(./@CUSTOMERROLETYPE))"/>
					<xsl:variable name="CustomerType" select="msg:GetCustomerType(string(./@CUSTOMERORDER),string(./@CUSTOMERROLETYPE))"/>
					<!-- Customer Information section -->
					<xsl:element name="CUSTOMERINFORMATION">
						<xsl:attribute name="HEADING"><xsl:value-of select="$Heading"/></xsl:attribute>
						<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@FIRSTFORENAME)"/></xsl:attribute>
						<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@SECONDFORENAME)"/></xsl:attribute>
						<xsl:attribute name="SURNAME"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@SURNAME)"/></xsl:attribute>
						<xsl:attribute name="TITLE"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@TITLE_TEXT)"/></xsl:attribute>
						<xsl:attribute name="GENDER"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@GENDER_TEXT)"/></xsl:attribute>
						<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@DATEOFBIRTH)"/></xsl:attribute>
						<xsl:attribute name="MARITALSTATUS"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@MARITALSTATUS_TEXT)"/></xsl:attribute>
						<xsl:attribute name="NUMBEROFDEPENDANTS"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@NUMBEROFDEPENDANTS)"/></xsl:attribute>
						<xsl:attribute name="NATIONALITY"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@NATIONALITY_TEXT)"/></xsl:attribute>
						<xsl:attribute name="UKRESIDENT"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/@UKRESIDENTINDICATOR))"/></xsl:attribute>
						<xsl:attribute name="STAYWHERE"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@COUNTRYOFRESIDENCY_TEXT)"/></xsl:attribute>
						<xsl:attribute name="REASONOFSTAY"><xsl:value-of select="string(./CUSTOMER/CUSTOMERVERSION/@REASONFORCOUNTRYOFRESIDENCY)"/></xsl:attribute>
						<xsl:attribute name="RIGHTTOWORK"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/@RIGHTTOWORK))"/></xsl:attribute>
						<xsl:attribute name="DIPLOMATICIMMUNITY"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/@DIPLOMATICIMMUNITY))"/></xsl:attribute>
						<xsl:attribute name="HOUSINGBENEFIT"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/@HOUSINGBENEFIT))"/></xsl:attribute>
						<xsl:choose>
							<xsl:when test="$CustomerType='1'">
								<xsl:attribute name="RIGHTTOWORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='610']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="DIPLOMATICIMMUNITYDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='620']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="SELLINGINTENSIONS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='640']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="HOUSINGBENEFITDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='630']/@MEMOENTRY"/></xsl:attribute>
							</xsl:when>
							<xsl:when test="$CustomerType='2'">
								<xsl:attribute name="RIGHTTOWORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='611']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="DIPLOMATICIMMUNITYDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='621']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="SELLINGINTENSIONS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='641']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="HOUSINGBENEFITDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='631']/@MEMOENTRY"/></xsl:attribute>
							</xsl:when>
							<xsl:when test="$CustomerType='3'">
								<xsl:attribute name="RIGHTTOWORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='612']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="DIPLOMATICIMMUNITYDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='622']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="SELLINGINTENSIONS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='642']/@MEMOENTRY"/></xsl:attribute>
								<xsl:attribute name="HOUSINGBENEFITDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='632']/@MEMOENTRY"/></xsl:attribute>
							</xsl:when>
						</xsl:choose>
						<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/ALIAS">
							<xsl:element name="OTHERNAME">
								<xsl:attribute name="OTHERNAME"><xsl:value-of select="concat(string(./PERSON/@FIRSTFORENAME), ' ',string(./PERSON/@SECONDFORENAME),' ',string(./PERSON/@SURNAME))"/></xsl:attribute>
								<xsl:attribute name="PREVTITLE"><xsl:value-of select="string(./PERSON/@TITLE_TEXT)"/></xsl:attribute>
								<xsl:attribute name="DATEOFNAMECHANGE"><xsl:value-of select="string(./@DATEOFCHANGE)"/></xsl:attribute>
							</xsl:element>
						</xsl:for-each>
						<!-- Customer Address section  -->
						<xsl:element name="ADDRESSES">
							<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/CUSTOMERADDRESS">
								<xsl:variable name="Address_Type" select="./@ADDRESSTYPE"/>
								<xsl:variable name="FlatNumber" select="./ADDRESS/@FLATNUMBER"/>
								<xsl:variable name="HouseName" select="./ADDRESS/@BUILDINGORHOUSENAME"/>
								<xsl:variable name="HouseNumber" select="./ADDRESS/@BUILDINGORHOUSENUMBER"/>
								<xsl:variable name="District" select="./ADDRESS/@DISTRICT"/>
								<xsl:variable name="Street" select="./ADDRESS/@STREET"/>
								<xsl:variable name="Town" select="./ADDRESS/@TOWN"/>
								<xsl:variable name="County" select="./ADDRESS/@COUNTY"/>
								<xsl:variable name="Country" select="./ADDRESS/@COUNTRY_TEXT"/>
								<xsl:variable name="Address" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
								<xsl:if test="($Address_Type = '1')">
									<xsl:element name="CURRENTADDRESS">
										<xsl:attribute name="ADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
										<xsl:attribute name="POSTCODE"><xsl:value-of select="./ADDRESS/@POSTCODE"/></xsl:attribute>
										<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="./@NATUREOFOCCUPANCY_TEXT"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="./@DATEMOVEDIN"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="./@DATEMOVEDOUT"/></xsl:attribute>
										<xsl:attribute name="INTENDEDACTION"><xsl:value-of select="./CURRENTPROPERTY/@INTENDEDACTION_TEXT"/></xsl:attribute>
										<xsl:attribute name="SELLINGPRICE"><xsl:value-of select="./CURRENTPROPERTY/@ESTIMATEDCURRENTVALUE"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<xsl:if test="($Address_Type = '2')">
									<xsl:element name="CORRESPONDENCEADDRESS">
										<xsl:attribute name="ADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
										<xsl:attribute name="POSTCODE"><xsl:value-of select="./ADDRESS/@POSTCODE"/></xsl:attribute>
										<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="./@NATUREOFOCCUPANCY_TEXT"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="./@DATEMOVEDIN"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="./@DATEMOVEDOUT"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<xsl:if test="($Address_Type = '3')">
									<xsl:element name="PREVIOUSADDRESS">
										<xsl:attribute name="ADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
										<xsl:attribute name="POSTCODE"><xsl:value-of select="./ADDRESS/@POSTCODE"/></xsl:attribute>
										<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="./@NATUREOFOCCUPANCY_TEXT"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="./@DATEMOVEDIN"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="./@DATEMOVEDOUT"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
							</xsl:for-each>
							<!-- Customer Address section ends  -->
							<!-- Telephone details section  -->
							<xsl:element name="TELEPHONEDETAILS">
								<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/CUSTOMERTELEPHONENUMBER">
									<xsl:variable name="Usage_Type" select="./@USAGE"/>
									<xsl:variable name="CountryCode" select="./@COUNTRYCODE"/>
									<xsl:variable name="AreaCode" select="./@AREACODE"/>
									<xsl:variable name="TeleohoneNumber" select="./@TELEPHONENUMBER"/>
									<xsl:variable name="ExtensionNumber" select="./@EXTENSIONNUMBER"/>
									<xsl:variable name="PreferredMethodOfContact" select="./@PREFERREDMETHODOFCONTACT"/>
									<xsl:variable name="Number" select="concat('{', msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($TeleohoneNumber), string($ExtensionNumber)),'}')"/>
									<xsl:attribute name="TYPE"><xsl:value-of select="$PreferredMethodOfContact"/></xsl:attribute>
									<xsl:if test="($Usage_Type = '1')">
										<xsl:element name="HOMETELEPHONE">
											<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
									<xsl:if test="($Usage_Type = '2')">
										<xsl:element name="BUSINESSTELEPHONE">
											<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
									<xsl:if test="($Usage_Type = '3')">
										<xsl:element name="MOBILETELEPHONE">
											<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
									<xsl:if test="($PreferredMethodOfContact = 'Y')">
										<xsl:element name="PREFERREDNUMBERTOCONTACT">
											<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
								</xsl:for-each>
							</xsl:element>
						</xsl:element>
					</xsl:element>
					<!-- Telephone details section ends -->
					<!-- Employment section -->
					<xsl:element name="EMPLOYMENTHISTORY">
						<xsl:attribute name="HEADING"><xsl:value-of select="$Heading"/></xsl:attribute>
						<xsl:element name="EMPLOYMENTSUMMARY">
							<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="concat(string(./CUSTOMER/CUSTOMERVERSION/@FIRSTFORENAME), ' ' , string(./CUSTOMER/CUSTOMERVERSION/@SURNAME))"></xsl:value-of></xsl:attribute>
							<xsl:attribute name="TAXDISTRICTANDREFERENCENUMBER"><xsl:value-of select="concat(string(./CUSTOMER/CUSTOMERVERSION/INCOMESUMMARY/@TAXOFFICE),' ',string(./CUSTOMER/CUSTOMERVERSION/INCOMESUMMARY/@TAXREFERENCENUMBER))"/></xsl:attribute>
							<xsl:attribute name="NATIONALINSURANCENUMBER"><xsl:value-of select="./CUSTOMER/CUSTOMERVERSION/@NATIONALINSURANCENUMBER"/></xsl:attribute>
							<xsl:attribute name="UKTAXPAYER"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/INCOMESUMMARY/@UKTAXPAYERINDICATOR))"/></xsl:attribute>
							<xsl:choose>
								<xsl:when test="$CustomerType='1'">
									<xsl:attribute name="UKTAXPAYERDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='700']/@MEMOENTRY"/></xsl:attribute>
								</xsl:when>
								<xsl:when test="$CustomerType='2'">
									<xsl:attribute name="UKTAXPAYERDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='701']/@MEMOENTRY"/></xsl:attribute>
								</xsl:when>
								<xsl:when test="$CustomerType='3'">
									<xsl:attribute name="UKTAXPAYERDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='702']/@MEMOENTRY"/></xsl:attribute>
								</xsl:when>
							</xsl:choose>
							<xsl:attribute name="SELFASSESS"><xsl:value-of select="msg:HandleYesNo(string(./CUSTOMER/CUSTOMERVERSION/INCOMESUMMARY/@SELFASSESS))"/></xsl:attribute>
						</xsl:element>
						<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/EMPLOYMENT">
							<xsl:variable name="MainStatus" select="./@MAINSTATUS"/>
							<xsl:variable name="Employment_Status" select="./@EMPLOYMENTSTATUS"/>
							<xsl:element name="EMPLOYMENT">
								<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="./@EMPLOYMENTSTATUS_TEXT"/></xsl:attribute>
								<xsl:attribute name="JOBTITLE"><xsl:value-of select="./@JOBTITLE"/></xsl:attribute>
								<xsl:attribute name="DATESTARTEDORESTABLISHED"><xsl:value-of select="./@DATESTARTEDORESTABLISHED"/></xsl:attribute>
								<xsl:attribute name="EMPLOYMENTTYPE"><xsl:value-of select="./@EMPLOYMENTTYPE_TEXT"/></xsl:attribute>
								<xsl:attribute name="NATUREOFBUSINESS"><xsl:value-of select="./@INDUSTRYTYPE_TEXT"/></xsl:attribute>
								<xsl:if test="./THIRDPARTY">
									<xsl:call-template name="CONTACTDETAILS">
										<xsl:with-param name="CONTACTDETAIL" select="./THIRDPARTY"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="./NAMEANDADDRESSDIRECTORY">
									<xsl:call-template name="CONTACTDETAILS">
										<xsl:with-param name="CONTACTDETAIL" select="./NAMEANDADDRESSDIRECTORY"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="($MainStatus = 'Y')">
									<xsl:element name="CURRENTEMPLOYMENT"/>
								</xsl:if>
								<xsl:if test="($MainStatus = 'N')">
									<xsl:element name="PREVEMPLOYMENT">
										<xsl:attribute name="DATEOFEMPLOYMENTFINISHED"><xsl:value-of select="./@DATELEFTORCEASEDTRADING"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<xsl:if test="($Employment_Status = '10')">
									<xsl:element name="EMPLOYED">
										<xsl:attribute name="PROBATIONARYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string(./EMPLOYEDDETAILS/@PROBATIONARYINDICATOR))"/></xsl:attribute>
										<xsl:choose>
											<xsl:when test="$CustomerType='1'">
												<xsl:attribute name="PROBATIONARYINDICATORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='710']/@MEMOENTRY"/></xsl:attribute>
												<xsl:attribute name="EMPLOYMENTRELATIONSHIPINDDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='720']/@MEMOENTRY"/></xsl:attribute>
											</xsl:when>
											<xsl:when test="$CustomerType='2'">
												<xsl:attribute name="PROBATIONARYINDICATORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='711']/@MEMOENTRY"/></xsl:attribute>
												<xsl:attribute name="EMPLOYMENTRELATIONSHIPINDDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='721']/@MEMOENTRY"/></xsl:attribute>
											</xsl:when>
											<xsl:when test="$CustomerType='3'">
												<xsl:attribute name="PROBATIONARYINDICATORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='712']/@MEMOENTRY"/></xsl:attribute>
												<xsl:attribute name="EMPLOYMENTRELATIONSHIPINDDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='722']/@MEMOENTRY"/></xsl:attribute>
											</xsl:when>
										</xsl:choose>
										<xsl:attribute name="NOTICEPROBLEMINDICATOR"><xsl:value-of select="msg:HandleYesNo(string(./EMPLOYEDDETAILS/@NOTICEPROBLEMINDICATOR))"/></xsl:attribute>
										<xsl:attribute name="PAYROLLNUMBER"><xsl:value-of select="./EMPLOYEDDETAILS/@PAYROLLNUMBER"/></xsl:attribute>
										<xsl:attribute name="PERCENTSHARESHELD"><xsl:value-of select="./EMPLOYEDDETAILS/@PERCENTSHARESHELD"/></xsl:attribute>
										<xsl:attribute name="EMPLOYMENTRELATIONSHIPIND"><xsl:value-of select="msg:HandleYesNo(string(./EMPLOYEDDETAILS/@EMPLOYMENTRELATIONSHIPIND))"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<!-- Self Employed details-->
								<xsl:if test="($Employment_Status = '20')">
									<xsl:element name="SELFEMPLOYED">
										<xsl:attribute name="DATEACQUIREINTEREST"><xsl:value-of select="./SELFEMPLOYEDDETAILS/@DATEFINANCIALINTERESTHELD"/></xsl:attribute>
										<xsl:attribute name="REGISTRATIONNUMBER"><xsl:value-of select="./SELFEMPLOYEDDETAILS/@REGISTRATIONNUMBER"/></xsl:attribute>
										<xsl:attribute name="VATNUMBER"><xsl:value-of select="./SELFEMPLOYEDDETAILS/@VATNUMBER"/></xsl:attribute>
										<xsl:attribute name="PERCENTSHARESHELD"><xsl:value-of select="./SELFEMPLOYEDDETAILS/@PERCENTSHARESHELD"/></xsl:attribute>
										<xsl:attribute name="COMPANYSTATUSTYPE"><xsl:value-of select="./SELFEMPLOYEDDETAILS/@COMPANYSTATUSTYPE_TEXT"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<!-- Contract Employee details -->
								<xsl:if test="($Employment_Status = '30')">
									<xsl:element name="CONTRACTWORKERS">
										<xsl:attribute name="LENGTHOFCONTRACTYEARS"><xsl:value-of select="./CONTRACTDETAILS/@LENGTHOFCONTRACTYEARS"/></xsl:attribute>
										<xsl:attribute name="LENGTHOFCONTRACTMONTHS"><xsl:value-of select="./CONTRACTDETAILS/@LENGTHOFCONTRACTMONTHS"/></xsl:attribute>
										<xsl:attribute name="CONTRACTDETAILSSTARTDATE"><xsl:value-of select="./@DATESTARTEDORESTABLISHED"/></xsl:attribute>
										<xsl:attribute name="CONTRACTDETAILSFINISHDATE"><xsl:value-of select="./@DATELEFTORCEASEDTRADING"/></xsl:attribute>
										<xsl:attribute name="LIKELYTOBERENEWED"><xsl:value-of select="msg:HandleYesNo(string(./CONTRACTDETAILS/@LIKELYTOBERENEWED))"/></xsl:attribute>
										<xsl:attribute name="NUMBEROFRENEWALS"><xsl:value-of select="./CONTRACTDETAILS/@NUMBEROFRENEWALS"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<xsl:apply-templates select="ACCOUNTANT" />
								<!-- Income Details -->
								<xsl:if test="contains(@EMPLOYMENTSTATUS_TYPES,'SELF') or contains(@EMPLOYMENTSTATUS_TYPES,'CON')">
									<xsl:element name="INCOMEDETAILS">
										<xsl:element name="EMPLOYED">
											<xsl:attribute name="INCOMETYPE">
												<xsl:text>Net Profit</xsl:text>
											</xsl:attribute>
											<xsl:attribute name="INCOMEAMOUNT">
												<xsl:value-of select="NETPROFIT/@YEAR1AMOUNT" />
											</xsl:attribute>										
										</xsl:element>
									</xsl:element>
								</xsl:if>								
								<xsl:if test="./EARNEDINCOME">
									<xsl:element name="INCOMEDETAILS">
										<xsl:attribute name="ALLOWABLEANNUALINCOME">											
											<xsl:value-of select="../INCOMESUMMARY/@ALLOWABLEANNUALINCOME"/>
										</xsl:attribute>																		
										<xsl:attribute name="HEADING"><xsl:value-of select="$Heading"/></xsl:attribute>
										<!--										<xsl:if test="($Employment_Status = '10')"> -->
										<xsl:for-each select="./EARNEDINCOME">
											<xsl:element name="EMPLOYED">
												<xsl:attribute name="INCOMETYPE"><xsl:value-of select="./@EARNEDINCOMETYPE_TEXT"/></xsl:attribute>
												<xsl:attribute name="INCOMEAMOUNT"><xsl:value-of select="./@EARNEDINCOMEAMOUNT * ./@PAYMENTFREQUENCYTYPE"/></xsl:attribute>
												<xsl:if test="./@EARNEDINCOMETYPE='65'">
													<xsl:choose>
														<xsl:when test="$CustomerType='1' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='810']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='810']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
														<xsl:when test="$CustomerType='2' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='811']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='811']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
														<xsl:when test="$CustomerType='3' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='812']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='812']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
													</xsl:choose>
												</xsl:if>
												<xsl:if test="./@EARNEDINCOMETYPE='99'">
													<xsl:choose>
														<xsl:when test="$CustomerType='1' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='820']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='820']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
														<xsl:when test="$CustomerType='2' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='821']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='821']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
														<xsl:when test="$CustomerType='3' and /RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='822']/@MEMOENTRY">
															<xsl:element name="DETAILS">
																<xsl:attribute name="DETAILSTEXT"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='822']/@MEMOENTRY"/></xsl:attribute>
															</xsl:element>
														</xsl:when>
													</xsl:choose>
												</xsl:if>
											</xsl:element>
										</xsl:for-each>
										<!-- </xsl:if> -->
									</xsl:element>
								</xsl:if>
							</xsl:element>
						</xsl:for-each>
					</xsl:element>
					<!-- **************************************** Landlord section moved to Customer address node - moved back below ***************************************** -->
					<xsl:if test="./CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[@TOTALCOLLATERALBALANCE > 0]">
						<xsl:element name="CURRENTRESIDENTIALLENDER">
							<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[@TOTALCOLLATERALBALANCE > 0]">
							<xsl:attribute name="CUSTOMERHEADING"><xsl:value-of select="$Heading"/></xsl:attribute>
							<xsl:variable name="TOTALCOLLATERALBALANCE" select="./@TOTALCOLLATERALBALANCE"/>
							<xsl:if test="../../../CUSTOMERADDRESS[@ADDRESSTYPE = '1']">
								<xsl:element name="LENDERDETAIL">
									<xsl:variable name="FlatNumber" select="../THIRDPARTY/ADDRESS/@FLATNUMBER"/>
									<xsl:variable name="HouseName" select="../THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME"/>
									<xsl:variable name="HouseNumber" select="../THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER"/>
									<xsl:variable name="District" select="../THIRDPARTY/ADDRESS/@DISTRICT"/>
									<xsl:variable name="Street" select="../THIRDPARTY/ADDRESS/@STREET"/>
									<xsl:variable name="Town" select="../THIRDPARTY/ADDRESS/@TOWN"/>
									<xsl:variable name="County" select="../THIRDPARTY/ADDRESS/@COUNTY"/>
									<xsl:variable name="Country" select="../THIRDPARTY/ADDRESS/@COUNTRY_TEXT"/>
									<xsl:variable name="Address" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
									<xsl:attribute name="LENDERNAME"><xsl:value-of select="../THIRDPARTY/@COMPANYNAME"/></xsl:attribute>
									<xsl:attribute name="LENDERADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
									<xsl:attribute name="LENDERPOSTCODE"><xsl:value-of select="../THIRDPARTY/ADDRESS/@POSTCODE"/></xsl:attribute>
									<xsl:for-each select="../THIRDPARTY/CONTACTDETAILS/CONTACTTELEPHONEDETAILS">
										<xsl:variable name="Telephone" select="@TELENUMBER"/>
										<xsl:variable name="Extension" select="@EXTENSIONNUMBER"/>
										<xsl:variable name="CountryCode" select="@CountryCode"/>
										<xsl:variable name="AreaCode" select="@AreaCode"/>
										<xsl:if test="@USAGE_TYPES= 'W'">
											<xsl:attribute name="LENDERPHONE"><xsl:value-of select="msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($Telephone), string($Extension))"/></xsl:attribute>
										</xsl:if>
										<xsl:if test="@USAGE_TYPES= 'F'">
											<xsl:attribute name="LENDERFAX"><xsl:value-of select="msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($Telephone), string($Extension))"/></xsl:attribute>
										</xsl:if>
									</xsl:for-each>
									<xsl:attribute name="ACCOUNTNUMBER">{<xsl:value-of select="../../ACCOUNT/@ACCOUNTNUMBER"/>}</xsl:attribute>
									<xsl:attribute name="HOUSENUMBER"><xsl:value-of select="../../../CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:attribute>
									<xsl:attribute name="HOUSEPOSTCODE"><xsl:value-of select="../../../CUSTOMERADDRESS/ADDRESS/@POSTCODE"/></xsl:attribute>
									<xsl:attribute name="MONTHLYPAYMENT"><xsl:value-of select="@TOTALMONTHLYCOST"/></xsl:attribute>
									<xsl:attribute name="BALANCEOUTSTANDING"><xsl:value-of select="@TOTALCOLLATERALBALANCE"/></xsl:attribute>
									<xsl:attribute name="MORTGAGESTARTDATE"><xsl:value-of select="./MORTGAGELOAN/@STARTDATE"/></xsl:attribute>
									<xsl:attribute name="MONTHSCURRENTLYINARREAR"><xsl:value-of select="//CUSTOMER/CUSTOMERVERSION/CUSTOMERARREARSHISTORY/@MAXIMUMNUMBEROFMONTHS"/></xsl:attribute>
									<xsl:attribute name="HOUSINGBENEFIT"><xsl:value-of select="msg:HandleYesNo(string(//CUSTOMER/CUSTOMERVERSION/@HOUSINGBENEFIT))"/></xsl:attribute>
									<xsl:attribute name="ORIGINALLOANAMOUNT"><xsl:value-of select="./MORTGAGELOAN/@ORIGINALLOANAMOUNT"/></xsl:attribute>
									<xsl:attribute name="ESTIMATEDPROPERTYVALUE"><xsl:value-of select="../../../CUSTOMERADDRESS/CURRENTPROPERTY/@ESTIMATEDCURRENTVALUE"/></xsl:attribute>
									<xsl:attribute name="REDEMPTIONSTATUS"><xsl:value-of select="COMBOVALUE/@VALUENAME"/></xsl:attribute>
									<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="@MONTHLYRENTALINCOME"/></xsl:attribute>
									<xsl:if test="./COMBOVALIDATION/@VALIDATIONTYPE = 'N'">
										<xsl:attribute name="OUTSTANDINGMORTGAGEBALANCE"><xsl:value-of select="sum(MORTGAGELOAN[COMBOVALIDATION/@VALIDATIONTYPE='N']/@OUTSTANDINGBALANCE)"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="OTHERPROPERYMORTGAGE"><xsl:value-of select="msg:HandleYesNo(string(@OTHERPROPERTYMORTGAGE))"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
						</xsl:for-each>
					</xsl:element>
					</xsl:if>
					<!-- Current Landlord Details -->
					<xsl:element name="CURRENTLANDLORD">
						<xsl:for-each select="./CUSTOMER/CUSTOMERVERSION/CUSTOMERADDRESS">
							<xsl:if test="./TENANCY">
								<xsl:variable name="FlatNumber" select="./ADDRESS/@FLATNUMBER"/>
								<xsl:variable name="HouseName" select="./ADDRESS/@BUILDINGORHOUSENAME"/>
								<xsl:variable name="HouseNumber" select="./ADDRESS/@BUILDINGORHOUSENUMBER"/>
								<xsl:variable name="District" select="./ADDRESS/@DISTRICT"/>
								<xsl:variable name="Street" select="./ADDRESS/@STREET"/>
								<xsl:variable name="Town" select="./ADDRESS/@TOWN"/>
								<xsl:variable name="County" select="./ADDRESS/@COUNTY"/>
								<xsl:variable name="Country" select="./ADDRESS/@COUNTRY_TEXT"/>
								<xsl:variable name="Address" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
								<xsl:element name="ADDRESSES">
									<xsl:element name="TENANCY">
										<xsl:attribute name="CUSTOMERHEADING"><xsl:value-of select="$Heading"/></xsl:attribute>
										<xsl:attribute name="CUSTOMERNAME"><xsl:value-of select="concat(string(../@FIRSTFORENAME), ' ' , string(../@SURNAME))"></xsl:value-of></xsl:attribute>
										<xsl:attribute name="RESIDENTIALADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
										<xsl:attribute name="ACCOUNTNUMBER">{<xsl:value-of select="./TENANCY/@ACCOUNTNUMBER"/>}</xsl:attribute>
										<xsl:attribute name="MONTHLYRENTALAMOUNT"><xsl:value-of select="./TENANCY/@MONTHLYRENTAMOUNT"/></xsl:attribute>
										<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="./@DATEMOVEDIN"/></xsl:attribute>
										<xsl:attribute name="MONTHSINARREARS"><xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/CUSTOMERARREARSHISTORY/COMBOVALIDATION/@VALIDATIONTYPE = 'O'"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/CUSTOMERARREARSHISTORY/@MAXIMUMNUMBEROFMONTHS"/></xsl:if></xsl:attribute>
										<xsl:if test="./TENANCY/THIRDPARTY">
											<xsl:call-template name="CONTACTDETAILS">
												<xsl:with-param name="CONTACTDETAIL" select="./TENANCY/THIRDPARTY"/>
											</xsl:call-template>
										</xsl:if>
										<xsl:if test="./TENANCY/NAMEANDADDRESSDIRECTORY">
											<xsl:call-template name="CONTACTDETAILS">
												<xsl:with-param name="CONTACTDETAIL" select="./TENANCY/NAMEANDADDRESSDIRECTORY"/>
											</xsl:call-template>
										</xsl:if>
									</xsl:element>
								</xsl:element>
							</xsl:if>
						</xsl:for-each>
					</xsl:element>
				</xsl:for-each>
				<!-- Current Landlord Details  end -->
				<!-- Payment History section -->
				<xsl:element name="PAYMENTHISTORY">
					<xsl:attribute name="BANKRUPTCYHISTORYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/FINANCIALSUMMARY/@BANKRUPTCYHISTORYINDICATOR))"/></xsl:attribute>
					<xsl:attribute name="CCJHISTORYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/FINANCIALSUMMARY/@CCJHISTORYINDICATOR))"/></xsl:attribute>
					<xsl:attribute name="DECLINEDMORTGAGEINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/FINANCIALSUMMARY/@DECLINEDMORTGAGEINDICATOR))"/></xsl:attribute>
					<xsl:attribute name="ARREARSHISTORYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/FINANCIALSUMMARY/@ARREARSHISTORYINDICATOR))"/></xsl:attribute>
					<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/BANKRUPTCYHISTORYLIST">
						<xsl:element name="BANKRUPTCYHISTORY">
							<xsl:attribute name="DATEOFDISCHARGE"><xsl:value-of select="BANKRUPTCYHISTORY/@DATEOFDISCHARGE"/></xsl:attribute>
							<xsl:attribute name="AMOUNTOFDEBT"><xsl:value-of select="BANKRUPTCYHISTORY/@AMOUNTOFDEBT"/></xsl:attribute>
							<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="BANKRUPTCYHISTORY/@MONTHLYREPAYMENT"/></xsl:attribute>
							<xsl:attribute name="IVA"><xsl:value-of select="msg:HandleYesNo(string(./BANKRUPTCYHISTORY/@IVA))"/></xsl:attribute>
							<xsl:attribute name="OTHERDETAILS"><xsl:value-of select="./BANKRUPTCYHISTORY/@OTHERDETAILS"/></xsl:attribute>
							<xsl:for-each select="./BANKRUPTCYHISTORY/CUSTOMERVERSIONBANKRUPTCYHISTORY/CUSTOMERROLE">
								<xsl:element name="CUSTOMERS">
									<xsl:attribute name="APPLICANT"><xsl:value-of select="msg:GetHeading(string(./@CUSTOMERORDER),string(./@CUSTOMERROLETYPE))"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
				<!-- Payment history ends -->
				<!-- commitments -->
				<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/COMMITMENTS">
					<xsl:element name="COMMITMENT">
						<xsl:attribute name="AGREEMENTTYPE"><xsl:value-of select="./@AGREEMENTTYPE_TEXT"/></xsl:attribute>
						<xsl:if test="not(contains(@AGREEMENTTYPE_TYPES,'CC'))">
							<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="./@MONTHLYREPAYMENT"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="TOTALOUTSTANDINGBALANCE"><xsl:value-of select="./@TOTALOUTSTANDINGBALANCE"/></xsl:attribute>
						<xsl:attribute name="ENDDATE"><xsl:value-of select="./@ENDDATE"/></xsl:attribute>
						<xsl:attribute name="PAIDFORBYBUSINESS"><xsl:value-of select="msg:HandleYesNo(string(./@PAIDFORBYBUSINESS))"/></xsl:attribute>
						<xsl:attribute name="TOBEREPAIDINDICATOR"><xsl:value-of select="msg:HandleYesNo(string(./@TOBEREPAIDINDICATOR))"/></xsl:attribute>
						<xsl:attribute name="ACCOUNTNUMBER">{<xsl:value-of select="./@ACCOUNTNUMBER"/>}</xsl:attribute>
						<xsl:attribute name="ARREARINMONTHS"><xsl:value-of select="./ARREARSHISTORY/@MAXIMUMNUMBEROFMONTHS"/></xsl:attribute>
						<xsl:if test="./THIRDPARTY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="./THIRDPARTY"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="./NAMEANDADDRESSDIRECTORY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="./NAMEANDADDRESSDIRECTORY"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:for-each select="./ACCOUNTRELATIONSHIP/CUSTOMERROLE">
							<xsl:element name="CUSTOMERS">
								<xsl:attribute name="APPLICANT"><xsl:value-of select="msg:GetHeading(string(./@CUSTOMERORDER),string(./@CUSTOMERROLETYPE))"/></xsl:attribute>
							</xsl:element>
						</xsl:for-each>
					</xsl:element>
				</xsl:for-each>
				<!-- commitments end -->
				<xsl:choose>
					<xsl:when test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION != '20' and $RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION != '90'">
						<xsl:element name="PURCHASEINFORMATION">
							<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED"/></xsl:attribute>
							<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE"/></xsl:attribute>
							<xsl:attribute name="PRIVATESALEINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@FAMILYSALEINDICATOR))"/></xsl:attribute>
							<xsl:attribute name="BUSINESSCONNECTION"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@BUSINESSCONNECTION))"/></xsl:attribute>
							<xsl:attribute name="FULLMARKETVALUE"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@FULLMARKETVALUE))"/></xsl:attribute>
							<xsl:attribute name="FULLMARKETVALUEDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1200']/@MEMOENTRY"/></xsl:attribute>
							<xsl:attribute name="DATEOFENTRY"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@DATEOFENTRY"/></xsl:attribute>
							<xsl:attribute name="DIRECTFINANCIALBENEFITIND"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWLOAN/@DIRECTFINANCIALBENEFITIND))"/></xsl:attribute>																
						</xsl:element>
					</xsl:when>
					<xsl:otherwise>
						<xsl:element name="REMORTGAGESDETAILS">
							<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE"/></xsl:attribute>
								<xsl:attribute name="FIRSTTOTALCOLLATERALBALANCE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@REMORTGAGEINDICATOR='1') and (@SECONDCHARGEINDICATOR='0')]/@TOTALCOLLATERALBALANCE"/></xsl:attribute>
								<xsl:attribute name="SECONDTOTALCOLLATERALBALANCE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@REMORTGAGEINDICATOR='1') and (@SECONDCHARGEINDICATOR='1')]/@TOTALCOLLATERALBALANCE"/></xsl:attribute>								
								<xsl:attribute name="DIRECTFINANCIALBENEFITIND"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWLOAN/@DIRECTFINANCIALBENEFITIND))"/></xsl:attribute>																
						</xsl:element>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTYDEPOSIT">
					<xsl:element name="SOURCEANDDEPOSIT">
						<xsl:attribute name="DEPOSITTYPE"><xsl:value-of select="./@SOURCEOFFUNDING_TEXT"/></xsl:attribute>
						<xsl:attribute name="DEPOSITAMOUNT"><xsl:value-of select="./@AMOUNT"/></xsl:attribute>
						<xsl:attribute name="SOURCEREASON"><xsl:value-of select="./@SOURCEREASON"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
				<xsl:element name="FINANCIALINCENTIVES">
					<xsl:attribute name="FINANCIALINCENTIVES"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@FINANCIALINCENTIVES))"/></xsl:attribute>
					<xsl:attribute name="FINANCIALINCENTIVESDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1210']/@MEMOENTRY"/></xsl:attribute>
				</xsl:element>
				<!-- New Property details -->
				<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY">
					<xsl:element name="MORTGAGEDPROPERTYDETAILS">
						<xsl:variable name="FlatNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER"/>
						<xsl:variable name="HouseName" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENAME"/>
						<xsl:variable name="HouseNumber" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER"/>
						<xsl:variable name="District" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@DISTRICT"/>
						<xsl:variable name="Street" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@STREET"/>
						<xsl:variable name="Town" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS//@TOWN"/>
						<xsl:variable name="County" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS//@COUNTY"/>
						<xsl:variable name="Country" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS//@COUNTRY_TEXT"/>
						<xsl:variable name="Address" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
						<xsl:attribute name="PROPERTYADDRESS"><xsl:value-of select="$Address"/></xsl:attribute>
						<xsl:attribute name="POSTCODE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@POSTCODE"/></xsl:attribute>
						<xsl:attribute name="TENURETYPE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@TENURETYPE_TEXT"/></xsl:attribute>
						<xsl:attribute name="UNEXPIREDTERMOFLEASEYEARS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYLEASEHOLD/@UNEXPIREDTERMOFLEASEYEARS"/></xsl:attribute>
						<xsl:attribute name="HOUSEBUILDERSGUARANTEE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@HOUSEBUILDERSGUARANTEE_TEXT"/></xsl:attribute>
						<xsl:attribute name="DESCRIPTIONOFPROPERTY"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@DESCRIPTIONOFPROPERTY_TEXT"/></xsl:attribute>
						<xsl:attribute name="YEARBUILT"><xsl:value-of select="concat('{',string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@YEARBUILT),'}')"/></xsl:attribute>
						<xsl:attribute name="NEWPROPERTYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@NEWPROPERTYINDICATOR))"/></xsl:attribute>
						<xsl:attribute name="TYPEOFPROPERTY"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@TYPEOFPROPERTY_TEXT"/></xsl:attribute>
						<xsl:attribute name="NUMBEROFSTOREYS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@NUMBEROFSTOREYS"/></xsl:attribute>
						<xsl:attribute name="NOFLATSINBLOCK"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@NOFLATSINBLOCK"/></xsl:attribute>
						<xsl:attribute name="FLOORNUMBER"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@FLOORNUMBER"/></xsl:attribute>
						<xsl:attribute name="FLATABOVECOMMERCIAL"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@FLATABOVECOMMERCIAL))"/></xsl:attribute>
						<xsl:attribute name="FLATABOVECOMMERCIALDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1240']/@MEMOENTRY"/></xsl:attribute>
						<!-- ROOMS -->
						<xsl:attribute name="RECEPTION"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYROOMTYPE[@ROOMTYPE='1']/@NUMBEROFROOMS"/></xsl:attribute>
						<xsl:attribute name="WCS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYROOMTYPE[@ROOMTYPE='4']/@NUMBEROFROOMS"/></xsl:attribute>
						<xsl:attribute name="BEDROOMS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYROOMTYPE[@ROOMTYPE='2']/@NUMBEROFROOMS"/></xsl:attribute>
						<xsl:attribute name="GARAGES"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@NUMBEROFGARAGES"/></xsl:attribute>
						<xsl:attribute name="KITCHENS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYROOMTYPE[@ROOMTYPE='5']/@NUMBEROFROOMS"/></xsl:attribute>
						<xsl:attribute name="OUTBUILDINGS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@NUMBEROFOUTBUILDINGS"/></xsl:attribute>
						<xsl:attribute name="BATHROOMS"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYROOMTYPE[@ROOMTYPE='3']/@NUMBEROFROOMS"/></xsl:attribute>
						<!-- ROOMS ENDS -->
						<xsl:attribute name="BUILDINGCONSTRUCTIONTYPE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@BUILDINGCONSTRUCTIONTYPE_TEXT"/></xsl:attribute>
						<xsl:attribute name="ROOFCONSTRUCTIONTYPE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@ROOFCONSTRUCTIONTYPE_TEXT"/></xsl:attribute>
						<xsl:attribute name="CONSTRUCTIONTYPEDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1250']/@MEMOENTRY"/></xsl:attribute>
						<xsl:attribute name="TOTALLANDTYPE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@TOTALLANDTYPE_TEXT"/></xsl:attribute>
						<xsl:attribute name="AGRICULTURALRESTRICTIONS"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@AGRICULTURALRESTRICTIONS))"/></xsl:attribute>
						<xsl:attribute name="AGRICULTURALRESTRICTIONSDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1260']/@MEMOENTRY"/></xsl:attribute>
						<xsl:attribute name="ENTIRELYOWNUSEINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@MAINRESIDENCEIND))"/></xsl:attribute>
						<xsl:attribute name="RESIDENTIALUSEONLYINDICATOR">
							<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@RESIDENTIALUSEONLYINDICATOR = '1'">
								<xsl:value-of select="msg:HandleYesNo('0')"/>
							</xsl:if>
							<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@RESIDENTIALUSEONLYINDICATOR = '0'">
								<xsl:value-of select="msg:HandleYesNo('1')"/>
							</xsl:if>
						</xsl:attribute>
						<xsl:attribute name="RESIDENTIALUSEONLYINDICATORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1270']/@MEMOENTRY"/></xsl:attribute>
						<xsl:attribute name="EXLOCALAUTHORITYINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@EXLOCALAUTHORITYINDICATOR))"/></xsl:attribute>
						<xsl:attribute name="EXLOCALAUTHORITYINDICATORDATE"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1295']/@MEMOENTRY"/></xsl:attribute>
						<xsl:attribute name="LETORPARTLETINDICATOR"><xsl:value-of select="msg:HandleYesNo(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@LETORPARTLETINDICATOR))"/></xsl:attribute>
						<xsl:attribute name="LETORPARTLETINDICATORDETAILS"><xsl:value-of select="/RESPONSE/APPLICATIONFORM/APPLICATION/MEMOPAD[@ENTRYTYPE='1280']/@MEMOENTRY"/></xsl:attribute>
						<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@MONTHLYRENTALINCOME"/></xsl:attribute>
						<xsl:attribute name="REGULATIONINDICATORDETAILS">
							<xsl:choose>
								<xsl:when test="contains(//APPLICATIONFACTFIND/@REGULATIONINDICATOR_TYPES,'R')">
									<xsl:text>Yes</xsl:text>
								</xsl:when>
								<xsl:otherwise>
									<xsl:text>No</xsl:text>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
						<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/OTHERRESIDENT">
							<xsl:element name="OTHERRESIDENT">
								<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="./PERSON/@FIRSTFORENAME"/></xsl:attribute>
								<xsl:attribute name="SURNAME"><xsl:value-of select="./PERSON/@SURNAME"/></xsl:attribute>
								<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="./PERSON/@DATEOFBIRTH"/></xsl:attribute>
								<xsl:attribute name="RELATIONSHIPTOAPPLICANT"><xsl:value-of select="./PERSON/@RELATIONSHIPTOAPPLICANT_TEXT"/></xsl:attribute>
							</xsl:element>
						</xsl:for-each>
						<xsl:element name="ACCESSTOTHEPROPERTY">
							<xsl:attribute name="VALUATIONTYPE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/@VALUATIONTYPE_TEXT"/></xsl:attribute>
							<xsl:attribute name="ACCESSCONTACTNAME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/@ARRANGEMENTSFORACCESS_TEXT"/></xsl:attribute>
							<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS">
								<xsl:for-each select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS">
									<xsl:attribute name="COMPANYNAME"><xsl:value-of select="@ACCESSCONTACTNAME"/></xsl:attribute>
									<xsl:variable name="CountryCode" select="@COUNTRYCODE"/>
									<xsl:variable name="AreaCode" select="@AREACODE"/>
									<xsl:variable name="TeleohoneNumber" select="@ACCESSTELEPHONENUMBER"/>
									<xsl:variable name="ExtensionNumber" select="@EXTENSIONNUMBER"/>
									<xsl:attribute name="NUMBER"><xsl:value-of select="concat('{',msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($TeleohoneNumber), string($ExtensionNumber)),'}')"/></xsl:attribute>
								</xsl:for-each>
							</xsl:if>
							<!--<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONESTATEAGENT/THIRDPARTY">
								<xsl:call-template name="CONTACTDETAILS">
									<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONESTATEAGENT/THIRDPARTY"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONESTATEAGENT/NAMEANDADDRESSDIRECTORY">
								<xsl:call-template name="CONTACTDETAILS">
									<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONESTATEAGENT/NAMEANDADDRESSDIRECTORY"/>
								</xsl:call-template>
							</xsl:if>-->
							<xsl:element name="VENDOR">
								<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYVENDOR/THIRDPARTY">
									<xsl:call-template name="CONTACTDETAILS">
										<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYVENDOR/THIRDPARTY"/>
									</xsl:call-template>
								</xsl:if>
								<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYVENDOR/THIRDPARTY/NAMEANDADDRESSDIRECTORY">
									<xsl:call-template name="CONTACTDETAILS">
										<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYVENDOR/THIRDPARTY/NAMEANDADDRESSDIRECTORY"/>
									</xsl:call-template>
								</xsl:if>
							</xsl:element>
						</xsl:element>
					</xsl:element>
				</xsl:if>
				<!-- Solicitor details -->
				<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP">
					<xsl:element name="SOLICITORDETAILS">
						<xsl:attribute name="SOLICITORFIRSTNAME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME"/></xsl:attribute>
						<xsl:attribute name="SOLICITORSURNAME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME"/></xsl:attribute>
						<xsl:attribute name="DXNUMBER">{<xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/THIRDPARTY/@DXID"/>}</xsl:attribute>
						<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/THIRDPARTY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/THIRDPARTY"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY"/>
							</xsl:call-template>
						</xsl:if>
					</xsl:element>
				</xsl:if>
				<!-- Income from existing property-->
				<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@ORIGINALNATUREOFLOAN='10' or @ORIGINALNATUREOFLOAN='11') and @REDEMPTIONSTATUS !='40']">
					<xsl:element name="INCOMEFROMEXISTINGPROPERTY">
						<xsl:attribute name="INVESTMENTPROPOWNED"><xsl:value-of select="count($RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@ORIGINALNATUREOFLOAN='10' or @ORIGINALNATUREOFLOAN='11') and @REDEMPTIONSTATUS !='40'])"/></xsl:attribute>
						<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="sum($RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@ORIGINALNATUREOFLOAN='10' or @ORIGINALNATUREOFLOAN='11') and @REDEMPTIONSTATUS !='40']/@MONTHLYRENTALINCOME)"/></xsl:attribute>
						<xsl:attribute name="MONTHLYMORTGAGEPAYMENTS"><xsl:value-of select="sum($RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/MORTGAGEACCOUNT[(@ORIGINALNATUREOFLOAN='10' or @ORIGINALNATUREOFLOAN='11') and @REDEMPTIONSTATUS !='40']/@TOTALMONTHLYCOST)"/></xsl:attribute>
					</xsl:element>
				</xsl:if>
				<!-- Bank building society details -->
				<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC">
					<xsl:element name="APPLICATIONBANKBUILDINGSOC">
						<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/THIRDPARTY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/THIRDPARTY"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/NAMEANDADDRESSDIRECTORY">
							<xsl:call-template name="CONTACTDETAILS">
								<xsl:with-param name="CONTACTDETAIL" select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/NAMEANDADDRESSDIRECTORY"/>
							</xsl:call-template>
						</xsl:if>
						<xsl:attribute name="ACCOUNTNAME"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/@ACCOUNTNAME"/></xsl:attribute>
						<xsl:attribute name="BANKSORTCODE"><xsl:value-of select="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/THIRDPARTY/@THIRDPARTYBANKSORTCODE"/></xsl:attribute>
						<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="concat('{',$RESPONSE/APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
					</xsl:element>
				</xsl:if>
				<!-- Fee Payment details -->
				<xsl:if test="$RESPONSE/APPLICATION/FEEPAYMENT[@PAYMENTEVENT='10']">
					<xsl:element name="FEEPAYMENT">
						<xsl:for-each select="$RESPONSE/APPLICATION/FEEPAYMENT[@PAYMENTEVENT='10']">
							<xsl:element name="PAYMENT">
								<xsl:attribute name="FEETYPE"><xsl:value-of select="@FEETYPE_TEXT"> </xsl:value-of></xsl:attribute>
								<xsl:attribute name="AMOUNTPAID"><xsl:value-of select="@AMOUNTPAID"/></xsl:attribute>
							</xsl:element>
						</xsl:for-each>
					</xsl:element>
				</xsl:if>
				<!-- Additional information -->
				<xsl:element name="ADDITIONALINFORMATION">
					<xsl:attribute name="ADDITIONALINFORMATIONTEXT"><xsl:value-of select="concat($RESPONSE/APPLICATION/MEMOPAD[@ENTRYTYPE='12']/@MEMOENTRY,'\par  ',$RESPONSE/APPLICATION/MEMOPAD[@ENTRYTYPE='1610']/@MEMOENTRY)"/></xsl:attribute>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<!-- Contact details template -->
	<xsl:template name="CONTACTDETAILS">
		<xsl:param name="CONTACTDETAIL"/>
		<xsl:variable name="FlatNumber" select="$CONTACTDETAIL/ADDRESS/@FLATNUMBER"/>
		<xsl:variable name="HouseName" select="$CONTACTDETAIL/ADDRESS/@BUILDINGORHOUSENAME"/>
		<xsl:variable name="HouseNumber" select="$CONTACTDETAIL/ADDRESS/@BUILDINGORHOUSENUMBER"/>
		<xsl:variable name="District" select="$CONTACTDETAIL/ADDRESS/@DISTRICT"/>
		<xsl:variable name="Street" select="$CONTACTDETAIL/ADDRESS/@STREET"/>
		<xsl:variable name="Town" select="$CONTACTDETAIL/ADDRESS/@TOWN"/>
		<xsl:variable name="County" select="$CONTACTDETAIL/ADDRESS/@COUNTY"/>
		<xsl:variable name="Country" select="$CONTACTDETAIL/ADDRESS/@COUNTRY_TEXT"/>
		<xsl:variable name="EmployerAddress" select="msg:HandleAddress(string($FlatNumber), string($HouseName), string($HouseNumber), string($Street), string($District), string($Town), string($County), string($Country))"/>
		<xsl:attribute name="COMPANYNAME"><xsl:value-of select="$CONTACTDETAIL/@COMPANYNAME"/></xsl:attribute>
		<xsl:attribute name="CONTACTFORENAME"><xsl:value-of select="$CONTACTDETAIL/CONTACTDETAILS/@CONTACTFORENAME"/></xsl:attribute>
		<xsl:attribute name="CONTACTSURNAME"><xsl:value-of select="$CONTACTDETAIL/CONTACTDETAILS/@CONTACTSURNAME"/></xsl:attribute>
		<xsl:attribute name="ADDRESS"><xsl:value-of select="$EmployerAddress"/></xsl:attribute>
		<xsl:attribute name="POSTCODE"><xsl:value-of select="$CONTACTDETAIL/ADDRESS/@POSTCODE"/></xsl:attribute>
		<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="$CONTACTDETAIL/CONTACTDETAILS/@EMAILADDRESS"/></xsl:attribute>
		<xsl:attribute name="ORGANISATIONTYPE"><xsl:value-of select="$CONTACTDETAIL/@ORGANISATIONTYPE"/></xsl:attribute>
		<xsl:for-each select="$CONTACTDETAIL/CONTACTDETAILS/CONTACTTELEPHONEDETAILS">
			<xsl:variable name="Usage_Type" select="./@USAGE_TYPES"/>
			<xsl:variable name="CountryCode" select="./@COUNTRYCODE"/>
			<xsl:variable name="AreaCode" select="./@AREACODE"/>
			<xsl:variable name="TeleohoneNumber" select="./@TELENUMBER"/>
			<xsl:variable name="ExtensionNumber" select="./@EXTENSIONNUMBER"/>
			<xsl:variable name="Number" select="concat('{',msg:HandlePhoneNumber(string($CountryCode), string($AreaCode), string($TeleohoneNumber), string($ExtensionNumber)),'}')"/>
			<xsl:if test="($Usage_Type = 'W')">
				<xsl:element name="BUSINESSTELEPHONE">
					<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
			<xsl:if test="($Usage_Type = 'F')">
				<xsl:element name="FAX">
					<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
			<xsl:if test="($Usage_Type = 'H')">
				<xsl:element name="HOME">
					<xsl:attribute name="NUMBER"><xsl:value-of select="$Number"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="ACCOUNTANT">
		<xsl:element name="ACCOUNTANT">
			<xsl:attribute name="QUALIFICATIONS"><xsl:value-of select="/@QUALIFICATIONS_TEXT"/></xsl:attribute>
			<xsl:attribute name="YEARSACTINGFORCUSTOMER"><xsl:value-of select="@YEARSACTINGFORCUSTOMER"/></xsl:attribute>
			<xsl:attribute name="ACCOUNTANTFIRM"><xsl:value-of select="@ACCOUNTANCYFIRMNAME"/></xsl:attribute>
			<xsl:if test="THIRDPARTY">
				<xsl:call-template name="CONTACTDETAILS">
					<xsl:with-param name="CONTACTDETAIL" select="THIRDPARTY"/>
				</xsl:call-template>
			</xsl:if>
			<xsl:if test="NAMEANDADDRESSDIRECTORY">
				<xsl:call-template name="CONTACTDETAILS">
					<xsl:with-param name="CONTACTDETAIL" select="NAMEANDADDRESSDIRECTORY"/>
				</xsl:call-template>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<xsl:template match="LOANCOMPONENT">
		<LOANCOMPONENT>
			<xsl:attribute name="PRODUCTCODE"><xsl:value-of select="concat('{',string(@MORTGAGEPRODUCTCODE),'}')"/></xsl:attribute>
			<xsl:attribute name="INTERESTRATE"><xsl:value-of select="string(@RESOLVEDRATE)"/></xsl:attribute>
			<xsl:attribute name="INTERESTPERIOD"><xsl:value-of select="string(INTERESTRATETYPE/@INTERESTRATEPERIOD) div 12"/></xsl:attribute>
			<xsl:attribute name="SUBPURPOSEOFLOAN"><xsl:value-of select="@SUBPURPOSEOFLOAN"/></xsl:attribute>
			<xsl:attribute name="LOANAMOUNT"><xsl:value-of select="@LOANAMOUNT"/></xsl:attribute>			
			<xsl:attribute name="LOANTERM"><xsl:value-of select="concat(sum(@TERMINYEARS), ' Years ', sum(@TERMINMONTHS), ' Months')"></xsl:value-of></xsl:attribute>			
		</LOANCOMPONENT>
	</xsl:template>
</xsl:stylesheet>
