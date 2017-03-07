<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		14/09/2006	EP1138		Created
	OS		24/01/2007	EP2_939		Modified according to new specs
	OS		13/02/2007	EP2_939		Removed redundant fields and updated according to new specs
	OS		26/02/2007	EP2_939		Added £ for currency fields, fixed number formatting errors
	OS		28/02/2007	EP2_939		Fixed error in DepositAmount calculations
	OS		02/04/2007	EP2_1452	Updated template according to new specs
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="AUTONUMBERING">
				<xsl:attribute name="NAME"><xsl:value-of select="'SECTION'"></xsl:value-of></xsl:attribute>
			</xsl:element>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION/ALIAS/PERSON"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERADDRESS[not(@DATEMOVEDOUT)]"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERADDRESS[@DATEMOVEDOUT]"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/INTERMEDIARYORGANISATION"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION/EMPLOYMENT"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/APPLICATIONINTRODUCER"/>
			<xsl:apply-templates select="/RESPONSE/DIPSUMMARY/APPLICATION/MEMOPAD"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:attribute name="ApplicationNumber">{<xsl:value-of select="@APPLICATIONNUMBER"/>}</xsl:attribute>
			<xsl:attribute name="PackagerReference">{<xsl:value-of select="@INGESTIONAPPLICATIONNUMBER"/>}</xsl:attribute>
			<xsl:attribute name="DIPDate"><xsl:value-of select="@APPLICATIONDATE"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="CUSTOMERVERSION">
		<xsl:element name="APPLICANT">
			<xsl:attribute name="Title"><xsl:value-of select="@TITLE_TEXT"/></xsl:attribute>
			<xsl:attribute name="Forename1"><xsl:value-of select="@FIRSTFORENAME"/></xsl:attribute>
			<xsl:attribute name="Forename2"><xsl:value-of select="@SECONDFORENAME"/></xsl:attribute>
			<xsl:attribute name="Surname"><xsl:value-of select="@SURNAME"/></xsl:attribute>
			<xsl:attribute name="ApplicantName"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(@TITLE_TEXT),
					string(@TITLEOTHER),
					string(@FIRSTFORENAME),																																	    									
					string(@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="DOB"><xsl:value-of select="concat('{',@DATEOFBIRTH,'}')"/></xsl:attribute>
			<xsl:attribute name="Age"><xsl:value-of select="msg:GetAge(string(@DATEOFBIRTH))"/></xsl:attribute>
			<xsl:attribute name="MaritalStatus"><xsl:value-of select="@MARITALSTATUS_TEXT"/></xsl:attribute>
			<xsl:attribute name="Gender"><xsl:value-of select="@GENDER_TEXT"/></xsl:attribute>
			<xsl:attribute name="Nationality"><xsl:value-of select="@NATIONALITY_TEXT"/></xsl:attribute>
			<xsl:attribute name="CustomerRole"><xsl:value-of select="../@CUSTOMERROLETYPE_TEXT"></xsl:value-of></xsl:attribute>
		</xsl:element>
	</xsl:template>
	
		<!--============================================================================================================-->
	<xsl:template match="PERSON">
		<xsl:element name="PREVIOUSNAMES">
			<xsl:attribute name="PreviousTitle"><xsl:value-of select="@TITLE_TEXT"/></xsl:attribute>
			<xsl:attribute name="PreviousForename1"><xsl:value-of select="@FIRSTFORENAME"/></xsl:attribute>
			<xsl:attribute name="PreviousForename2"><xsl:value-of select="@SECONDFORENAME"/></xsl:attribute>
			<xsl:attribute name="PreviousSurname"><xsl:value-of select="@SURNAME"/></xsl:attribute>
			<xsl:attribute name="FullName"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(../../@TITLE_TEXT),
					string(../../@TITLEOTHER),
					string(../../@FIRSTFORENAME),																																	    									
					string(../../@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="DateOfChange"><xsl:value-of select="../@DATEOFCHANGE"/></xsl:attribute>
		</xsl:element>
	</xsl:template>

	<!--============================================================================================================-->
	<xsl:template match="CUSTOMERADDRESS">
		<xsl:element name="ADDRESSHISTORY">
			<xsl:attribute name="FullName"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(../CUSTOMERVERSION/@TITLE_TEXT),
					string(../CUSTOMERVERSION/@TITLEOTHER),
					string(../CUSTOMERVERSION/@FIRSTFORENAME),																																	    									
					string(../CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="FullAddress"><xsl:value-of select="msg:DealWithStraightAddress(
					string(ADDRESS/@FLATNUMBER),
					string(ADDRESS/@BUILDINGORHOUSENAME),
					string(ADDRESS/@BUILDINGORHOUSENUMBER),
					string(ADDRESS/@STREET),
					string(ADDRESS/@DISTRICT),
					string(ADDRESS/@TOWN),
					string(ADDRESS/@COUNTY),
					string(ADDRESS/@POSTCODE))"/></xsl:attribute>
			<xsl:if test="@DATEMOVEDOUT">
				<xsl:attribute name="TimeAtAddress"><xsl:value-of select="msg:GetYears(string(@DATEMOVEDIN), string(@DATEMOVEDOUT))"/><xsl:text> years </xsl:text><xsl:value-of select="msg:GetMonths(string(@DATEMOVEDIN), string(@DATEMOVEDOUT))"/><xsl:text> months</xsl:text></xsl:attribute>
			</xsl:if>
			<xsl:if test="not(@DATEMOVEDOUT) and @DATEMOVEDIN">
				<xsl:attribute name="TimeAtAddress"><xsl:value-of select="msg:GetYears(string(@DATEMOVEDIN), '')"/><xsl:text> years </xsl:text><xsl:value-of select="msg:GetMonths(string(@DATEMOVEDIN), '')"/><xsl:text> months</xsl:text></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="Occupancy"><xsl:value-of select="@NATUREOFOCCUPANCY_TEXT"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:element name="APPLICATIONDETAILS">			
			<!--Phase 2 fields-->
			<xsl:attribute name="Scheme"><xsl:value-of select="@PRODUCTSCHEME_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="Status"><xsl:value-of select="@APPLICATIONINCOMESTATUS_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="Category"><xsl:value-of select="@NATUREOFLOAN_TEXT"/></xsl:attribute>
			<xsl:attribute name="TypeOfApp"><xsl:value-of select="@TYPEOFAPPLICATION_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="TypeOfBuyer">
				<xsl:choose>
					<xsl:when test="@TYPEOFAPPLICATION='10'"><xsl:value-of select="../COMBOVALIDATIONTN/COMBOVALUE/@VALUENAME"></xsl:value-of></xsl:when>
					<xsl:otherwise><xsl:value-of select="../COMBOVALIDATIONTR/COMBOVALUE/@VALUENAME"></xsl:value-of></xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
			<xsl:attribute name="Score"><xsl:value-of select="../APPLICATIONCREDITCHECK[@SUCCESSINDICATOR=1][last()]/CREDITCHECKSCORE[last()]/@SCORE"/></xsl:attribute>
			<xsl:attribute name="SchemeEligibility"><xsl:value-of select="../APPLICATIONCREDITCHECK[@SUCCESSINDICATOR=1][last()]/BESPOKEDECISION/COMBOVALIDATIONSG/COMBOVALUE/@VALUENAME"/></xsl:attribute>
			<xsl:attribute name="ExpReference"><xsl:value-of select="concat('{', ../APPLICATIONCREDITCHECK[@SUCCESSINDICATOR=1][last()]/@CREDITCHECKREFERENCENUMBER, '}')"/></xsl:attribute>

			<xsl:variable name="vAmountRequested"><xsl:value-of select="sum(../QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED)"></xsl:value-of></xsl:variable>
			<xsl:variable name="vDrawDown"><xsl:value-of select="sum(../QUOTATION/MORTGAGESUBQUOTE/@DRAWDOWN)"></xsl:value-of></xsl:variable>
			<xsl:variable name="vPurchasePrice"><xsl:value-of select="sum(../QUOTATION/MORTGAGESUBQUOTE/@PURCHASEPRICEORESTIMATEDVALUE)"></xsl:value-of></xsl:variable>
			<xsl:variable name="vDiscountAmount"><xsl:value-of select="sum(../NEWPROPERTY/@DISCOUNTAMOUNT)"></xsl:value-of></xsl:variable>

			<xsl:attribute name="AdvanceAmount"><xsl:value-of select="($vAmountRequested - $vDrawDown)"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="AmountRequested"><xsl:value-of select="$vAmountRequested"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="PurchasePrice"><xsl:value-of select="$vPurchasePrice"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="DepositAmount"><xsl:value-of select="sum(../NEWPROPERTYDEPOSIT/@AMOUNT)"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="MainDepositSource"><xsl:value-of select="../NEWPROPERTYDEPOSIT/@SOURCEOFFUNDING_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="LTV"><xsl:value-of select="concat(../QUOTATION/MORTGAGESUBQUOTE/@LTV, '%')"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="DiscountedPurchasePrice"><xsl:value-of select="($vPurchasePrice - $vDiscountAmount)"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="Discount"><xsl:value-of select="$vDiscountAmount"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="PropertyLocation"><xsl:value-of select="../NEWPROPERTY/@PROPERTYLOCATION_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="BTLRentalIncome"><xsl:value-of select="../NEWPROPERTY/@MONTHLYRENTALINCOME"></xsl:value-of></xsl:attribute>
			
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="EMPLOYMENT">
		<xsl:element name="EMPLOYMENTSTATUS">
			<xsl:attribute name="EmploymentStatus"><xsl:value-of select="@EMPLOYMENTSTATUS_TEXT"/></xsl:attribute>
			<xsl:attribute name="TimeEmployed">
				<xsl:value-of select="msg:GetYears(string(@DATESTARTEDORESTABLISHED), '')"/>
				<xsl:text> years </xsl:text>
				<xsl:value-of select="msg:GetMonths(string(@DATESTARTEDORESTABLISHED), '')"/>
				<xsl:text> months</xsl:text></xsl:attribute>
			<xsl:attribute name="GrossIncome"><xsl:value-of select="../INCOMESUMMARY/@ALLOWABLEANNUALINCOME"/></xsl:attribute>
			<xsl:attribute name="FullName"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(../@TITLE_TEXT),
					string(../@TITLEOTHER),
					string(../@FIRSTFORENAME),																																	    									
					string(../@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="JobTitle"><xsl:value-of select="@JOBTITLE"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="ContractRenewable">
				<xsl:if test="@EMPLOYMENTSTATUS='30'">
					<xsl:choose>
						<xsl:when test="CONTRACTDETAILS[@LIKELYTOBERENEWED='1']"><xsl:value-of select="'Yes'"></xsl:value-of></xsl:when>
						<xsl:otherwise><xsl:value-of select="'No'"></xsl:value-of></xsl:otherwise>
					</xsl:choose>
				</xsl:if>
			</xsl:attribute>
			<xsl:attribute name="AllowableIncome"><xsl:value-of select="../INCOMESUMMARY/@NETALLOWABLEANNUALINCOME"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="LOANCOMPONENT">
		<xsl:element name="LOANCOMPONENT">
			<xsl:attribute name="ProductCode"><xsl:value-of select="concat('{', @MORTGAGEPRODUCTCODE, '}')"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="PurposeOfLoan"><xsl:value-of select="@PURPOSEOFLOAN_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="SubPurposeOfLoan"><xsl:value-of select="@SUBPURPOSEOFLOAN_TEXT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="Amount"><xsl:value-of select="@LOANAMOUNT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="Term"><xsl:value-of select="concat(sum(@TERMINYEARS), ' Years ', sum(@TERMINMONTHS), ' Months')"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="RepaymentType"><xsl:value-of select="@REPAYMENTMETHOD_TEXT"></xsl:value-of></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="INTRODUCER">
	<xsl:choose>
		<xsl:when test="@INTRODUCERTYPE='1'">
			<xsl:element name="INTRODUCERBROKER">
				<xsl:attribute name="BrokerCompany">
				<xsl:choose>
					<xsl:when test="INTRODUCERFIRM/ARFIRM"><xsl:value-of select="INTRODUCERFIRM/ARFIRM/@ARFIRMNAME"></xsl:value-of></xsl:when>
					<xsl:when test="INTRODUCERFIRM/PRINCIPALFIRM"><xsl:value-of select="INTRODUCERFIRM/PRINCIPALFIRM/@PRINCIPALFIRMNAME"></xsl:value-of></xsl:when>
				</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name="IntroducerType"><xsl:value-of select="@INTRODUCERTYPE_TEXT"></xsl:value-of></xsl:attribute>
				<xsl:attribute name="BrokerIndividual">
					<xsl:value-of select="msg:GetApplicantNameWithInitials(string(ORGANISATIONUSER/@USERTITLE_TEXT), '',
																	string(ORGANISATIONUSER/@USERFORENAME),
																	string(ORGANISATIONUSER/@USERSURNAME))"></xsl:value-of>
				</xsl:attribute>
				<xsl:attribute name="BrokerIndividualEmail"><xsl:value-of select="@EMAILADDRESS"></xsl:value-of></xsl:attribute>
				<xsl:attribute name="BrokerIndividualTelephone">
					<xsl:value-of select="concat('{', ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@COUNTRYCODE,
																	ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@AREACODE,
																	ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@TELENUMBER, '}')"></xsl:value-of>
				</xsl:attribute>
			</xsl:element>
		</xsl:when>
		<xsl:when test="@INTRODUCERTYPE='2'">
			<xsl:element name="INTRODUCERPACKAGER">
				<xsl:attribute name="PackagerCompany">
				<xsl:choose>
					<xsl:when test="INTRODUCERFIRM/PRINCIPALFIRM"><xsl:value-of select="INTRODUCERFIRM/PRINCIPALFIRM/@PRINCIPALFIRMNAME"></xsl:value-of></xsl:when>
				</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name="IntroducerType"><xsl:value-of select="@INTRODUCERTYPE_TEXT"></xsl:value-of></xsl:attribute>
				<xsl:attribute name="PackagerIndividual">
					<xsl:value-of select="msg:GetApplicantNameWithInitials(string(ORGANISATIONUSER/@USERTITLE_TEXT), '',
																	string(ORGANISATIONUSER/@USERFORENAME),
																	string(ORGANISATIONUSER/@USERSURNAME))"></xsl:value-of>
				</xsl:attribute>
				<xsl:attribute name="PackagerIndividualEmail"><xsl:value-of select="@EMAILADDRESS"></xsl:value-of></xsl:attribute>
				<xsl:attribute name="PackagerIndividualTelephone">
					<xsl:value-of select="concat('{', ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@COUNTRYCODE,
																	ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@AREACODE,
																	ORGANISATIONUSERCONTACTDETAILS/CONTACTTELEPHONEDETAILS/@TELENUMBER, '}')"></xsl:value-of>
				</xsl:attribute>
			</xsl:element>
		</xsl:when>
	</xsl:choose>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="MEMOPAD">
		<xsl:if test="@MEMOENTRY">
			<xsl:element name="MEMOPAD">
				<xsl:attribute name="MemoPadEntry"><xsl:value-of select="@MEMOENTRY"></xsl:value-of></xsl:attribute>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
