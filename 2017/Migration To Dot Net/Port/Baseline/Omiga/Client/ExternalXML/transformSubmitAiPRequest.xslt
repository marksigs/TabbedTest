<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile				$Workfile: transformSubmitAiPRequest.xslt $	Copyright © 2006 Vertex Financial Services
Current Version:	$Revision: 17 $
Last Modified:		$JustDate: 29/03/07 $
Modified By:		$Author: Lesliem $
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

History:

Author		Date				Description
IK				24/06/2006	EP859 - day 0 data migration enhancements
SAB			03/07/2006	EP937 - Updated population of INTERMEDIARYPANELID and removed MEMOPAD creations
BSC			04/07/2006	EP939 - Added MEMOPAD creations to cater for Loan Amount, Applicant(s) Income and Product Code
IK				19/09/2006	EP2_1 fix namespace directice
LDM 		20/09/2006  	EP1116  For buy to let set self cert to true and income to 0 
IK				25/10/2006	associate with EP2_1
IK				25/10/2006	associate with EP2_10 - address targetting
IK				14/11/2006	EP2_106 - multi-component loans
IK				23/11/2006	EP2_170 - add APPLICATIONINTRODUCER, add ADDITIONALDIPDATA (memopad) drop PACKAGER
IK				23/11/2006	EP2_170 - do nor create QUOTATION hierarchy via ingestion (add null CRUD_OP)
IK				28/11/2006	EP2_224 - drop PACKAGER for real
IK				07/12/2006	EP2_342 - create NEWLOAN in all circumstances
IK				11/12/2006	EP2_398 - default CHANNELID
MCh 		11/12/2006	EP2_239 - Added UKTAXPAYERINDICATOR and moved INCOMESUMMARY/@ALLOWABLEANNUALINCOME logic.
MCh			13/12/2006	EP2_241 - Phase 2: EMPLOYMENT/DATESTARTEDORESTABLISHED now passed from web as a date. 
MCh			13/12/2006	EP2_238 - Added LOANSLIABILITIES processing ACCOUNT/LOANSLIABILITIES using stripNS 
MCh 		13/12/2006	EP2_237 - Added ALIAS call
IK				18/12/2006	EP2_571 - resequence default REQUEST attributes
IK				20/12/2006	EP2_585 - remove test for EARNEDINCOMEAMOUNT=1 (phase 1 frig), plus (temporary) fix for empty (01/01/0001) dates
IK				20/12/2006	EP2_582 - if previous name is ingested from Web, ALIASTYPE to default to Alias.
PSC			04/01/2007	EP2_633 - Use value CUSTOMERROLETYPE rather than defaulting to 1
IK				09/01/2007	EP2_702 - create ACCOUNTRELATIONSHIP records
IK				31/01/2007	EP2_1152	default USERAUTHORITYLEVEL on REQUEST
PE			30/01/2007	EP2_865 - Default PAYMENTFREQUENCYTYPE to 1 (Annual)
IK				02/01/2007	EP2_1050 - TYPEOFBUYER now input from web
IK				02/01/2007	EP2_599 - fix MEMOPAD entries for different income types
IK				08/02/2007	EP2_1049 - NEWPROPERTY,NEWPROPERTYDEPOSIT,UNEARNEDINCOME now input
IK				19/02/2007	EP2_1172 add TYPEOFVALUATION to APPLICATION
IK				12/03/2007	EP2_1352 add DIPDRAWDOWN to NEWLOAN
PSC			20/03/2007	EP2_1872 add LOCATION to APPLICATIONFACTFIND
IK				26/03/2007	EP2_1647 differentiate Customer & Guarantor on MEMOPAD entries
====================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:msg="http://Request.SubmitAiP.Omiga.vertex.co.uk" xmlns:OmigaAiPData="http://OmigaAiPData.Omiga.vertex.co.uk" xmlns:user="http://vertex.com/omiga4">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<msxsl:script language="JScript" implements-prefix="user"><![CDATA[
		function toPC(v)
		{
			var s = " " + v.toLowerCase();
			var a;
			while(a = s.match(/ [a-z]|'[a-z]|-[a-z]|mc[a-z]|Mc[a-z]/))
				s = s.substr(0,a.lastIndex -1) + s.substr(a.lastIndex -1,1).toUpperCase() + s.substring(a.lastIndex);
			s = s.replace(/ Bfpo /g," BFPO ");
			s = s.replace(/ Hms /g," HMS ");
			s = s.replace(/ Po /g," PO ");
			s = s.replace(/ Von /g," von ");
			s = s.replace(/ Van /g," van ");
			s = s.replace(/ De /g," de ");
					
			return(s.substring(1));
		}
		function toUC(v)
		{
			return(v.toUpperCase());
		}
		function startDate(v)
		{
			if(isNaN(parseInt(v))) return("");
			var dt = new Date();
			dt.setMonth((dt.getMonth() - parseInt(v)));
			var day = (dt.getDate() +100).toString().substr(1);
			var month = (dt.getMonth() +101).toString().substr(1);
			return(day + "/" + month + "/" + dt.getFullYear());
		}
	]]></msxsl:script>
	<xsl:variable name="firstPass">
		<xsl:for-each select="*">
			<xsl:call-template name="stripNS"/>
		</xsl:for-each>
	</xsl:variable>
	<xsl:template name="stripNS">
		<xsl:element name="{local-name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="local-name(.)='EXPERIANREF'">
				<xsl:value-of select="."/>
			</xsl:if>
			<xsl:if test="local-name(.)='ADDRESSTARGETING'">
				<xsl:value-of select="."/>
			</xsl:if>
			<xsl:for-each select="*">
				<xsl:call-template name="stripNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/">
		<xsl:for-each select="msxsl:node-set($firstPass)/REQUEST">
			<xsl:element name="REQUEST">
				<xsl:attribute name="OPERATION">SubmitAiP</xsl:attribute>
				<xsl:attribute name="USERID">epsom</xsl:attribute>
				<xsl:attribute name="UNITID">epsom</xsl:attribute>
				<xsl:attribute name="CHANNELID">2</xsl:attribute>
				<xsl:attribute name="omigaClient">epsom</xsl:attribute>
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:apply-templates select="*"/>
			</xsl:element>
		</xsl:for-each>
		<xsl:for-each select="msxsl:node-set($firstPass)/ATR_REQUEST">
			<xsl:element name="REQUEST">
				<xsl:attribute name="OPERATION">SubmitAiP</xsl:attribute>
				<xsl:attribute name="USERID">epsom</xsl:attribute>
				<xsl:attribute name="UNITID">epsom</xsl:attribute>
				<xsl:attribute name="CHANNELID">2</xsl:attribute>
				<xsl:attribute name="omigaClient">epsom</xsl:attribute>
				<xsl:attribute name="ADDRESSTARGETREQ">TRUE</xsl:attribute>
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:apply-templates select="*"/>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="local-name(../.) = 'REQUEST'">
				<!-- ep2_1172 -->
				<xsl:if test="APPLICATIONFACTFIND/NEWPROPERTY[@VALUATIONTYPE]">
					<xsl:attribute name="TYPEOFVALUATION"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/@VALUATIONTYPE"/></xsl:attribute>
				</xsl:if>
				<xsl:element name="USERHISTORY"/>
			</xsl:if>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="PACKAGER"/>
	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:element name="APPLICATIONFACTFIND">
			<xsl:if test="NEWPROPERTY[@PROPERTYLOCATION]">
				<xsl:attribute name="LOCATION"><xsl:value-of select="NEWPROPERTY/@PROPERTYLOCATION"/></xsl:attribute>
			</xsl:if>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="@NATUREOFLOAN='11'">
				<xsl:attribute name="SELFCERTIND">true</xsl:attribute>
			</xsl:if>
			<xsl:if test="local-name(../../.) = 'REQUEST'">
				<xsl:attribute name="NUMBEROFAPPLICANTS"><xsl:value-of select="count(CUSTOMERROLE)"/></xsl:attribute>
				<xsl:attribute name="INGESTIONMORTGAGEPRODUCTCODE"><xsl:value-of select="QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@MORTGAGEPRODUCTCODE"/></xsl:attribute>
				<xsl:attribute name="INGESTIONINTERESTRATE"><xsl:value-of select="QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT/@FINALRATE"/></xsl:attribute>
				<xsl:attribute name="PACKAGERINDIVIDUALTELEPHONE"><xsl:value-of select="/REQUEST/APPLICATION/PACKAGER/@PACKAGERINDIVIDUALTELEPHONE"/></xsl:attribute>
				<xsl:attribute name="PACKAGERINDIVIDUALEMAIL"><xsl:value-of select="/REQUEST/APPLICATION/PACKAGER/@PACKAGERINDIVIDUALEMAIL"/></xsl:attribute>
			</xsl:if>
			<!-- EP2_170 -->
			<xsl:apply-templates select="APPLICATIONINTRODUCER"/>
			<xsl:apply-templates select="APPLICATIONLEGALREP"/>
			<xsl:apply-templates select="QUOTATION"/>
			<xsl:apply-templates select="FINANCIALSUMMARY"/>
			<xsl:apply-templates select="CUSTOMERROLE"/>
			<!-- EP2_702 -->
			<xsl:apply-templates select="ACCOUNT"/>
			<!-- EP2_1049 -->
			<xsl:apply-templates select="NEWPROPERTY"/>
			<xsl:apply-templates select="NEWPROPERTYDEPOSIT"/>
			<!-- EP2_170 -->
			<xsl:apply-templates select="ADDITIONALDIPDATA"/>
			<xsl:if test="local-name(../../.) = 'REQUEST'">
				<xsl:element name="APPLICATIONSALUTATION">
					<xsl:attribute name="_IGNORENOWROWS">true</xsl:attribute>
				</xsl:element>
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Dip ingestion - Amount Requested &#163;<xsl:value-of select="QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
			<xsl:for-each select="QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGEPRODUCTCODE]">
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Dip ingestion - Product Code <xsl:value-of select="@MORTGAGEPRODUCTCODE"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
			<!-- EP2_1647 -->
			<!-- EP2_599 -->
			<xsl:for-each select="CUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT/EARNEDINCOME[@EARNEDINCOMETYPE = '10'][@EARNEDINCOMEAMOUNT != '0']">
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Customer <xsl:value-of select="../../../../@CUSTOMERORDER"/> Basic Income &#163;<xsl:value-of select="@EARNEDINCOMEAMOUNT"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="CUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT/EARNEDINCOME[@EARNEDINCOMETYPE != '10'][@EARNEDINCOMEAMOUNT != '0']">
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Customer <xsl:value-of select="../../../../@CUSTOMERORDER"/> Other Non-Guaranteed / Irregular Income &#163;<xsl:value-of select="@EARNEDINCOMEAMOUNT"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
			<!-- EP2_599 ends -->
			<xsl:for-each select="CUSTOMERROLE[@CUSTOMERROLETYPE='2']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT/EARNEDINCOME[@EARNEDINCOMETYPE = '10'][@EARNEDINCOMEAMOUNT != '0']">
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Guarantor  Basic Income &#163;<xsl:value-of select="@EARNEDINCOMEAMOUNT"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="CUSTOMERROLE[@CUSTOMERROLETYPE='2']/CUSTOMER/CUSTOMERVERSION/EMPLOYMENT/EARNEDINCOME[@EARNEDINCOMETYPE != '10'][@EARNEDINCOMEAMOUNT != '0']">
				<xsl:element name="MEMOPAD">
					<xsl:attribute name="ENTRYTYPE">9</xsl:attribute>
					<xsl:attribute name="MEMOENTRY">Guarantor  Other Non-Guaranteed / Irregular Income &#163;<xsl:value-of select="@EARNEDINCOMEAMOUNT"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
			<!-- EP2_1647 ends -->
		</xsl:element>
	</xsl:template>
	<xsl:template match="QUOTATION">
		<xsl:element name="NEWLOAN">
			<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="MORTGAGESUBQUOTE/@AMOUNTREQUESTED"/></xsl:attribute>
			<!-- EP2_1352 -->
			<xsl:if test="MORTGAGESUBQUOTE[@DRAWDOWN][@DRAWDOWN != 0]">
				<xsl:attribute name="DIPDRAWDOWN"><xsl:value-of select="MORTGAGESUBQUOTE/@DRAWDOWN"/></xsl:attribute>
			</xsl:if>
			<!-- EP2_1352 ends -->
		</xsl:element>
		<xsl:if test="MORTGAGESUBQUOTE/LOANCOMPONENT[@MORTGAGEPRODUCTCODE]">
			<!-- EP2_170 -->
			<xsl:call-template name="nullCrudOp"/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="CUSTOMERVERSION">
		<xsl:element name="CUSTOMERVERSION">
			<xsl:for-each select="@*">
				<xsl:choose>
					<xsl:when test="name() = 'FIRSTFORENAME'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'SECONDFORENAME'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'SURNAME'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'MARITALSTATUS'">
						<xsl:choose>
							<xsl:when test=". = '60'">
								<xsl:attribute name="MARITALSTATUS">25</xsl:attribute>
							</xsl:when>
							<xsl:when test=". = '70'">
								<xsl:attribute name="MARITALSTATUS">60</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:copy/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:copy/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
			<xsl:apply-templates select="CUSTOMERTELEPHONENUMBER"/>
			<xsl:apply-templates select="CUSTOMERADDRESS">
				<xsl:sort select="@ADDRESSTYPE"/>
				<xsl:sort data-type="number" select="concat(substring(@DATEMOVEDIN,7), substring(@DATEMOVEDIN,4,2), substring(@DATEMOVEDIN,1,2))"/>
			</xsl:apply-templates>
			<xsl:apply-templates select="EMPLOYMENT"/>
			<xsl:apply-templates select="INCOMESUMMARY"/>
			<!-- EP2_1049 -->
			<xsl:apply-templates select="UNEARNEDINCOME"/>
			<xsl:for-each select="ALIAS">
				<!-- EP2_582  -->
				<xsl:element name="ALIAS">
					<xsl:attribute name="ALIASTYPE">10</xsl:attribute>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:apply-templates/>
				</xsl:element>
			</xsl:for-each>
			<xsl:element name="CUSTOMERSALUTATION">
				<xsl:attribute name="_IGNORENOWROWS">true</xsl:attribute>
			</xsl:element>
			<xsl:element name="CUSTOMERCHANNEL">
				<xsl:attribute name="_IGNORENOWROWS">true</xsl:attribute>
				<xsl:attribute name="USERID"><xsl:value-of select="ancestor::REQUEST/@USERID"/></xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERADDRESS">
		<xsl:element name="CUSTOMERADDRESS">
			<xsl:for-each select="@*">
				<!-- temp. fix -->
				<xsl:choose>
					<xsl:when test=".='01/01/0001'"/>
					<xsl:otherwise>
						<xsl:copy/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ADDRESS">
		<xsl:element name="ADDRESS">
			<xsl:for-each select="@*">
				<xsl:choose>
					<xsl:when test="name() = 'BUILDINGORHOUSENAME'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<!--  SW 14/06/2006 EP718	Remove the text Flat from FLATNUMBER attribute -->
					<xsl:when test="name() = 'FLATNUMBER' and contains(translate(../@FLATNUMBER,'flat','FLAT'),'FLAT')">
						<xsl:attribute name="FLATNUMBER"><xsl:value-of select="substring-after(translate(../@FLATNUMBER,'abcdefghijklmnopqrstuvwxyz ','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'FLAT')"/></xsl:attribute>
					</xsl:when>
					<xsl:when test="name() = 'STREET'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'DISTRICT'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'TOWN'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'COUNTY'">
						<xsl:call-template name="toProper"/>
					</xsl:when>
					<xsl:when test="name() = 'POSTCODE'">
						<xsl:call-template name="toUpper"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:copy/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
			<xsl:if test="@COUNTRY=BF">
				<xsl:attribute name="BFPO">1</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="EMPLOYMENT">
		<xsl:element name="EMPLOYMENT">
			<xsl:attribute name="MAINSTATUS"><xsl:value-of select="'1'"/></xsl:attribute>
			<xsl:for-each select="@*">
				<!-- temp. fix -->
				<xsl:choose>
					<xsl:when test="name(.)='DATELEFTORCEASEDTRADING' and .='01/01/0001'"/>
					<xsl:otherwise>
						<xsl:copy/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="INCOMESUMMARY">
		<xsl:element name="INCOMESUMMARY">
			<xsl:variable name="NatureOfLoan" select="/REQUEST/APPLICATION/APPLICATIONFACTFIND[1]/@NATUREOFLOAN"/>
			<xsl:variable name="UKTaxPayerIndicator" select="@UKTAXPAYERINDICATOR"/>
			<xsl:attribute name="ALLOWABLEANNUALINCOME"><xsl:choose><xsl:when test="$NatureOfLoan and ($NatureOfLoan = 11)">0</xsl:when><xsl:otherwise><xsl:value-of select="..//EARNEDINCOME/@EARNEDINCOMEAMOUNT"/></xsl:otherwise></xsl:choose></xsl:attribute>
			<xsl:if test="$UKTaxPayerIndicator!=''">
				<xsl:attribute name="UKTAXPAYERINDICATOR"><xsl:value-of select="$UKTaxPayerIndicator"/></xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERROLE">
		<xsl:element name="CUSTOMERROLE">
			<xsl:copy-of select="@CUSTOMERROLETYPE"/>
			<xsl:attribute name="CUSTOMERORDER"><xsl:value-of select="position()"/></xsl:attribute>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="local-name(.)='EXPERIANREF'">
				<xsl:value-of select="."/>
			</xsl:if>
			<xsl:if test="local-name(.)='ADDRESSTARGETING'">
				<xsl:value-of select="."/>
			</xsl:if>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="toProper">
		<xsl:attribute name="{name()}"><xsl:value-of select="user:toPC(normalize-space(.))"/></xsl:attribute>
	</xsl:template>
	<xsl:template name="toUpper">
		<xsl:attribute name="{name()}"><xsl:value-of select="user:toUC(normalize-space(.))"/></xsl:attribute>
	</xsl:template>
	<!-- EP2_170 -->
	<xsl:template name="nullCrudOp">
		<xsl:element name="{local-name()}">
			<xsl:attribute name="CRUD_OP">NULL</xsl:attribute>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="nullCrudOp"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>

	<!-- EP2_702 -->

	<xsl:template match="ACCOUNT">
		<xsl:element name="ACCOUNT">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="ACCOUNTRELATIONSHIP">

		<xsl:variable name="cO"><xsl:value-of select="@CUSTOMERORDER"/></xsl:variable>

		<xsl:element name="ACCOUNTRELATIONSHIP">

			<xsl:for-each select="msxsl:node-set($firstPass)/REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE">
				<xsl:if test="@CUSTOMERORDER = $cO">
					<xsl:attribute name="CUSTOMERNUMBER">xpath://REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER = '<xsl:value-of select="$cO"/>']/@CUSTOMERNUMBER</xsl:attribute>
					<xsl:attribute name="CUSTOMERVERSIONNUMBER">xpath://REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER = '<xsl:value-of select="$cO"/>']/@CUSTOMERVERSIONNUMBER</xsl:attribute>
					<xsl:attribute name="CUSTOMERROLETYPE">xpath://REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER = '<xsl:value-of select="$cO"/>']/@CUSTOMERROLETYPE</xsl:attribute>
				</xsl:if>
			</xsl:for-each>

		</xsl:element>

	</xsl:template>

	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>

	<!-- EP2_702 ends -->

	<xsl:template match="EARNEDINCOME">
		<xsl:element name="EARNEDINCOME">
			<xsl:copy-of select="@*"/>
			<xsl:if test="not(@PAYMENTFREQUENCYTYPE)">
				<xsl:attribute name="PAYMENTFREQUENCYTYPE">
					<xsl:text>1</xsl:text>
				</xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>	

	<xsl:template match="NEWPROPERTY">
		<xsl:element name="NEWPROPERTY">
			<xsl:attribute name="VALUATIONTYPE">1</xsl:attribute>
			<xsl:attribute name="PROPERTYLOCATION">10</xsl:attribute>
			<xsl:copy-of select="@*"/>
		</xsl:element>
	</xsl:template>

</xsl:stylesheet>
