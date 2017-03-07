<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\code\WebServiceInterface\rawGetApplicationDataResponse.xml?>
<!-- Peter Edney - 16/03/2006 - MAR1444 - Added new CUSTOMERADDRESS template to exclude address types TH, TP and TP1 -->
<!-- IK - 19/09/2006 - EP2_1 drop MEMOPAD -->
<!-- IK - 14/11/1006 - EP2_57 drop APPLICATIONSTAGE,APPLICATIONSALUTATION,PACKAGERDETAIL,CUSTOMERSALUTATION, CUSTOMERCHANNEL -->
<!-- IK - 23/11/1006 - EP2_170 add APPLICATIONINTRODUCER, ADDITIONALDIPDATA (MEMOPAD), reinstate APPLICATIONSTAGE -->
<!-- IK - 01/12/1006 - EP2_279 drop MEMOPAD -->
<!-- IK - 09/01/2007 - EP2_702 - restructure for ACCOUNTRELATIONSHIP records -->
<!-- IK - 12/01/2007 - EP2_700 - drop OTHERARREARSACCOUNT -->
<!-- IK - 12/01/2007 - EP2_756 - drop CUSTOMERADDRESS for MORTGAGEACCOUNT  (type 5) -->
<!-- IK - 13/01/2007 - EP2_840 - return MEMOPAD (not ADDITIONALDIPDATA) -->
<!-- IK - 02/02/2007 - EP2_1207 - return INCOMESUMMARY -->
<!--PSC 06/03/2007	  EP2_1846 - Determine ADDTOLOAN for MORTGAGEONEOFFCOSTS -->

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:apply-templates select="@*"/>
			<xsl:apply-templates select="*"/>
		</xsl:copy>
	</xsl:template>

	<xsl:template match="@BMDECLARATIONIND"/>
	<xsl:template match="@NAMEANDADDRESSACTIVEFROM"/>
	<xsl:template match="@NAMEANDADDRESSACTIVETO"/>
	<xsl:template match="@DATEOFENTRY"/>
	<xsl:template match="@PRODUCTCODESEARCHIND"/>
	<!-- Peter Edney - 27/01/2006 - MAR1134 -->
	<!-- <xsl:template match="@EARLIESTCOMPLETIONDATE"/> -->
	<xsl:template match="@SALESLEADSTATUS"/>
	<xsl:template match="@FTTRANSFEROFEQUITYREQUIRED"/>
	<xsl:template match="@LASTOFFERED"/>
	<xsl:template match="@KFIREQUIREDINDICATOR"/>
	<xsl:template match="@CONFIRMEDMAXBORROWING"/>
	<xsl:template match="@CONFIRMEDCALCULATEDINCMULTIPLE"/>
	<xsl:template match="@CONFIRMEDMAXBORROWING"/>
	<xsl:template match="@CONFIRMEDCALCULATEDINCMULTIPLIERTYPE"/>
	<xsl:template match="@PRODUCTSWITCHRETAINPRODUCTIND"/>
	<xsl:template match="@CURRENTRATEEXPIRYDATE"/>
	<xsl:template match="@CARDFAILREASON"/>

	<xsl:template match="@*"><xsl:copy/></xsl:template>

	<xsl:template match="APPLICATIONBANKBUILDINGSOC">
		<xsl:copy>
			<xsl:apply-templates select="@*"/>
			<xsl:if test="not(@DDEXPLAINEDIND)">
				<xsl:attribute name="DDEXPLAINEDIND">0</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="CASEACTIVITY"/>
	<xsl:template match="APPLICATIONPRIORITY"/>
	<xsl:template match="USERHISTORY"/>
	<!-- EP2_840 -->
	<xsl:template match="ADDITIONALDIPDATA"/>
	<!-- Peter Edney - 16/03/2006 - MAR1444 -->
	<!-- Exclude address types TH, TP and TP1-->
	<xsl:template match="CUSTOMERADDRESS">
		<!-- EP2_756 -->
		<xsl:if test="@ADDRESSTYPE!=5 and @ADDRESSTYPE!=6 and @ADDRESSTYPE!=7 and @ADDRESSTYPE!=8">
			<xsl:copy-of select="."/>
		</xsl:if>
	</xsl:template>		
	<!-- IK - 14/11/1006 - EP2_57 drop APPLICATIONSTAGE,APPLICATIONSALUTATION, PACKAGERDETAIL,CUSTOMERSALUTATION, CUSTOMERCHANNEL -->
	<!-- IK - 23/11/1006 - EP2_170 reinstate APPLICATIONSTAGE -->
	<xsl:template match="APPLICATIONSALUTATION"/>
	<xsl:template match="PACKAGERDETAIL"/>
	<xsl:template match="CUSTOMERSALUTATION"/>
	<xsl:template match="CUSTOMERCHANNEL"/>
	<!-- EP2_700 -->
	<xsl:template match="OTHERARREARSACCOUNT"/>
	
	<xsl:template match="MORTGAGEPRODUCT"/>
	<xsl:template match="MORTGAGELENDER"/>

	<xsl:template match="MORTGAGEONEOFFCOST">
		<xsl:variable name="mortgageLender" select="../LOANCOMPONENT/MORTGAGEPRODUCT/MORTGAGELENDER[1]" />
		<xsl:copy>
			<xsl:apply-templates select="@*[name()!='IDENTIFIER']"/>
			<xsl:choose>
				<xsl:when test="@ADDTOLOAN='1'">
					<xsl:copy-of select="@ADDTOLOAN"/>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='ADM' and $mortgageLender/@ALLOWADMINFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='ARR' and $mortgageLender/@ALLOWARRGEMTFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='DEE' and $mortgageLender/@ALLOWDEEDSRELFEEADD='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='MIG' and $mortgageLender/@ALLOWMIGFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='POR' and $mortgageLender/@ALLOWPORTINGFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='REI' and $mortgageLender/@ALLOWREINSPTFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='REV' and $mortgageLender/@ALLOWREVALUATIONFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='SEA' and $mortgageLender/@ALLOWSEALINGFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='TTF' and $mortgageLender/@ALLOWTTFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='SRF' and $mortgageLender/@ALLOWTTFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='VAL' and $mortgageLender/@ALLOWVALNFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='TPV' and $mortgageLender/@ALLOWVALNFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='LEG' and $mortgageLender/@ALLOWLEGALFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='PSF' and $mortgageLender/@ALLOWPRODUCTSWITCHFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='OTH' and $mortgageLender/@ALLOWOTHERFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='IAF' and $mortgageLender/@ALLOWNONLENDINSADMINFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='FLF' and $mortgageLender/@ALLOWFURTHERLENDINGFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='AB' and $mortgageLender/@ALLOWFURTHERLENDINGFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
				<xsl:when test="@IDENTIFIER='CLI' and $mortgageLender/@ALLOWCREDITLIMITINCFEEADDED='1'">
						<xsl:attribute name="ADDTOLOAN">0</xsl:attribute>
				</xsl:when>
			</xsl:choose>
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>
