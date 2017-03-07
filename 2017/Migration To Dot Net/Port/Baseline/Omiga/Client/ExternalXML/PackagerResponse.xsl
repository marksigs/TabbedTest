<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description	
	PB 		15/09/2006 	EP1107
	PB 		19/09/2006 	EP1107 		Minor fix
	OS		25/01/2007	EP2_896		Modified according to updated specs
	GHun	31/01/2007	EP2_1147	Merged EP1205 Now includes CASETASK
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-packager.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE/APPLICATION"/>
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE//CUSTOMERROLE[@CUSTOMERROLETYPE='1']/CUSTOMERVERSION"/>
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE/APPLICATION/APPLICATIONINTRODUCER"/>
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE/NEWLOAN[position()=1]"/>
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE/RISKASSESSMENT[position()=1]"/>
			<xsl:apply-templates select="/RESPONSE/PACKAGERRESPONSE/APPLICATION/CASESTAGE/CASETASK[@TASKSTATUS='10']"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="@APPLICATIONNUMBER"/>}</xsl:attribute>
			<xsl:attribute name="PACKAGERREFERENCE">{<xsl:value-of select="@INGESTIONAPPLICATIONNUMBER"/>}</xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="CUSTOMERVERSION">
		<xsl:element name="APPLICANT">
			<xsl:attribute name="SURNAME"><xsl:value-of select="@SURNAME"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="NEWLOAN">
		<xsl:element name="NEWLOAN">
			<xsl:variable name="vTOTALLOANAMOUNT"><xsl:value-of select="sum(@AMOUNTREQUESTED)"/></xsl:variable>
			<xsl:variable name="vPPESTVAL"><xsl:value-of select="sum(../APPLICATION/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"></xsl:value-of></xsl:variable>
			<xsl:attribute name="TOTALLOANAMOUNT"><xsl:value-of select="$vTOTALLOANAMOUNT"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="PPESTVAL"><xsl:value-of select="$vPPESTVAL"></xsl:value-of></xsl:attribute>
			<xsl:attribute name="LTV"><xsl:value-of select="format-number(($vTOTALLOANAMOUNT div $vPPESTVAL), '###0.00%')"></xsl:value-of></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="RISKASSESSMENT">
		<xsl:element name="RISKASSESSMENT">
			<xsl:attribute name="DIPRESPONSE"><xsl:value-of select="@UNDERWRITERDECISION_TEXT"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="CASETASK">
		<xsl:element name="CASETASK">
			<xsl:attribute name="TASKNAME"><xsl:value-of select="@TASKNAME"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
