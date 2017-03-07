<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description

	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
    DS		29/12/2006	EP2_504		Added new templates for EP2_504 template
	MAH		09/01/2007	EP2_42		Altered order of ARFIRM
	MAH		19/01/2007	EP2_472		Added braces to date to preserve format
	===============================================================================================================-->
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="INTERMEDIARY">
		<xsl:element name="BROKERDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="ADDRESS"/>
			<xsl:apply-templates select="CONTACTDETAILS"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="APPLICATIONINTRODUCER">
		<xsl:element name="BROKERDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="ARFIRM"/>
			<xsl:apply-templates select="PRINCIPALFIRM"/>
			<xsl:call-template name="APPLICATIONINTRODUCER"/>
		</xsl:element>
	</xsl:template>
	
	<!--============================================================================================================-->
</xsl:stylesheet>
