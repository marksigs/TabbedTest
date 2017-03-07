<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		04/07/2006	EP543			Incorporate 'Other' title, e.g. Lord, Baron, Reverend etc...
	PE		15/08/2006	EP103			Reworked. Abstracted javascript functions
	DS		26/01/2007	EP2_483		Reworked, as per modified spec
	DS	 	14/02/2007	 EP2_1360	 Formatted date
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
		<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
		<xsl:attribute name="LENDERORGURREF" ><xsl:value-of select="concat('{',/RESPONSE/PREVIOUSLENDERREFERENCE/APPLICATIONFACTFIND/CUSTOMERROLE/APPLICATIONMORTGAGEACCOUNTS/MORTGAGEACCOUNT/ACCOUNT/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
		<xsl:for-each select="/RESPONSE/PREVIOUSLENDERREFERENCE/APPLICATIONFACTFIND/CUSTOMERROLE">
				<xsl:element name="INFORMATION">
					<xsl:apply-templates select="current()"/>
					<xsl:if test="position() > 1">
						<xsl:element name="PAGEBREAK"/>
					</xsl:if>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="MORTGAGEACCOUNT">

	<xsl:element name="LENDERDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="ACCOUNT/THIRDPARTY/ADDRESS"/>
			<xsl:apply-templates select="ACCOUNT/THIRDPARTY/CONTACTDETAILS"/>
			<xsl:apply-templates select="ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
			<xsl:apply-templates select="ACCOUNT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
			
		</xsl:element>
		
		
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
