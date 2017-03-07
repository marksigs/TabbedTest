<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		04/07/2006	EP543			Incorporate 'Other' title, e.g. Lord, Baron, Reverend etc...
	PE		15/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		07/02/2007	EP2_484		Modified to correctly map loans to applicants
	DS	 	13/02/2007	EP2_1360	Formatted date
	OS		14/02/2007	EP2_484		Reverted to earlier version, main changes were required only in PrepareTemplate.xml
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/SECUREDLOANREFERENCE/APPLICATIONFACTFIND/APPLICATIONLOANSLIABILITIES/LOANSLIABILITIES[@AGREEMENTTYPE_TYPE_S and CUSTOMERROLE/@CUSTOMERNUMBER]"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="LOANSLIABILITIES">
		<xsl:element name="LETTER">
			<xsl:element name="SECUREDLOANDETAILS">
				<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
				<xsl:apply-templates select="ACCOUNT/THIRDPARTY/ADDRESS"/>
				<xsl:apply-templates select="ACCOUNT/THIRDPARTY/CONTACTDETAILS"/>
				<xsl:apply-templates select="ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
				<xsl:apply-templates select="ACCOUNT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
			</xsl:element>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="current()"/>
			</xsl:call-template>
			<xsl:if test="current()/CUSTOMERROLE/BANKCREDITCARD">
				<xsl:element name="BANKCCDETAILS">
					<xsl:attribute name="ADDITIONALDETAILS">
						<xsl:value-of select="current()/CUSTOMERROLE/BANKCREDITCARD[position()=1]/@ADDITIONALDETAILS"></xsl:value-of>
					</xsl:attribute>
				</xsl:element>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
