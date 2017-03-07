<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Abstracted packeger templates.
	OS		29/12/2006	EP2_455		Added new template for packager data
	DS		07/02/2007	EP2_455		Modified GetDate() function call to fix formatting issues
	===============================================================================================================-->
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="INTERMEDIARY">
		<xsl:element name="PACKAGERDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="ADDRESS"/>
			<xsl:apply-templates select="CONTACTDETAILS"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="APPLICATIONINTRODUCER">
		<xsl:element name="PACKAGERDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="PRINCIPALFIRM"/>
			<xsl:call-template name="APPLICATIONINTRODUCER"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
