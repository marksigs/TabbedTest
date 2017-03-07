<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		07/08/2006	EP103			Reworked. Abstracted javascript functions
	DS		06/02/2007	EP2_479		Attaching this file to this defect
	OS		12/02/2007	EP2_479		Reformatted date to be processed as a string
	===============================================================================================================-->
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="ACCOUNTANT">
		<xsl:element name="ACCOUNTANTDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="THIRDPARTY/ADDRESS"/>
			<xsl:apply-templates select="THIRDPARTY/CONTACTDETAILS"/>
			<xsl:apply-templates select="NAMEANDADDRESSDIRECTORY/ADDRESS"/>
			<xsl:apply-templates select="NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
