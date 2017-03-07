<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Abstracted solicitor templates.
	OS		12/02/2007	EP2_813		Modified date attribute
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="APPLICATIONLEGALREP">
		<xsl:element name="SOLICITORDETAILS">
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:apply-templates select="THIRDPARTY/ADDRESS"/>
			<xsl:apply-templates select="THIRDPARTY/CONTACTDETAILS"/>
			<xsl:apply-templates select="NAMEANDADDRESSDIRECTORY/ADDRESS"/>
			<xsl:apply-templates select="NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
