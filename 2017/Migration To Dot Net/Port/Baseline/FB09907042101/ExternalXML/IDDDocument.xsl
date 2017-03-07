<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description

	OS		15/01/2007	EP2_809		XSL file created
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="IDDDETAILS">
				<xsl:attribute name="TTFEE"><xsl:value-of select="RESPONSE/IDDDOCUMENT/TTFEEBAND/@AMOUNT"/></xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->    	
</xsl:stylesheet>
