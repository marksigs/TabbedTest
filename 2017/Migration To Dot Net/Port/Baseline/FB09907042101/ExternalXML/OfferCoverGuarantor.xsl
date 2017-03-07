<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description

	OS		10/01/2007	EP2_724		XSL file created
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:call-template name="GUARANTOR">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/OFFERCOVERGUARANTOR/APPLICATION/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']"/>
			</xsl:call-template>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/OFFERCOVERGUARANTOR/APPLICATION"/>
			</xsl:call-template>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->    	
</xsl:stylesheet>
