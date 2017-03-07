<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	OS		15/12/2006	EP2_475		Added Application type
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-solicitor.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/REQUESTSOLICITORBANKDETAILS/APPLICATIONLEGALREP"/>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/REQUESTSOLICITORBANKDETAILS"/>
			</xsl:call-template>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->    
</xsl:stylesheet>
