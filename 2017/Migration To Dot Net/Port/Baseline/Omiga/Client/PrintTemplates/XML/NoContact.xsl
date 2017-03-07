<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">

<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB 		13/06/2006 	EP721			Not using correspondence address
	OS 		14/12/2006 	EP2_469		Added application type description, included library xsl
	OS		11/01/2007 	EP2_469		Rewrote according to new abstacted library
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
   
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:call-template name="APPLICANTINFO">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/NOCONTACT"/>
			</xsl:call-template>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->    	
	
</xsl:stylesheet>