<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">

<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date				AQR			Description
	MAH	187/01/2007	EP2796 		Rewritten
	OS		12/02/2007		EP2_796		Removed DDDetails element.
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
   
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:attribute name="ACCOUNTNAME"><xsl:value-of select="/RESPONSE/TELEPHONEDDCONFIRMATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/@ACCOUNTNAME"/></xsl:attribute>
			<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="concat('{',/RESPONSE/TELEPHONEDDCONFIRMATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
			<xsl:attribute name="SORTCODE"><xsl:value-of select="concat('{',/RESPONSE/TELEPHONEDDCONFIRMATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC/THIRDPARTY/@THIRDPARTYBANKSORTCODE, '}')"/></xsl:attribute>
			<xsl:attribute name="DEBITAMOUNT"><xsl:value-of select="/RESPONSE/TELEPHONEDDCONFIRMATION/APPLICATIONFACTFIND/QUOTATION/@TOTALQUOTATIONCOST"/></xsl:attribute>
			<xsl:call-template name="APPLICANTINFO">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/TELEPHONEDDCONFIRMATION"/>
			</xsl:call-template>
			<!--xsl:element name="DDDETAILS"-->
			<!--/xsl:element-->
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->    	
	
</xsl:stylesheet>