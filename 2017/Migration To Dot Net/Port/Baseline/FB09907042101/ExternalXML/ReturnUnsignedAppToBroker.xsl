<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date				AQR			Description

	MAH		15/12/2006	EP2_473	Based on Packager template details for Unsigned application chasing.
	MAH		09/01/2007	EP2_473	Added APPLICATIONINTRODUCER
	MAH		19/01/2007	EP2_473	Updated APPLICANT
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-broker.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="APPLICANT">
				<xsl:apply-templates select="/RESPONSE/RETURNUNSIGNEDAPPTOBROKER"/>
			</xsl:element>
			<xsl:apply-templates select="/RESPONSE/RETURNUNSIGNEDAPPTOBROKER/APPLICATION/APPLICATIONINTRODUCER"/>
			<!--<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/RETURNUNSIGNEDAPPTOBROKER"/>
			</xsl:call-template>-->
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
