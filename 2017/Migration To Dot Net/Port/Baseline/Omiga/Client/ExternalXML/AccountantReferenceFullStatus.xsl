<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		04/07/2006	EP543			Incorporate 'Other' title, e.g. Lord, Baron, Reverend etc...
	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	OS		10/01/2007	EP2_479		Modified to incorporate multiple letter generation functionality
	OS		17/01/2007	EP2_479		Modified to handle guarantors in the same template
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-accountant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustRole" select="count(RESPONSE/ACCOUNTANTREFERENCEFULLSTATUS/CUSTOMERROLE[EMPLOYMENT/ACCOUNTANT])"/>
			<xsl:for-each select="/RESPONSE/ACCOUNTANTREFERENCEFULLSTATUS/CUSTOMERROLE[EMPLOYMENT/ACCOUNTANT]">
					<xsl:element name="INFORMATION">
						<xsl:apply-templates/>
						<xsl:if test="position() > 1">
							<xsl:element name="PAGEBREAK"/>
						</xsl:if>
					</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
