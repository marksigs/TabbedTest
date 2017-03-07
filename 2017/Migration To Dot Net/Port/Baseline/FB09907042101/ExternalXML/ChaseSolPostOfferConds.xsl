<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	OS		11/01/2007	EP2_813		Modified according to new specifications for offer date
	OS		12/02/2007	EP2_813		Modified offer date to display in the correct format
	OS		14/02/2007	EP2_813		Modified offer date according to new specs
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-solicitor.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/CHASESOLPOSTOFFERCONDS/APPLICATIONLEGALREP"/>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/CHASESOLPOSTOFFERCONDS"/>
			</xsl:call-template>
			<xsl:apply-templates select="/RESPONSE/CHASESOLPOSTOFFERCONDS/APPLICATIONOFFER[position()=1]"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="APPLICATIONOFFER">
		<xsl:element name="APPOFFER">
			 <xsl:if test="@OFFERISSUEDATE">
				<xsl:attribute name="OFFERISSUEDDATE"><xsl:value-of select="concat('{',msg:GetDate(string(@OFFERISSUEDATE)),'}')"/></xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
