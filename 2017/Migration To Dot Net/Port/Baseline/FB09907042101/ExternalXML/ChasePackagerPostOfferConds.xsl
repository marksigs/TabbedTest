<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description

	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	DS 	    12/02/2007	EP2_738		Formatted the date
	OS		13/02/2007	EP2_1361	Removed APPLICATIONCONDITIONS node and updated OFFERISSUEDATE according to specs
	OS		14/02/2007	EP2_813		Modified offer date according to new specs
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-packager.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/CHASEPACKAGERPOSTOFFERCONDS/APPLICATION/APPLICATIONINTRODUCER"/>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/CHASEPACKAGERPOSTOFFERCONDS"/>
			</xsl:call-template>
			<xsl:apply-templates select="/RESPONSE/CHASEPACKAGERPOSTOFFERCONDS/APPLICATIONOFFER[position()=1]"/>
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
