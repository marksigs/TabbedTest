<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-solicitor.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/REVSOLSINSTRSNORTHERNIRELAND/APPLICATIONLEGALREP"/>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/REVSOLSINSTRSNORTHERNIRELAND"/>
			</xsl:call-template>
			<xsl:element name="TCWORDING">
				<xsl:attribute name="TCEDITIONYEAR">
					<xsl:value-of select="concat('{', /RESPONSE/REVSOLSINSTRSNORTHERNIRELAND/GLOBALPARAMETER[@NAME='TemplateTCEditionYearNIreland']/@STRING,'}')"/>
				</xsl:attribute>
			</xsl:element>
			<xsl:apply-templates select="/RESPONSE/REVSOLSINSTRSNORTHERNIRELAND/CUSTOMERROLE/NEWPROPERTY/COMBOVALIDATION[@VALIDATIONTYPE='L']"/>
			<xsl:apply-templates select="/RESPONSE/REVSOLSINSTRSNORTHERNIRELAND/OTHERRESIDENT/PERSON"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="COMBOVALIDATION">
		<xsl:element name="LEASEHOLD">
			<xsl:attribute name="LANDLORDNOTICEOFCHARGE"><xsl:text>Notice of Charge to Landlords</xsl:text></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="PERSON">
		<xsl:if test="msg:GetAge(string(@DATEOFBIRTH))>=18">
			<xsl:element name="OTHERRESIDENTS">
				<xsl:attribute name="OTHERRESIDENTNAME">
					<xsl:variable name="salutation" select="concat(string(@FIRSTFORENAME),' ',string(@SURNAME))"/>
					<xsl:value-of select="concat('deed of consent for ', $salutation, ' to complete')"/>
				</xsl:attribute>
			</xsl:element>			
		</xsl:if>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
