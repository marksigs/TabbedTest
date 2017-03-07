<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: RunXMLUpgrade

History:

Version Author		Date       	Description
01.00	LDM		28/03/2006	Created . EP6 Epsom. Transform KYC data from experian prior to saving away
01.01	AW		04/05/2006	EP446 Amended paths for policy rules
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="Response" select="/GEODS"/>
	<xsl:variable name="KYCRoot" select="$Response/REQUEST"/>
	<xsl:variable name="Autp" select="$KYCRoot/AUTP"/>
	<xsl:variable name="PolicyTextArray" select="$KYCRoot/AUTP/POLICYTEXTS/HRPOLICYTEXT"/>
	
<xsl:template match="/">
<EXPERIAN>
	<xsl:element name="EXP">
		<xsl:attribute name="EXPERIANREF"><xsl:value-of select="$KYCRoot/@EXP_ExperianRef"/></xsl:attribute>
	</xsl:element>
	
	<xsl:element name="AUTP">
		<xsl:attribute name="NOPRIMDATAIDANDCURRADD"><xsl:value-of select="$Autp/NUMPRIMDATIACA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMSOURCEIDANDCURRADD"><xsl:value-of select="$Autp/NUMPRIMSRCIACA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDPRIMDATAIDANDCURRADD"><xsl:value-of select="$Autp/OLDPRIMIACA"/></xsl:attribute>
		<xsl:attribute name="NOSECDATAIDANDCURRADD"><xsl:value-of select="$Autp/NUMSECDATIACA"/></xsl:attribute>
		<xsl:attribute name="NOSECSOURCEIDANDCURRADD"><xsl:value-of select="$Autp/NUMSECSRCIACA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDSECDATAIDANDCURRADD"><xsl:value-of select="$Autp/OLDSECIACA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMDATAONLYCURRADD"><xsl:value-of select="$Autp/NUMPRIMDATAOCA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMSOURCEONLYCURRADD"><xsl:value-of select="$Autp/NUMPRIMSRCAOCA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDPRIMDATAONLYCURRADD"><xsl:value-of select="$Autp/OLDPRIMAOCA"/></xsl:attribute>
		<xsl:attribute name="NOSECDATAONLYCURRADD"><xsl:value-of select="$Autp/NUMSECDATAOCA"/></xsl:attribute>
		<xsl:attribute name="NOSECSOURCEONLYCURRADD"><xsl:value-of select="$Autp/NUMSECSRCAOCA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDSECDATAONLYCURRADD"><xsl:value-of select="$Autp/OLDSECAOCA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMDATAIDANDPREVADD"><xsl:value-of select="$Autp/NUMPRIMDATIAPA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMSOURCEIDANDPREVADD"><xsl:value-of select="$Autp/NUMPRIMSRCIAPA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDPRIMDATAIDANDPREVADD"><xsl:value-of select="$Autp/OLDPRIMIAPA"/></xsl:attribute>
		<xsl:attribute name="NOSECDATAIDANDPREVADD"><xsl:value-of select="$Autp/NUMSECDATIAPA"/></xsl:attribute>
		<xsl:attribute name="NOSECSOURCEIDANDPREVADD"><xsl:value-of select="$Autp/NUMSECSRCIAPA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDSECDATAIDANDPREVADD"><xsl:value-of select="$Autp/OLDSECIAPA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMDATAONLYPREVADD"><xsl:value-of select="$Autp/NUMPRIMDATAOPA"/></xsl:attribute>
		<xsl:attribute name="NOPRIMSOURCEONLYPREVADD"><xsl:value-of select="$Autp/NUMPRIMSRCAOPA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDPRIMDATAONLYPREVADD"><xsl:value-of select="$Autp/OLDPRIMAOPA"/></xsl:attribute>
		<xsl:attribute name="NOSECDATAONLYPREVADD"><xsl:value-of select="$Autp/NUMSECDATAOPA"/></xsl:attribute>
		<xsl:attribute name="NOSECSOURCEONLYPREVADD"><xsl:value-of select="$Autp/NUMSECSRCAOPA"/></xsl:attribute>
		<xsl:attribute name="DATEOLDSECDATAONLYPREVADD"><xsl:value-of select="$Autp/OLDSECAOPA"/></xsl:attribute>
		<xsl:attribute name="NOOFAGEMATCHPRIM"><xsl:value-of select="$Autp/AGEMATCHPRIM"/></xsl:attribute>
		<xsl:attribute name="NOOFAGEMATCHSEC"><xsl:value-of select="$Autp/AGEMATCHSEC"/></xsl:attribute>
		<xsl:attribute name="NOOFTIMECURRADDMATCHPRIM"><xsl:value-of select="$Autp/TACAMATCHPRIM"/></xsl:attribute>
		<xsl:attribute name="NOOFTIMECURRADDMATCHSEC"><xsl:value-of select="$Autp/TACAMATCHSEC"/></xsl:attribute>
		<xsl:attribute name="AUTHDECISION"><xsl:value-of select="$Autp/DECISION"/></xsl:attribute>
		<xsl:attribute name="AUTHDECISIONTEXT"><xsl:value-of select="$Autp/DECISIONTEXT"/></xsl:attribute>
		<xsl:attribute name="AUTHINDEX"><xsl:value-of select="$Autp/AUTHINDEX"/></xsl:attribute>
		<xsl:attribute name="AUTHINDEXTEXT"><xsl:value-of select="$Autp/AUTHINDEXTEXT"/></xsl:attribute>
		<xsl:attribute name="IDCONFIRMEDLEVEL"><xsl:value-of select="$Autp/IDCONFLEVEL"/></xsl:attribute>
		<xsl:attribute name="IDCONFIRMEDLEVELTEXT"><xsl:value-of select="$Autp/IDCONFLEVELTEXT"/></xsl:attribute>
		<xsl:attribute name="NOHIGHRISKPOLICYRULES"><xsl:value-of select="$Autp/NUMHRPOLICYRULES"/></xsl:attribute>
		<xsl:attribute name="CIFASREF"><xsl:value-of select="$Autp/CIFASREF"/></xsl:attribute>
	</xsl:element>

	<xsl:call-template name="CreatePolicys"/>
	
</EXPERIAN>
</xsl:template>

<xsl:template name="CreatePolicys">
	<xsl:variable name="PolicyRuleArray" select="$KYCRoot/AUTP/POLICYRULES/HRPOLICYRULE"/>
	<xsl:if test="$PolicyRuleArray">
		<xsl:element name="POLICYBLOCK">
			<xsl:for-each select="$PolicyRuleArray">
				<xsl:element name="POLICY">
					<xsl:attribute name="HIGHRISKPOLICYRULE"><xsl:value-of select="."/></xsl:attribute>
					<xsl:call-template name="GetRuleText">
						<xsl:with-param name="Position"><xsl:value-of select="position()"/></xsl:with-param>
					</xsl:call-template>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:if>
</xsl:template>

<xsl:template name="GetRuleText">
	<xsl:param name="Position"/>
		<xsl:if test="$PolicyTextArray[$Position]">
			<xsl:attribute name="HIGHRISKPOLICYTEXT">
				<xsl:value-of select="$PolicyTextArray[number($Position)]"/>
			</xsl:attribute>	
		</xsl:if>
</xsl:template>
	
</xsl:stylesheet>
