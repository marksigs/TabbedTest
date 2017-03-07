<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Archive                  $Archive: /Epsom Phase2/2 INT Code/ExternalXML/IntroducerPipeLineResponse.xslt $
Workfile                 $Workfile: IntroducerPipeLineResponse.xslt $
Current Version   	$Revision: 4 $
Last Modified       	$Modtime: 29/03/07 17:49 $
Modified By          	$Author: Lesliem $

Prog					Date					Description
IK						16/01/2007		EP2_816 additional data
IK						23/03/2007		EP2_2037 APPLICATIONUNDERWRITING.UNDERWRITERDECISION overrides RISKASSESSMENT.UNDERWRITERDECISION
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	
	<xsl:template match="node()">
		<xsl:apply-templates select="node()"/>
	</xsl:template>		
	
	<xsl:template match="CASE">
		<xsl:element name="{name()}">
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="node()"/>
			<xsl:apply-templates select="node()"  mode="AddChildren"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="APPLICATIONFACTFIND|LATESTHISTORY|OWNINGUNIT|STAGE|CASESTAGE|RISKASSESSMENT|MORTGAGESUBQUOTE|APPLICATIONUNDERWRITING">
		<xsl:copy-of select="@*"/>
		<xsl:apply-templates select="node()" />
	</xsl:template>
	
	<xsl:template match="APPLICANTS">
		<xsl:copy-of select="@*[local-name()!='GUARANTOR']"/>
		<xsl:if test="@GUARANTOR">
			<xsl:attribute name="GUARANTOR">
				<xsl:choose>
					<xsl:when test="@GUARANTOR='1'">true</xsl:when>
					<xsl:otherwise>false</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>		
		</xsl:if>
		<xsl:apply-templates select="node()" />
	</xsl:template>

	<xsl:template match="ORIGINALOWNER|CURRENTOWNER" mode="AddChildren">
		<xsl:element name="{name()}">
			<xsl:copy-of select="@*"/>
		</xsl:element>
		<xsl:apply-templates select="node()" mode="AddChildren"/>
	</xsl:template>

	<!-- EP2_816 -->
	<xsl:template match="VALUATION[count(@*) > 0 or count(../@*)]" mode="AddChildren">
		<xsl:element name="{name()}">
			<xsl:copy-of select="@*"/>
			  <xsl:choose>
						<xsl:when test="../@VALUATIONSTATUS='1'">
							<xsl:attribute name="VALUATIONACCEPTABLE">true</xsl:attribute>
							<xsl:attribute name="VALUATIONREFERRED">false</xsl:attribute>
						</xsl:when>
						<xsl:when test="../@VALUATIONSTATUS='0'">
							<xsl:attribute name="VALUATIONACCEPTABLE">false</xsl:attribute>
							<xsl:attribute name="VALUATIONREFERRED">true</xsl:attribute>
						</xsl:when>
					</xsl:choose>
		</xsl:element>
		<xsl:apply-templates select="node()" mode="AddChildren"/>
	</xsl:template>

	<xsl:template match="MAINAPPLICANT" mode="AddChildren">
		<xsl:element name="{name()}">
			<xsl:copy-of select="@*"/>
			<xsl:copy-of select="ADDRESS/@POSTCODE"/>
		</xsl:element>
		<xsl:apply-templates select="node()" mode="AddChildren"/>
	</xsl:template>

	<!-- EP2_816 -->
	<xsl:template match="PACKAGER" mode="AddChildren">
		<xsl:if test="count(@*) > 0">
			<xsl:element name="{name()}">
				<xsl:copy-of select="@*"/>
			</xsl:element>
		</xsl:if>
		<xsl:apply-templates select="node()" mode="AddChildren"/>
	</xsl:template>
	
</xsl:stylesheet>
