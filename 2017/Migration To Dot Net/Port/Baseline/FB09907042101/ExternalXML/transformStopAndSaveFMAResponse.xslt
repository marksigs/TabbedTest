<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Archive                  $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformStopAndSaveFMAResponse.xslt $
Workfile                 $Workfile: transformStopAndSaveFMAResponse.xslt $
Current Version   	$Revision: 1 $
Last Modified       	$Modtime: 28/11/06 10:07 $
Modified By          	$Author: Dbarraclough $

Prog					Date					Description
IK						18/11/2006		EP2_134 created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0"  xmlns="http://Response.SubmitStopAndSaveFMA.Omiga.vertex.co.uk" xmlns:msgdt="http://msgtypes.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="MESSAGE"/>
			<xsl:apply-templates select="ERROR"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MESSAGE">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="MESSAGE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ERROR">
		<xsl:element name="ERROR">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="msgdtNS">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="{local-name(.)}">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:value-of select="."/>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>

