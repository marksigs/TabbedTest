<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Archive                  $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformSubmitFMAResponse.xslt $
Workfile                 $Workfile: transformSubmitFMAResponse.xslt $
Current Version   	$Revision: 2 $
Last Modified       	$Modtime: 26/03/07 13:50 $
Modified By          	$Author: Lesliem $

Prog					Date					Description
IK						23/11/2006		EP2_159 created
IK						21/03/2007		EP2_1062 application form response
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0"  xmlns="http://Response.SubmitFMA.Omiga.vertex.co.uk" xmlns:msgdt="http://msgtypes.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE" namespace="http://Response.SubmitFMA.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="DOCUMENTCONTENTS ">
					<xsl:element name="DOCUMENTCONTENTS" namespace="http://Response.SubmitFMA.Omiga.vertex.co.uk">
						<xsl:copy-of select="@*"/>
					</xsl:element>
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

