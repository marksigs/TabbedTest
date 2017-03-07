<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="Response">
		<xsl:call-template name="addNS"/>
	</xsl:template>
	<xsl:template name="addNS">
		<xsl:element name="{local-name()}" namespace="http://Alpha.FlexiCalc.Response.vertex.co.uk">
			<xsl:copy-of select="@*"/>
			<xsl:for-each select="*">
				<xsl:call-template name="addNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
