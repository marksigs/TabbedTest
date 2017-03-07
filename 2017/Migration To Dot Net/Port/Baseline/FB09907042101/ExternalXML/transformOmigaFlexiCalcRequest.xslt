<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:om="http://Request.FlexiCalc.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="om:OmigaFlexiCalcRequest">
		<xsl:for-each select="*">
			<xsl:call-template name="stripNS"/>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="stripNS">
		<xsl:element name="{local-name()}">
			<xsl:copy-of select="@*"/>
			<xsl:for-each select="*">
				<xsl:call-template name="stripNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
