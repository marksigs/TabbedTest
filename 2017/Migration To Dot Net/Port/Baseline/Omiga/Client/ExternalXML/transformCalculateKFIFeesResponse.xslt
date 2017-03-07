<?xml version="1.0" encoding="UTF-8"?>
<!-- MAR1166 GHun -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>
	
		<xsl:template match="ONEOFFCOSTLIST">
		<xsl:element name="FEELIST">
			<xsl:apply-templates select="ONEOFFCOST"/>
		</xsl:element>
	</xsl:template>

	<xsl:template match="ONEOFFCOST">
		<xsl:element name="FEE">
			<xsl:copy-of select="@*[name()='IDENTIFIER' or name()='DESCRIPTION' or name()='AMOUNT' or name()='CANBEADDEDTOLOAN' or name()='CANBEREFUNDED']"></xsl:copy-of>
		</xsl:element>
	</xsl:template>

</xsl:stylesheet>
