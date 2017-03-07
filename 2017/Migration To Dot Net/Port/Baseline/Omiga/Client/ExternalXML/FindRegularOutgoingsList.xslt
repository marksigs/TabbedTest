<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
		<xsl:template match="REGULAROUTGOINGSLIST">
			<xsl:element name="REGULAROUTGOINGSLIST">
				<xsl:apply-templates select="REGULAROUTGOINGS">
				   <xsl:sort select="REGULAROUTGOINGSTYPE"/>
				</xsl:apply-templates>
			</xsl:element>
		</xsl:template>
	      <xsl:template match="REGULAROUTGOINGS">
			<xsl:copy-of select='.'/>
		</xsl:template>	
</xsl:stylesheet>
