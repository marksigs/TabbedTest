<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/">
		<xsl:element name="CASETASKS">
				<xsl:for-each select="RESPONSE/CASETASK[not(@CASEID=preceding::CASETASK/@CASEID)]">
					<xsl:element name="CASETASK">
						<xsl:attribute name="CASEID">
							<xsl:value-of select="@CASEID"/>
						</xsl:attribute>
						<xsl:attribute name="STAGEID">
							<xsl:value-of select="@STAGEID"/>
						</xsl:attribute>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
	</xsl:template>
</xsl:stylesheet>
