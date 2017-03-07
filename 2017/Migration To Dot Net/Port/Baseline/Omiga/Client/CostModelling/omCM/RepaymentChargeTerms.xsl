<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/">
		<xsl:apply-templates select="Control/Mortgage"/>
	</xsl:template>
	<xsl:template match="Mortgage">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="RepaymentChargeTerms">
				<!-- Sort and deduplicate all the EarlyRepaymentCharge Durations -->
				<xsl:for-each select="ParameterOverride[@Parameter='EarlyRepaymentCharge']/ParameterOverrideValue[not(@Duration=preceding::ParameterOverride[@Parameter='EarlyRepaymentCharge']/ParameterOverrideValue/@Duration)]">
				<xsl:sort data-type="number" select="number(@Duration)"/>
					<xsl:value-of select="@Duration"/><xsl:text> </xsl:text>
				</xsl:for-each>
			</xsl:attribute>

			<xsl:for-each select="ElementGroup">
				<xsl:element name="ElementGroup">
					<xsl:variable name="GroupNum" select="substring-after(@Id, 'ElementGroup')"/>
					<xsl:attribute name="Id"><xsl:value-of select="@Id"/></xsl:attribute>
					<!-- Find all the ParameterOverrideIds that need to be linked to the current ElementGroup -->
					<xsl:attribute name="ParameterOverrideIds">
						<xsl:for-each select="../ParameterOverride[substring-before(substring-after(@Id, 'EG'), '-PO') = $GroupNum]">
							<xsl:value-of select="@Id"/><xsl:text> </xsl:text>
						</xsl:for-each>
					</xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
