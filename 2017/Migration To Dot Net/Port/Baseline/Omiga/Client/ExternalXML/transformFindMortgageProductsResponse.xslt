<?xml version="1.0" encoding="UTF-8"?>
<!--SR 19/09/2006 EP2_1 modified namespaces for Epsom -->
<!--SR 23/10/2006 EP2_1 modified to check in -->
<?altova_samplexml C:\omiga4Projects\projectMars\xml\FindMortgageProducts\rawResponse.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="*">
		<xsl:choose>
			<xsl:when test="RESPONSE">
					<xsl:apply-templates/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:copy>
					<xsl:apply-templates select="@*"/>
					<xsl:apply-templates/>
				</xsl:copy>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="@*">
		<xsl:choose>
			<xsl:when test="name() ='RAW'"/>
			<xsl:when test="name() ='TEXT'"/>
			<xsl:otherwise>
				<xsl:copy/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
