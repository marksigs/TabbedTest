<?xml version="1.0" encoding="UTF-8"?>
<!-- GHun	26/04/2007	EP2_2142 Created -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>

	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="APPLICATION">
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{', RESPONSE/APPLICATION/@APPLICATIONNUMBER, '}')"/></xsl:attribute>
			</xsl:element>
			<xsl:element name="BROKER">
				<xsl:attribute name="NAME">
					<xsl:choose>
						<xsl:when test="string(RESPONSE/BROKER/@FIRMNAME)"><xsl:value-of select="concat(RESPONSE/BROKER/@NAME, ' of ', RESPONSE/BROKER/@FIRMNAME)"/></xsl:when>
						<xsl:otherwise><xsl:value-of select="RESPONSE/BROKER/@NAME"/></xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:element>
			<xsl:element name="WEBSITE">
				<xsl:attribute name="URL"><xsl:value-of select="/RESPONSE/GLOBALPARAMETER/@STRING"/></xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>

</xsl:stylesheet>
