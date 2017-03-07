<?xml version="1.0" encoding="UTF-8"?>
<!-- GHun	26/04/2007  EP2_2142 Created -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>

	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="APPLICATION">
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{', /RESPONSE/EMAILDIPRESPONSE/@APPLICATIONNUMBER, '}')"/></xsl:attribute>
				<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="/RESPONSE/EMAILDIPRESPONSE/@CORRESPONDENCESALUTATION"/></xsl:attribute>
			</xsl:element>
			<xsl:element name="WEBSITE">
				<xsl:attribute name="URL"><xsl:value-of select="/RESPONSE/EMAILDIPRESPONSE/GLOBALPARAMETER/@STRING"/></xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>

</xsl:stylesheet>
