<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\code\WebServicesTest\FindBusinessForCustomerWS\ResponseFiles\Untitled27.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.FindBusinessForCustomer.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="tpRequestChild"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="tpRequestChild">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:value-of select="."/>
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
