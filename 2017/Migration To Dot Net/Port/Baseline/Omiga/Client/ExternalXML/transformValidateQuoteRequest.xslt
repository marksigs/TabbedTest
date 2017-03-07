<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\code\WebClientsTest\xml\validateQuoteRequest.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.ValidateQuote.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
			<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">VALIDATEQUOTEDATA</xsl:attribute>
			<xsl:attribute name="postProcRef">validateQuoteData.xslt</xsl:attribute>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="msg:APPLICATION">
		<xsl:element name="VALIDATEQUOTEDATA">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="@ACTIVEQUOTEIND = 'true'">
				<xsl:attribute name="ACTIVEQUOTEIND">1</xsl:attribute>
			</xsl:if>
			<xsl:if test="@ACCEPTEDQUOTEIND = 'true'">
				<xsl:attribute name="ACCEPTEDQUOTEIND">1</xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
