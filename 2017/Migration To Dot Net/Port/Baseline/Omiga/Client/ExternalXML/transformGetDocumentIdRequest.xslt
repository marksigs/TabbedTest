<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\xml\getDocumentIdRequest.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  xmlns:msg="http://Request.GetDocumentID.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
			<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">DOCUMENTID</xsl:attribute>
			<xsl:attribute name="postProcRef">transformGetDocumentIdResponse.xslt</xsl:attribute>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="msg:APPLICATION">
		<xsl:element name="DOCUMENTID">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
