<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\code\WebServiceInterface\GetNextMessageSequenceNumberRequest.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.GetNextMessageSequenceNumber.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
			<xsl:attribute name="CRUD_OP">UPDATE</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">FIRSTTITLEPARAMETER</xsl:attribute>
			<xsl:element name="FIRSTTITLEPARAMETER">
				<xsl:attribute name="BOOLEAN">1</xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
