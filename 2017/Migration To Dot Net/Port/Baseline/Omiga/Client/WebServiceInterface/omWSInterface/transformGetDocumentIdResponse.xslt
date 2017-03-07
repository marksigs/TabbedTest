<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\xml\getDocumentIdGuidResponse.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:choose>
				<xsl:when test="@TYPE='SUCCESS'">
					<xsl:choose>
						<xsl:when test="DOCUMENTID[@DOCUMENTGUID]">
							<xsl:attribute name="TYPE"><xsl:value-of select="@TYPE"/></xsl:attribute>
							<xsl:element name="APPLICATION">
								<xsl:for-each select="DOCUMENTID">
									<xsl:for-each select="@*">
										<xsl:copy/>
									</xsl:for-each>
								</xsl:for-each>
							</xsl:element>
						</xsl:when>
						<xsl:when test="DOCUMENTID[@TASKPENDING]">
							<xsl:attribute name="TYPE">WARNING</xsl:attribute>
							<xsl:element name="MESSAGE">
								<xsl:element name="MESSAGETYPE">warning</xsl:element>
								<xsl:element name="MESSAGETEXT">Document not yet produced, please try later</xsl:element>
							</xsl:element>
						</xsl:when>
						<xsl:when test="DOCUMENTID[@TASKEXPIRED]">
							<xsl:attribute name="TYPE">WARNING</xsl:attribute>
							<xsl:element name="MESSAGE">
								<xsl:element name="MESSAGETYPE">error</xsl:element>
								<xsl:element name="MESSAGETEXT">Error producing document (task expired)</xsl:element>
							</xsl:element>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="TYPE">WARNING</xsl:attribute>
							<xsl:element name="MESSAGE">
								<xsl:element name="MESSAGETYPE">error</xsl:element>
								<xsl:element name="MESSAGETEXT">Error producing document (no document task)</xsl:element>
							</xsl:element>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:apply-templates/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>
