<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\xml\robH\GetDocumentListResponse.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="DOCUMENTDETAILS">
		<xsl:element name="DOCUMENTDETAILS">
			<xsl:for-each select="@FILEGUID">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@FILENETIMAGEREF">
				<xsl:attribute name="DOCUMENTGUID">
					<xsl:value-of select="." />
				</xsl:attribute>
			</xsl:for-each>
			<xsl:for-each select="@DOCUMENTVERSION">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@DOCUMENTNAME">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@DOCUMENTDESCRIPTION">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@HOSTTEMPLATEID">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@STAGEID">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@USERNAME">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@UNITNAME">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@PRINTLOCATION">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@CUSTOMERNAME">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@RECIPIENTNAME">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@EVENTDATE">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@SOURCESYSTEM">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@DOCUMENTGROUP">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@TEMPLATEID">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@LANGUAGE">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@SEARCHKEY1">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@SEARCHKEY2">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@SEARCHKEY3">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@EVENTKEY">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@ARCHIVEDATE">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@UNITID">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@CREATIONDATE">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@PRINTDATE">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@APPLICATIONNUMBER">
				<xsl:copy/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
