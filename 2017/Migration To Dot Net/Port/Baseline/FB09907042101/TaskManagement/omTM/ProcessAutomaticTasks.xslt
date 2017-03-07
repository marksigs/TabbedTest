<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="RESPONSE/CASESTAGE">
		<xsl:element name="CASETASKLIST">
			<xsl:for-each select="CASETASK[@AUTOMATICTASKIND='1']">
			<xsl:sort select="@TASKID"/>
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>

