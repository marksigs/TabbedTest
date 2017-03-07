<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
		<xsl:template match='/'>
			<xsl:element name="BANKRUPTCYHISTORYLIST">
				<xsl:for-each select='.//BANKRUPTCYHISTORY'>
					<xsl:sort order='descending' select='concat(substring(DATEDECLARED, 7,4), substring(DATEDECLARED, 4,2), substring(DATEDECLARED, 1,2))'/>
					<xsl:copy-of select='.'/>
				</xsl:for-each>
			</xsl:element>
		</xsl:template>
</xsl:stylesheet>

