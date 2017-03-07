<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/">
		<xsl:element name="CONTACTHISTORYLIST">
			<xsl:for-each select="CONTACTHISTORYLIST/CONTACTHISTORY">
			<xsl:sort order="ascending" select="USERNAME "/>
				<xsl:element name="CONTACTHISTORY">
					<xsl:element name="CUSTOMERNUMBER">
						<xsl:value-of select="CUSTOMERNUMBER"/>
					</xsl:element>
					<xsl:element name="CONTACTHISTORYDATETIME">
						<xsl:value-of select="CONTACTHISTORYDATETIME"/>
					</xsl:element>
					<xsl:element name="CONTACTREASONCODE">
						<xsl:attribute name="TEXT">
							<xsl:value-of select="CONTACTREASONCODE/@TEXT"/>
						</xsl:attribute>
						<xsl:value-of select="CONTACTREASONCODE"/>
					</xsl:element>
					<xsl:element name="OTHERREASONTEXT">
						<xsl:value-of select="OTHERREASONTEXT"/>
					</xsl:element>
					<xsl:element name="CONTACTTEXT">
						<xsl:value-of select="CONTACTTEXT"/>
					</xsl:element>
					<xsl:element name="USERID">
						<xsl:value-of select="USERID"/>
					</xsl:element>
					<xsl:element name="USERNAME">
						<xsl:value-of select="USERNAME"/>
					</xsl:element>
					<xsl:element name="UNITID">
						<xsl:value-of select="UNITID"/>
					</xsl:element>
					<xsl:element name="UNITNAME">
						<xsl:value-of select="UNITNAME"/>
					</xsl:element>
					<xsl:element name="STATUSINDICATOR">
						<xsl:value-of select="STATUSINDICATOR"/>
					</xsl:element>	
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
