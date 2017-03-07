<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:element name="OMIGAUSERLIST">
				<xsl:for-each select="OMIGAUSER">
					<xsl:element name="OMIGAUSER">
						<xsl:element name="USERID"><xsl:value-of select="@USERID"/></xsl:element>
						<xsl:element name="ACCESSTYPE">
							<xsl:attribute name="TEXT"><xsl:value-of select="ORGANISATIONUSER/cv/@valuename"/>	</xsl:attribute>
							<xsl:value-of select="@ACCESSTYPE"/>
						</xsl:element>
						<xsl:element name="OMIGAUSERACTIVEFROM">
							<xsl:call-template name="formatDate">
								<xsl:with-param name="dateString">
									<xsl:value-of select="@OMIGAUSERACTIVEFROM"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:element>
						<xsl:element name="OMIGAUSERACTIVETO">
							<xsl:call-template name="formatDate">
								<xsl:with-param name="dateString">
									<xsl:value-of select="@OMIGAUSERACTIVETO"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:element>
						<xsl:element name="NOTES"><xsl:value-of select="@NOTES"/></xsl:element>
						<xsl:element name="CHANGEPASSWORDINDICATOR"><xsl:value-of select="@CHANGEPASSWORDINDICATOR"/></xsl:element>
						<xsl:element name="WORKINGHOURTYPE"><xsl:value-of select="@WORKINGHOURTYPE"/></xsl:element>
						<xsl:element name="CREDITCHECKACCESS"><xsl:value-of select="@CREDITCHECKACCESS"/></xsl:element>
						<xsl:element name="WORKGROUPUSER"><xsl:value-of select="@CREDITCHECKACCESS"/></xsl:element>
						<xsl:element name="USERFORENAME"><xsl:value-of select="ORGANISATIONUSER/@USERFORENAME"/></xsl:element>
						<xsl:element name="USERSURNAME"><xsl:value-of select="ORGANISATIONUSER/@USERSURNAME"/></xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="formatDate">
		<xsl:param name="dateString"/>
		<xsl:if test="string-length($dateString) != 0">
			<xsl:value-of select="concat(substring($dateString,9,2),'/',substring($dateString,6,2),'/',substring($dateString,1,4))"/>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>
