<?xml version="1.0" encoding="UTF-8"?>
<!-- SR   19/09/2006      EP2_1 : Modified namespaces for Epsom -->
<!-- EP2_1152 IK 31/01/2007	default USERAUTHORITYLEVEL on REQUEST -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.GetCaseTrackingData.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<!-- EP2_1152 -->
			<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
			<xsl:attribute name="OPERATION">GetCaseTrackingAllData</xsl:attribute>
			<xsl:for-each select="msg:OPI/msg:APPLICATION">
				<xsl:element name="APPLICATION">
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
