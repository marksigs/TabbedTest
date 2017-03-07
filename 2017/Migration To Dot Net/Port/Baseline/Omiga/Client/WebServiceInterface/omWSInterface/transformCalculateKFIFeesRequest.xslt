<?xml version="1.0" encoding="utf-8"?>
<!-- MAR301 GHun -->
<!-- MAR1061 PEdney -->
<!-- MAR1571 Bill Coates - Extract PURCHASEPRICE from ApplicationFactFind, not PURCHASEPRICEORESTIMATEDVALUE-->
<!-- EP2_1152 IK 31/01/2007	default USERAUTHORITYLEVEL on REQUEST -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/|comment()|processing-instruction()">
		<xsl:copy>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="*">
	  <!-- EP2_1152 -->
	  <xsl:choose>
		<xsl:when test="local-name()='REQUEST'">
				<xsl:element name="REQUEST">
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
					<xsl:apply-templates select="*"/>
				</xsl:element>
		</xsl:when>
		<xsl:otherwise>
			<xsl:element name="{local-name()}">
				<xsl:if test="local-name()='MORTGAGESUBQUOTE'">
					<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE">
						<xsl:value-of select="../../@PURCHASEPRICE"/>
					</xsl:attribute> 
				</xsl:if>
				<xsl:apply-templates select="@*|node()"/>			
		</xsl:element>		
		</xsl:otherwise>
	</xsl:choose>
	</xsl:template>
	<xsl:template match="@*">
		<xsl:attribute name="{local-name()}"><xsl:value-of select="."/></xsl:attribute>
	</xsl:template>
</xsl:stylesheet>
