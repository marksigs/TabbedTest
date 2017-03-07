<?xml version="1.0" encoding="UTF-8"?>
<!-- IK - 19/09/2006 - EP2_1 standard Omiga.vertex.co.uk namespace identifier  -->
<!--IK 31/01/2007 EP2_1152 default USERAUTHORITYLEVEL on REQUEST -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.GetApplicationData.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:if test="@USERID">
				<xsl:attribute name="USERID"><xsl:value-of select="@USERID"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="@UNITID">
				<xsl:attribute name="UNITID"><xsl:value-of select="@UNITID"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="@CHANNELID">
				<xsl:attribute name="CHANNELID"><xsl:value-of select="@CHANNELID"/></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
			<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">AIPINGESTION</xsl:attribute>
			<xsl:attribute name="postProcRef">transformGetApplicationDataResponse.xslt</xsl:attribute>
			<!-- EP2_1152 -->
			<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
			<xsl:for-each select="msg:APPLICATIONDETAILS">
				<xsl:element name="APPLICATION">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="@APPLICATIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
