<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="node()|@*">
		<xsl:copy>
			<xsl:apply-templates select="node()|@*"/>
		</xsl:copy>
	</xsl:template>
	
	<xsl:template match="CUSTOMERADDRESSLIST">
		<xsl:apply-templates select="CUSTOMERADDRESS"/>
	</xsl:template>
	
	<xsl:template match="CUSTOMERADDRESS">
		<xsl:copy>
			<xsl:apply-templates select="@*"/>
				<xsl:apply-templates select="ADDRESS/@*"/>
			</xsl:copy>
	</xsl:template>
		
	<xsl:template match="CUSTOMER">
		<xsl:copy>
			<xsl:apply-templates select="@*"/>
				<xsl:apply-templates select="CUSTOMERVERSION/@*"/>
				<xsl:apply-templates select="CUSTOMERVERSION/*"/>
		</xsl:copy>
	</xsl:template>
	
	<xsl:template match="CUSTOMERTELEPHONENUMBERLIST|AREASOFINTERESTLIST|@DATEMOVEDIN|@MARITALSTATUS|@MARITALSTATUS_TEXT|@NATUREOFOCCUPANCY|@NATUREOFOCCUPANCY_TEXT|@NUMBEROFDEPENDANTS|@CUSTOMERSTATUS|@MEMBEROFSTAFF|@TIMEATBANKYEARS|@TIMEATBANKMONTHS">
	</xsl:template>
	
</xsl:stylesheet>
