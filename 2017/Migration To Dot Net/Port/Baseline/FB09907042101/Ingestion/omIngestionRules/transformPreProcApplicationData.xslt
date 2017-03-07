<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:			$Workfile: POSSS020.xslt $
Copyright:			Copyright © 2006 Vertex Financial Services
Current Version:	$Revision: 4 $
Last Modified:		$JustDate:  8/02/06 $
Modified By:		$Author: Ikemp $
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Description: transformPreProcApplicationData

History:

Author	Date       		Description
IK			31/01/2006 	MAR1124 used by omIngestionRules.PreProcBO.CompareData
IK			08/02/2006	MAR1238 drop USERHISTORY

================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:for-each select="@*"><xsl:copy-of select="."/></xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="ACCOUNTRELATIONSHIP">
		<xsl:variable name="thisAccountGuid" select="@ACCOUNTGUID"/>
		<xsl:variable name="thisRoleType" select="../../../@CUSTOMERROLETYPE"/>
		<xsl:variable name="thisCustOrder" select="../../../@CUSTOMERORDER"/>
		<xsl:if test="not(//RESPONSE/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE[(@CUSTOMERROLETYPE = $thisRoleType) and (@CUSTOMERORDER &lt; $thisCustOrder)]/CUSTOMER/CUSTOMERVERSION/ACCOUNTRELATIONSHIP[@ACCOUNTGUID= $thisAccountGuid])">
			<xsl:copy>
				<xsl:for-each select="@*"><xsl:copy-of select="."/></xsl:for-each>
				<xsl:apply-templates select="*"/>
			</xsl:copy>
		</xsl:if>
	</xsl:template>
	<xsl:template match="OTHERCUSTOMERACCOUNTRELATIONSHIP">
		<xsl:variable name="thisCust" select="ancestor::CUSTOMER/@CUSTOMERNUMBER"/>
		<xsl:if test="@CUSTOMERNUMBER != $thisCust">
			<xsl:copy-of select="."/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="USERHISTORY"/>
</xsl:stylesheet>
