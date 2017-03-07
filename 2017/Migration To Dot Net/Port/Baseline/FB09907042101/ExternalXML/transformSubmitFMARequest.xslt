<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Archive                  $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformSubmitFMARequest.xslt $
Workfile                $Workfile: transformSubmitFMARequest.xslt $
Current Version   	$Revision: 5 $
Last Modified       	$Modtime: 1/02/07 10:08 $
Modified By          	$Author: Dbarraclough $

Prog					Date				Description
Peter Edney		10/03/2006	MAR1357 - Addedd LOANREPAYMENTINDICATOR attribute to LOANSLIABILITIES node.
Helen Aldred   	15/03/2006	MAR1385 - Added CustomerRoleType to AccountRelationship.
IK						21/09/2006	EP2_1 - 'product' namespace identifiers
IK						18/11/2006	EP2_134 omiga client identifier
IK						23/11/2006	EP2_159 - as EP2_134 (StopAndSaveFMA)
IK						09/01/2007	EP2_352 - remote introducer, always CREATE USERHISTORY
IK						09/01/2007	EP2_702 - create ACCOUNTRELATIONSHIP records, use imported common functionality
IK						31/01/2007	EP2_1152	default USERAUTHORITYLEVEL on REQUEST
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
	<xsl:import href="importFMARequest.xslt"/>
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="firstPass">
		<xsl:for-each select="*">
			<xsl:call-template name="stripNS"/>
		</xsl:for-each>
	</xsl:variable>
	<xsl:template name="stripNS">
		<xsl:element name="{local-name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="stripNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/">
		<xsl:for-each select="msxsl:node-set($firstPass)/REQUEST">
			<xsl:element name="REQUEST">
				<xsl:attribute name="OPERATION">SubmitFMA</xsl:attribute>
				<xsl:attribute name="USERID">epsom</xsl:attribute>
				<xsl:attribute name="UNITID">epsom</xsl:attribute>
				<xsl:attribute name="omigaClient">epsom</xsl:attribute>
				<xsl:attribute name="CRUD_OP">IUPDATE</xsl:attribute>
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:apply-templates select="*"/>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>
