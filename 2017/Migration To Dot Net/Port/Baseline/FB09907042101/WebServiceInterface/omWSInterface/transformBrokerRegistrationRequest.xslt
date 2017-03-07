<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformBrokerRegistrationRequest.xslt $
Workfile:             $Workfile: transformBrokerRegistrationRequest.xslt $
Current Version   $Revision: 6 $
Last Modified      $Modtime: 10/04/07 14:04 $
Modified By        $Author: Mheys $

Copyright			Copyright © 2006 Vertex Financial Services
Description			

History:

Author		Date				    Description
SR			16/10/2006		EP2_11  - New 
SR			19/10/2006		EP2_11 No postProcRef required. This is handled in omWSInterfaceBO.
SR			31/10/2006		EP2_26	spelling correction
SR			24/11/2006		EP2_169  add Organisation.UserSurname and UserForeName to Introducer 
IK				31/01/2007		EP2_1152	default USERAUTHORITYLEVEL on REQUEST
IK				23/03/2007		EP2_1190	additional usp_CreateIntroducerFirm parameters
SR			10/04/2007		EP2_2270 add attrib TRADINGAS to INTRODUCERFIRM
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="http://vertex.com/omiga4">
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
		<xsl:for-each select="msxsl:node-set($firstPass)/BROKERREGISTER_REQUEST">
			<xsl:element name="REQUEST">
				<xsl:for-each select="@*">
					<xsl:copy-of select="."/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">CREATE</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">BROKER_REGISTER</xsl:attribute>
				<xsl:apply-templates/>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="INTRODUCER">
		<xsl:element name="INTRODUCER">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:attribute name="USERSURNAME"><xsl:value-of select="ORGANISATIONUSER/@USERSURNAME"/></xsl:attribute>
			<xsl:attribute name="USERFORENAME"><xsl:value-of select="ORGANISATIONUSER/@USERFORENAME"/></xsl:attribute>
			<xsl:attribute name="FSAREF">
			<xsl:choose>
				<xsl:when test="@ARINDICATOR='1'"><xsl:value-of select="INTRODUCERFIRM/ARFIRM/@FSAARFIRMREF"/></xsl:when>
				<xsl:otherwise><xsl:value-of select="INTRODUCERFIRM/PRINCIPALFIRM/@FSAREF"/></xsl:otherwise>
			</xsl:choose>
			</xsl:attribute>
			<xsl:element name="INTRODUCERFIRM">
				<xsl:if test="@ARINDICATOR='1'">
					<xsl:attribute name="PRINCIPALFIRMID"><xsl:value-of select="INTRODUCERFIRM/PRINCIPALFIRM/@PRINCIPALFIRMID"/></xsl:attribute>
					<xsl:attribute name="ARFIRMID"><xsl:value-of select="INTRODUCERFIRM/ARFIRM/@ARFIRMID"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="INTRODUCERFIRM/@TRADINGAS">
					<xsl:attribute name="TRADINGAS"><xsl:value-of select="INTRODUCERFIRM/@TRADINGAS"/></xsl:attribute>
				</xsl:if>
			</xsl:element>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ORGANISATIONUSER">
		<xsl:element name="ORGANISATIONUSER">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
