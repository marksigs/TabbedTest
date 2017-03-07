<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformGetBrokerProfileResponse.xslt $
Workfile:             $Workfile: transformGetBrokerProfileResponse.xslt $
Current Version   $Revision: 2 $
Last Modified      $Modtime: 23/01/07 9:34 $
Modified By        $Author: Dbarraclough $

Copyright			Copyright © 2006 Vertex Financial Services
Description			

History:

Author		Date				    Description
SR			16/10/2006			EP2_11 - New 
SR			19/10/2006			EP2_11
SR			19/10/2006			EP2_11
SR			21/01/2007		 	EP2_328 - Add FirmPermissions to Response
================================================================================-->
<xsl:stylesheet version="2.0" xmlns="http://Response.BrokerRegAndMaintain.Omiga.vertex.co.uk" xmlns:IntroProf="http://IntroducerDetail.Omiga.vertex.co.uk" xmlns:msgdt="http://msgtypes.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="MESSAGE"/>
			<xsl:apply-templates select="ERROR"/>
			<xsl:apply-templates select="INTRODUCER"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MESSAGE">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="MESSAGE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ERROR">
		<xsl:element name="ERROR">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="msgdtNS">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="{local-name(.)}">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:value-of select="."/>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="INTRODUCER">
		<xsl:element name="INTRODUCER">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="INTRODUCERFIRM"/>
			<xsl:apply-templates select="ORGANISATIONUSER"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="INTRODUCERFIRM">
		<xsl:element name="INTRODUCERFIRM" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="PRINCIPALFIRM"/>
			<xsl:apply-templates select="ARFIRM"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="PRINCIPALFIRM">	
		<xsl:element name="PRINCIPALFIRM" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="FIRMPERMISSIONS"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ARFIRM">	
		<xsl:element name="ARFIRM" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="FIRMPERMISSIONS"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="FIRMPERMISSIONS">
		<xsl:element name="FIRMPERMISSIONS" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>	
	<xsl:template match="ORGANISATIONUSER">
		<xsl:element name="ORGANISATIONUSER" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="PASSWORD"/>
			<xsl:apply-templates select="ORGANISATIONUSERCONTACTDETAILS" />
		</xsl:element>
	</xsl:template>
	<xsl:template match="PASSWORD">
		<xsl:element name="PASSWORD" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ORGANISATIONUSERCONTACTDETAILS">
		<xsl:element name="ORGANISATIONUSERCONTACTDETAILS" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="CONTACTDETAILS"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CONTACTDETAILS">
		<xsl:element name="CONTACTDETAILS" namespace="http://IntroducerDetail.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="CONTACTTELEPHONEDETAILS"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CONTACTTELEPHONEDETAILS">
		<xsl:element name="CONTACTTELEPHONEDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>		
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
