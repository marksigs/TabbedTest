<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/FindIntroducersRequest.xslt $
Workfile:             $Workfile: FindIntroducersRequest.xslt $
Current Version   $Revision: 9 $
Last Modified      $Modtime: 20/04/07 18:18 $
Modified By        $Author: Mheys $

Copyright				Copyright © 2006 Vertex Financial Services
Description			epsom specific FindIntroducers response transformation

History:

Author		Date				Defect		Description
IK				23/10/2006	EP2_21	created
IK				08/11/2006	EP2_35	add FSAONLYINDIVIDUAL search
IK				14/11/2006	EP2_89	add PACKAGERINDICATOR search
IK				14/11/2006	EP2_89	add PACKAGINGASSOCIATIONINDICATOR search
IK				31/01/2007	EP2_1152	default USERAUTHORITYLEVEL on REQUEST
PSC			08/03/2007	EP2_1190	add new parameters for ARFIRM search
IK				20/04/2007	EP2_2504 no FIRMPERMISSIONS for PRINCIPALFIRM where PACKAGERINDICATOR '1'
									add explicit child node FIRMCLUBNETWORKASSOCIATION only
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:msg="http://Request.SubmitAiP.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
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
				<xsl:attribute name="OPERATION">FindIntroducers</xsl:attribute>
				<xsl:attribute name="omigaClient">epsom</xsl:attribute>
				<xsl:attribute name="RETURNREQUEST">Y</xsl:attribute>
				<xsl:attribute name="postProcRef">FindIntroducersResponse.xslt</xsl:attribute>
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:apply-templates select="INTRODUCERLIST"/>
				<xsl:for-each select="LENGTHSEARCH">
					<xsl:copy-of select="."/>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="INTRODUCERLIST">
		<xsl:if test="@BROKERINDICATOR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">PRINCIPALFIRM</xsl:attribute>
				<xsl:element name="PRINCIPALFIRM">
					<xsl:attribute name="PACKAGERIND">0</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@ARFIRMINDICATOR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">ARFIRM</xsl:attribute>
				<xsl:element name="ARFIRM">
					<xsl:call-template name="targets"/>
					<xsl:call-template name="networkParameters"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@MORTGAGECLUBINDICATOR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">MORTGAGECLUBNETWORKASSOCIATION</xsl:attribute>
				<xsl:element name="MORTGAGECLUBNETWORKASSOCIATION">
					<xsl:attribute name="PACKAGERIND">0</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@PACKAGINGASSOCIATIONINDICATOR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">MORTGAGECLUBNETWORKASSOCIATION</xsl:attribute>
				<xsl:element name="MORTGAGECLUBNETWORKASSOCIATION">
					<xsl:attribute name="PACKAGERIND">1</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@OMIGAINDIVIDUALBROKERAR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">INTRODUCERDETAILS</xsl:attribute>
				<!-- switch off for now
				<xsl:attribute name="COMBOLOOKUP">Y</xsl:attribute> -->
				<xsl:element name="INTRODUCERDETAILS">
					<xsl:attribute name="BROKERTYPE">AR</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@OMIGAINDIVIDUALBROKERDA='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">INTRODUCERDETAILS</xsl:attribute>
				<!-- switch off for now
				<xsl:attribute name="COMBOLOOKUP">Y</xsl:attribute> -->
				<xsl:element name="INTRODUCERDETAILS">
					<xsl:attribute name="BROKERTYPE">DA</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@OMIGAINDIVIDUALPACKAGER='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">INTRODUCERDETAILS</xsl:attribute>
				<!-- switch off for now
				<xsl:attribute name="COMBOLOOKUP">Y</xsl:attribute> -->
				<xsl:element name="INTRODUCERDETAILS">
					<xsl:attribute name="BROKERTYPE">PA</xsl:attribute>
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@FSAONLYINDIVIDUAL='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">FSAINDIVIDUAL</xsl:attribute>
				<!-- switch off for now
				<xsl:attribute name="COMBOLOOKUP">Y</xsl:attribute> -->
				<xsl:element name="FSAINDIVIDUAL">
					<xsl:call-template name="targets"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@PACKAGERINDICATOR='1'">
			<xsl:element name="OPERATION">
				<xsl:attribute name="SCHEMA_NAME">FindIntroducersWSSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">READ</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">PRINCIPALFIRM</xsl:attribute>
				<xsl:element name="PRINCIPALFIRM">
					<xsl:attribute name="PACKAGERIND">1</xsl:attribute>
					<xsl:call-template name="targets"/>
					<!-- EP2_2504 -->
					<xsl:element name="FIRMCLUBNETWORKASSOCIATION"/>
				</xsl:element>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template name="targets">
		<xsl:for-each select="../TARGETSEARCH/@*">
			<xsl:copy-of select="."/>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="networkParameters">
		<xsl:copy-of select="../LENGTHSEARCH/ARFIRMLENGTH/@*[name()='ALLPARENTPRINCIPALS' or name()='MORTGAGENETWORKS' or name()='SELECTEDNETWORK']"/>
	</xsl:template>
</xsl:stylesheet>
