<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformUpdateBrokerProfileRequest.xslt $
Workfile:             $Workfile: transformUpdateBrokerProfileRequest.xslt $
Current Version   $Revision: 3 $
Last Modified      $Modtime: 1/02/07 10:08 $
Modified By        $Author: Dbarraclough $

Copyright			Copyright © 2006 Vertex Financial Services
Description			

History:

Author		Date				    Description
SR			16/10/2006			EP2_11  - New 
SR			19/10/2006			EP2_11  - remove attrib 'postProcRef'. This is handled in omWSInterfaceBO
SR			31/10/2006			EP2_26 - spelling correction. Handle creation of new INTRODUCERFIRM while updating Broker.
SR			31/10/2006			EP2_26 - add CRUD_OP attrib with value NULL to nodes PRINCIPALFIRM and ARFIRM so that these will not be created when creating
											new INTRODUCERFIRM records
IK				31/01/2007		EP2_1152	default USERAUTHORITYLEVEL on REQUEST
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
		<xsl:for-each select="msxsl:node-set($firstPass)/UPDATEBROKERPROFILE_REQUEST">
			<xsl:element name="REQUEST">
				<xsl:for-each select="@*">
					<xsl:copy-of select="."/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
				<xsl:attribute name="CRUD_OP">IUPDATE</xsl:attribute>
				<xsl:attribute name="ENTITY_REF">BROKER_PROFILE</xsl:attribute>
				<xsl:for-each select="*">
					<xsl:call-template name="showElem"/>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="showElem">
		<xsl:choose>
			<xsl:when test="local-name()='ARFIRM'">
				<xsl:element name="ARFIRM">
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:attribute name="FSAREF"><xsl:value-of select="@FSAARFIRMREF"/></xsl:attribute>
					<xsl:attribute name="CRUD_OP">NULL</xsl:attribute>
				</xsl:element>
			</xsl:when>
			<xsl:when test="local-name()='PRINCIPALFIRM'">
				<xsl:element name="PRINCIPALFIRM">
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:attribute name="CRUD_OP">NULL</xsl:attribute>
				</xsl:element>
			</xsl:when>			
			<xsl:otherwise>
				<xsl:element name="{local-name()}">
					<xsl:for-each select="@*">
						<xsl:copy/>
					</xsl:for-each>
					<xsl:for-each select="*">
						<xsl:call-template name="showElem"/>
					</xsl:for-each>
				</xsl:element>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
