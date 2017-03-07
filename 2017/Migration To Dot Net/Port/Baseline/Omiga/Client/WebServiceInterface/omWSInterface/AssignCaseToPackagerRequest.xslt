<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/AssignCaseToPackagerRequest.xslt $
Workfile:             $Workfile: AssignCaseToPackagerRequest.xslt $
Current Version   $ $
Last Modified      $Modtime: 1/02/07 10:08 $
Modified By        $Author: Dbarraclough $

Copyright			Copyright © 2006 Vertex Financial Services
Description			Assign case to packager request transformation

History:

Author		Date				Defect		Description
PSC			12/12/2006	EP2_434		Created
IK				31/01/2007	EP2_1152	default USERAUTHORITYLEVEL on REQUEST
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:msg="http://Request.SubmitAiP.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	
	<xsl:template match="*">
		<xsl:element name="{local-name()}">
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="node()"/>
		</xsl:element>			
	</xsl:template>
	
	<xsl:template match="*[local-name(.)='REQUEST']">
		<xsl:element name="REQUEST">
			<xsl:copy-of select="@*"/>
			<xsl:attribute name="CRUD_OP">CREATE</xsl:attribute>
			<xsl:attribute name="ENTITY_NAME">ASSIGNCASETOPACKAGER</xsl:attribute>
			<xsl:attribute name="SCHEMA_NAME">WebServiceSchema</xsl:attribute>
			<!-- EP2_1152 -->
			<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
			<xsl:apply-templates select="node()"/>
		</xsl:element>
</xsl:template>
	
	<xsl:template match="*[local-name(.)='ASSIGNCASE']">
		<xsl:element name="ASSIGNCASETOPACKAGER">
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="node()"/>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
