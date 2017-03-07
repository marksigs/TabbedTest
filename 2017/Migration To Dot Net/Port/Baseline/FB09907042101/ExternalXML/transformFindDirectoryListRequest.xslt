<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformFindDirectoryListRequest.xslt $
Workfile:             $Workfile: transformFindDirectoryListRequest.xslt $
Current Version   $Revision: 2 $
Last Modified      $Modtime: 1/02/07 10:08 $
Modified By        $Author: Dbarraclough $

Copyright				Copyright © 2007 Vertex Financial Services
Description			epsom specific FindDirectoryList response transformation

History:

Author		Date				Defect		Description
IK				31/01/2007	EP2_1152	default USERAUTHORITYLEVEL on REQUEST
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.FindDirectoryList.Omiga.vertex.co.uk">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<!-- EP2_1152 -->
			<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
			<xsl:for-each select="msg:NAMEANDADDRESSDIRECTORY">
				<xsl:element name="NAMEANDADDRESSDIRECTORY">
					<xsl:for-each select="*">
						<xsl:element name="{name()}">
							<xsl:value-of select="."/>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
