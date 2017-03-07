<?xml version="1.0" encoding="UTF-8"?>
<!-- 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Author		Date			Description
PEdney	12/12/2006	EP2_421 - Created
PEdney	14/12/2006	EP2_480 - Ensure combovlaue items are present even when the list of web combo value items do not match the list of omiga web combo items.
PEdney	15/12/2006	EP2_526 - Add empty validation type to stop Amatica web failing.
PEdney	05/02/2007	EP2_1233 - Modify GetComboList node path.
====================================================================================
  -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--======================================================================================================-->
	<xsl:template match="/">
		<xsl:apply-templates select="/RESPONSE"/>
	</xsl:template>
	<!--======================================================================================================-->
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="COMBOLISTWS/COMBOGROUP"/>
		</xsl:element>
	</xsl:template>
	<!--======================================================================================================-->
	<xsl:template match="COMBOGROUP">
		<COMBOGROUP>
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="COMBOVALUE"/>
		</COMBOGROUP>
	</xsl:template>
	<!--======================================================================================================-->
	<xsl:template match="COMBOVALUE">
		<xsl:variable name="groupname">
			<xsl:value-of select="../@GROUPNAME"/>
		</xsl:variable>
		<xsl:variable name="valueid">
			<xsl:value-of select="@VALUEID"/>
		</xsl:variable>
		<xsl:if test="/RESPONSE/COMBOGROUP[@GROUPNAME=$groupname]/COMBOVALUE[@VALUEID=$valueid]">
			<COMBOVALUE>
				<xsl:copy-of select="/RESPONSE/COMBOGROUP[@GROUPNAME=$groupname]/COMBOVALUE[@VALUEID=$valueid]/@*"/>
				<xsl:copy-of select="/RESPONSE/COMBOGROUP[@GROUPNAME=$groupname]/COMBOVALUE[@VALUEID=$valueid]/COMBOVALIDATION[@VALIDATIONTYPE]"/>
				<COMBOVALIDATION VALIDATIONTYPE=""/>
			</COMBOVALUE>
		</xsl:if>
		<xsl:if test="not(/RESPONSE/COMBOGROUP[@GROUPNAME=$groupname]/COMBOVALUE[@VALUEID=$valueid])">
			<xsl:copy-of select="."/>
		</xsl:if>
	</xsl:template>
	<!--======================================================================================================-->
</xsl:stylesheet>
