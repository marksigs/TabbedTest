<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
<!--
	
	File: ConditionRemovalBroker.xsl
	Created By: P Buck
	Purpose: Provides data transformation for broker's copy of condition removal letter
	
	History
	=================================================================================================
	PB		04/01/2007		EP2_505		Template: 11b Condition Removal - Broker.
	DS 		13/02/2007        EP2_1360 	Formatted date
	DS 		14/02/2007        EP2_505 		Removed old code and added standard broker code
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-broker.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/CONDITIONREMOVALBROKER/APPLICATION/APPLICATIONINTRODUCER"/>
			<xsl:call-template name="APPLICANT">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/CONDITIONREMOVALBROKER"/>
			</xsl:call-template>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->

</xsl:stylesheet>
