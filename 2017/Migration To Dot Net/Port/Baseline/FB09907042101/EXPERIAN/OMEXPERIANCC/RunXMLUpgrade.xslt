<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: RunXMLUpgrade

History:

Version Author		Date       Description
01.00	LDM			22/03/2006	 Created . EP6 Epsom. New call to upgrade an enquiry footprint to a full application footprint
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="GlobalParam" select="/RESPONSE/GLOBALPARAMETER"/>
	<xsl:variable name="AppCreditCheck" select="/RESPONSE/APPLICATIONCREDITCHECK"/>

	<xsl:template match="/">
		<xsl:element name="EXPERIAN">
			<xsl:element name="ESERIES">
				<xsl:element name="FORM_ID">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianCCFormId']/@STRING"/>
				</xsl:element>
			</xsl:element>
			<xsl:element name="EXP">
				<xsl:element name="EXPERIANREF">
						<xsl:value-of select="/RESPONSE/APPLICATIONCREDITCHECK[1]/@CREDITCHECKREFERENCENUMBER"/>
				</xsl:element>
			</xsl:element>
			<xsl:element name="AC01">
				<xsl:element name="CLIENTKEY">
							<xsl:value-of select="$GlobalParam[@NAME='ExperianCCClientKey']/@STRING"/>
				</xsl:element>
				<xsl:element name="ENTRYPOINT">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianCCEntryPoint']/@STRING"/>
				</xsl:element>
				<xsl:element name="FUNCTION">0130</xsl:element>
			</xsl:element>
		</xsl:element>	
	</xsl:template>		
</xsl:stylesheet>
