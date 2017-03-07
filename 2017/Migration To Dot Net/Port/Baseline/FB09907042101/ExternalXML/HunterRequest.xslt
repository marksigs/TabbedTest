<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: RunXMLCreditCheck

History:

Version 	Author		Date       		Description
01.00   	SAB 		27/04/2006	EP463 New tranformation
01.01   	LDM			30/05/2006	EP627 Apply mars change MAR1395 in experianhibo(ver 13) to transform   
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="Response" select="/RESPONSE"/>
	<xsl:variable name="Customer" select="$Response/CUSTOMER"/>
	<xsl:variable name="PropertyType" select="$Response/PROPERTY[@GROUPNAME='PropertyType']"/>
	<xsl:variable name="PropertyDesc" select="$Response/PROPERTY[@GROUPNAME='PropertyDescription']"/>
	<xsl:variable name="Employer" select="$Response/EMPLOYER"/>
	<xsl:variable name="Security" select="$Response/SECADADDRESS"/>
	<xsl:variable name="Solicitor" select="$Response/SOLADDRESS"/>	
	<xsl:variable name="Bank" select="$Response/BANK"/>	
	<xsl:variable name="GlobalParam" select="$Response/GLOBALPARAMETER"/>	
	<xsl:template match="/">
		<xsl:element name="EXPERIAN">
			<xsl:element name="ESERIES">
				<xsl:element name="FORM_ID">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianHunterFormID']/@STRING"/>
				</xsl:element>
			</xsl:element>
			<xsl:element name="EXP">
				<xsl:element name="EXPERIANREF">
					<xsl:value-of select="$Response/APPLICATIONCREDITCHECK/@CREDITCHECKREFERENCENUMBER"/>
				</xsl:element>
			</xsl:element>
			<xsl:element name="AC01">
				<xsl:element name="CLIENTKEY">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianHunterClientKey']/@STRING"/>
				</xsl:element>
				<xsl:element name="ENTRYPOINT">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianHunterEntryPoint']/@STRING"/>
				</xsl:element>
				<xsl:element name="FUNCTION">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianHunterFunction']/@STRING"/>
				</xsl:element>
			</xsl:element>
			<xsl:for-each select="$Customer">
				<xsl:element name="HUNI">
					<xsl:element name="APPLICANTTYPE">
						<xsl:choose>
							<xsl:when test="position() = 1">1</xsl:when>
							<xsl:otherwise>2</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="APPNO"><xsl:value-of select="$Response/APPLICATION/@APPLICATIONNUMBER"/></xsl:element>
					<xsl:element name="SEQUENCENUMBER"><xsl:value-of select="concat('0', position())"/></xsl:element>
					<xsl:element name="SECADDFLAT"><xsl:value-of select="$Security/@FLATNUMBER"/></xsl:element>					
					<xsl:element name="SECADDHOUSENAME"><xsl:value-of select="$Security/@BUILDINGORHOUSENAME"/></xsl:element>		
					<xsl:element name="SECADDHOUSENUMBER"><xsl:value-of select="$Security/@BUILDINGORHOUSENUMBER"/></xsl:element>
					<xsl:element name="SECADDSTREET"><xsl:value-of select="$Security/@STREET"/></xsl:element>					
					<xsl:element name="SECADDDISTRICT"><xsl:value-of select="$Security/@DISTRICT"/></xsl:element>
					<xsl:element name="SECADDTOWN"><xsl:value-of select="$Security/@TOWN"/></xsl:element>					
					<xsl:element name="SECADDCOUNTY"><xsl:value-of select="$Security/@COUNTY"/></xsl:element>
					<xsl:element name="SECADDPOSTCODE"><xsl:value-of select="$Security/@POSTCODE"/></xsl:element>					
					<xsl:call-template name="ProcessACCTYPE"/>
					<xsl:call-template name="ProcessEMP">
						<xsl:with-param name="CustNo"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:with-param>
					</xsl:call-template>
					<xsl:element name="SOLNAME"><xsl:value-of select="$Solicitor/@SOLNAME"/></xsl:element>		
					<xsl:element name="SOLBUILDINGNAME"><xsl:value-of select="$Solicitor/@BUILDINGORHOUSENAME"/></xsl:element>
					<xsl:element name="SOLBUILDINGNUMBER"><xsl:value-of select="$Solicitor/@BUILDINGORHOUSENUMBER"/></xsl:element>					
					<xsl:element name="SOLSTREET"><xsl:value-of select="$Solicitor/@STREET"/></xsl:element>
					<xsl:element name="SOLDISTRICT"><xsl:value-of select="$Solicitor/@DISTRICT"/></xsl:element>					
					<xsl:element name="SOLTOWN"><xsl:value-of select="$Solicitor/@TOWN"/></xsl:element>
					<xsl:element name="SOLCOUNTY"><xsl:value-of select="$Solicitor/@COUNTY"/></xsl:element>					
					<xsl:element name="SOLPOSTCODE"><xsl:value-of select="$Solicitor/@POSTCODE"/></xsl:element>
					<xsl:element name="BSORTCODE"><xsl:value-of select="$Bank/@BSORTCODE"/></xsl:element>
					<xsl:element name="BACCTNUMBER"><xsl:value-of select="$Bank/@BACCTNUMBER"/></xsl:element>					
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="ProcessACCTYPE">
		<xsl:element name="ACCTYPE">
			<xsl:choose>
				<xsl:when test="$PropertyType[@VALIDATIONTYPE='H']">
					<xsl:choose>
						<xsl:when test="$PropertyDesc[@VALIDATIONTYPE='HD']">D</xsl:when>				
						<xsl:when test="$PropertyDesc[@VALIDATIONTYPE='HS']">S</xsl:when>				
						<xsl:when test="$PropertyDesc[@VALIDATIONTYPE='HT']">T</xsl:when>				
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$PropertyType[@VALIDATIONTYPE='B']">B</xsl:when>
				<xsl:when test="$PropertyType[@VALIDATIONTYPE='F']">F</xsl:when>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template name="ProcessEMP">
		<xsl:param name="CustNo"/>
		<xsl:variable name="Employer" select="$Employer[@CUSTOMERNUMBER=$CustNo]"/>
		<xsl:element name="EMPLOYERNAME"><xsl:value-of select="substring($Employer/@EMPLOYERNAME, 1, 26)"/></xsl:element>
		<xsl:element name="EMPBUILDINGNAME"><xsl:value-of select="$Employer/@BUILDINGORHOUSENAME"/></xsl:element>
		<xsl:element name="EMPBUILDINGNUMBER"><xsl:value-of select="$Employer/@BUILDINGORHOUSENUMBER"/></xsl:element>				
		<xsl:element name="EMPSTREET"><xsl:value-of select="$Employer/@STREET"/></xsl:element>
		<xsl:element name="EMPDISTRICT"><xsl:value-of select="$Employer/@DISTRICT"/></xsl:element>
		<xsl:element name="EMPTOWN"><xsl:value-of select="$Employer/@TOWN"/></xsl:element>
		<xsl:element name="EMPCOUNTY"><xsl:value-of select="$Employer/@COUNTY"/></xsl:element>
		<xsl:element name="EMPPOSTCODE"><xsl:value-of select="$Employer/@POSTCODE"/></xsl:element>
	</xsl:template>
</xsl:stylesheet>