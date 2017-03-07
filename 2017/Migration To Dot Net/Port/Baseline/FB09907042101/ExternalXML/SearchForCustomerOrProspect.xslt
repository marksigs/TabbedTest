<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

	<xsl:template match="SearchForCustomerOrProspectResponseType">
		<xsl:element name="RESPONSE">
			<xsl:choose>
				<xsl:when test="ErrorCode=0">
					<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
					<xsl:if test="count(ClientDetails) > 0">
						<xsl:element name="CUSTOMERLIST">
							<xsl:apply-templates select="ClientDetails"/>
						</xsl:element>
					</xsl:if>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="TYPE">APERR</xsl:attribute>
					<xsl:element name="ERROR">
						<xsl:element name="NUMBER"><xsl:value-of select="ErrorCode"></xsl:value-of></xsl:element>
						<xsl:element name="DESCRIPTION"><xsl:value-of select="ErrorMessage"></xsl:value-of></xsl:element>
						<xsl:element name="SOURCE">DirectGateway</xsl:element>
						<xsl:element name="VERSION"></xsl:element>
					</xsl:element>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="ClientDetails">
		<xsl:element name="CUSTOMER">
			<xsl:attribute name="OTHERSYSTEMCUSTOMERNUMBER"><xsl:value-of select="Number"/></xsl:attribute>
			<xsl:attribute name="SURNAME"><xsl:value-of select="LastName"/></xsl:attribute>
			<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="FirstName"/></xsl:attribute>
			<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="MiddleName"/></xsl:attribute>
			<xsl:attribute name="CORRESPONDENCESALUTATION"><xsl:value-of select="concat(FirstName,' ',LastName)"/></xsl:attribute>
			<xsl:apply-templates  select="DateOfBirth"/>
			<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="AddressDetail[1]/HouseOrBuildingName"/></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="AddressDetail[1]/HouseOrBuildingNumber"/></xsl:attribute>
			<xsl:attribute name="FLATNUMBER"><xsl:value-of select="AddressDetail[1]/FlatNameOrNumber"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="AddressDetail[1]/Street"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="AddressDetail[1]/District"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="AddressDetail[1]/TownOrCity"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="AddressDetail[1]/County"/></xsl:attribute>
			<xsl:attribute name="COUNTRY"><xsl:value-of select="AddressDetail[1]/CountryCode"/></xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="AddressDetail[1]/PostCode"/></xsl:attribute>
			<xsl:attribute name="ADDRESSTYPE"><xsl:value-of select="AddressDetail[1]/AddressType"/></xsl:attribute>
			<xsl:attribute name="OTHERSYSTEMTYPE"><xsl:value-of select="../OtherCustomerType"/></xsl:attribute>
			<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="Email1"/></xsl:attribute>
			<xsl:attribute name="GENDER"><xsl:value-of select="Gender"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERSTATUS"><xsl:value-of select="ClientStatus"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="DateOfBirth">
		<xsl:attribute name="DATEOFBIRTH">
			<xsl:call-template name="ConvertDate"/>
		</xsl:attribute>
	</xsl:template>

	<xsl:template name="ConvertDate">
		<xsl:variable name="Year"><xsl:value-of select="substring(.,1,4)"/></xsl:variable>
		<xsl:variable name="Month"><xsl:value-of select="substring(.,6,2)"/></xsl:variable>
		<xsl:variable name="Day"><xsl:value-of select="substring(.,9,2)"/></xsl:variable>
		<xsl:value-of select="concat($Day,'/',$Month,'/',$Year)"/>
	</xsl:template>

	
</xsl:stylesheet>

