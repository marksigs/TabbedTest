<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="REQUEST">
			<xsl:copy>
				<xsl:apply-templates select="CUSTOMERS/CUSTOMER"/>
			</xsl:copy>
	</xsl:template>
	
	<xsl:template match="CUSTOMER">
		<xsl:element name="UpdateCustomerOrProspectRequestType">
			<xsl:element name="CustomerNumber"><xsl:value-of select="@CIFNUMBER"/></xsl:element>
			<xsl:element name="ClientDetail">
				<xsl:element name="OMIGACUSTOMERNUMBER"><xsl:value-of select="@OMIGACUSTOMERNUMBER"/></xsl:element>
				<xsl:element name="Number"><xsl:value-of select="@CIFNUMBER"/></xsl:element>
				<xsl:element name="NamePrefix"><xsl:value-of select="@TITLECODE"/></xsl:element>
				<xsl:element name="FirstName"><xsl:value-of select="@FORENAME1"/></xsl:element>
				<xsl:element name="MiddleName"><xsl:value-of select="@FORENAME2"/></xsl:element>
				<xsl:element name="LastName"><xsl:value-of select="@SURNAME"/></xsl:element>
				<xsl:element name="Email1"><xsl:value-of select="@EMAILADDRESS"/></xsl:element>
				<xsl:element name="Gender"><xsl:value-of select="@GENDERCODE"/></xsl:element>
				<xsl:element name="SpecialNeed"><xsl:value-of select="@SPECIALNEEDS"/></xsl:element>
				<xsl:apply-templates select="@DATEOFBIRTH"/>
				<xsl:element name="NationalInsuranceNumber"><xsl:value-of select="@NATIONALINSURANCENUMBER"/></xsl:element>
				<xsl:element name="DontSolicit"><xsl:value-of select="@MAILSHOTREQUIRED"/></xsl:element>
				<xsl:apply-templates select="ADDRESS"/>
				<xsl:apply-templates select="PHONENUMBER"/>
			</xsl:element>
			<xsl:element name="MortgagePassword"><xsl:value-of select="@PASSWORD"/></xsl:element>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="ADDRESS">
		<xsl:element name="AddressDetail">
			<xsl:element name="AddressType"><xsl:value-of select="@ADDRESSTYPE"/></xsl:element>
			<xsl:element name="HouseOrBuildingName"><xsl:value-of select="@HOUSENAME"/></xsl:element>
			<xsl:element name="HouseOrBuildingNumber"><xsl:value-of select="@HOUSENUMBER"/></xsl:element>
			<xsl:element name="FlatNameOrNumber"><xsl:value-of select="@FLATNUMBER"/></xsl:element>
			<xsl:element name="Street"><xsl:value-of select="@STREET"/></xsl:element>
			<xsl:element name="District"><xsl:value-of select="@DISTRICT"/></xsl:element>
			<xsl:element name="TownOrCity"><xsl:value-of select="@TOWN"/></xsl:element>
			<xsl:element name="PostCode"><xsl:value-of select="@POSTCODE"/></xsl:element>
			<xsl:element name="County"><xsl:value-of select="@COUNTY"/></xsl:element>
			<xsl:element name="CountryCode"><xsl:value-of select="@COUNTRYCODE"/></xsl:element>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="PHONENUMBER">
		<xsl:element name="PhoneDetail">
			<xsl:element name="AreaCode"><xsl:value-of select="@AREACODE"/></xsl:element>
			<xsl:element name="LocalNumber"><xsl:value-of select="@NUMBER"/></xsl:element>
			<xsl:element name="Extension"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:element>
			<xsl:element name="PhoneType"><xsl:value-of select="@USAGECODE"/></xsl:element>	
		</xsl:element>
	</xsl:template>
	
		<xsl:template match="@DATEOFBIRTH">
		<xsl:element name="DateOfBirth">
			<xsl:call-template name="ConvertDate"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template name="ConvertDate">
		<xsl:variable name="Year"><xsl:value-of select="substring(.,7,4)"/></xsl:variable>
		<xsl:variable name="Month"><xsl:value-of select="substring(.,4,2)"/></xsl:variable>
		<xsl:variable name="Day"><xsl:value-of select="substring(.,1,2)"/></xsl:variable>
		<xsl:value-of select="concat($Year,'/',$Month,'/',$Day)"/>
	</xsl:template>

	
</xsl:stylesheet>
