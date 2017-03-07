<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

	<xsl:template match="RESPONSE">
		<xsl:copy>
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:apply-templates select="GetCustomerOrProspectDetailsResponseType/ClientDetails"/>
		</xsl:copy>
	</xsl:template>
	
	<xsl:template match="GetCustomerOrProspectDetailsResponseType">
		<xsl:element name="RESPONSE">
			<xsl:choose>
				<xsl:when test="ErrorCode=0">
					<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
					<xsl:apply-templates select="ClientDetails"/>
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
			<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="Number"/></xsl:attribute>
			<xsl:attribute name="SURNAME"><xsl:value-of select="LastName"/></xsl:attribute>
			<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="FirstName"/></xsl:attribute>
			<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="MiddleName"/></xsl:attribute>
			<xsl:attribute name="TITLE"><xsl:value-of select="NamePrefix"/></xsl:attribute>
			<xsl:apply-templates  select="DateOfBirth"/>
			<xsl:attribute name="GENDER"><xsl:value-of select="Gender"/></xsl:attribute>
			<xsl:attribute name="CONTACTEMAILADDRESS"><xsl:value-of select="Email1"/></xsl:attribute>
			<xsl:attribute name="NATIONALINSURANCENUMBER"><xsl:value-of select="NationalInsuranceNumber"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERSTATUS"><xsl:value-of select="ClientStatus"/></xsl:attribute>
			<xsl:attribute name="SPECIALNEEDS"><xsl:value-of select="SpecialNeed"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERKYCSTATUS"><xsl:value-of select="../KYCStatus"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERKYCADDRESSFLAG"><xsl:value-of select="../KYCAddressFlag"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERKYCIDFLAG"><xsl:value-of select="../KYCIdentificationFlag"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERCATEGORY"><xsl:value-of select="../CustomerCategory"/></xsl:attribute>
			<xsl:attribute name="MAILSHOTREQUIRED"><xsl:value-of select="DontSolicit"/></xsl:attribute>
			<xsl:attribute name="PROSPECTPASSWORDEXISTS"><xsl:value-of select="../PasswordStatus"/></xsl:attribute>
			<xsl:apply-templates select="AddressDetail"/>
		
			<xsl:if test="count(PhoneDetail)>0">
				<xsl:element name="TELEPHONENUMBERLIST">
					<xsl:apply-templates select="PhoneDetail"/>
				</xsl:element>
			</xsl:if>			
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="AddressDetail[AddressType='H']">
		<xsl:element name="CURRENTADDRESS">
			<xsl:call-template name="FormatAddress"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template match="AddressDetail[AddressType='M']">
		<xsl:element name="CORRESPONDENCEADDRESS">
			<xsl:call-template name="FormatAddress"/>
		</xsl:element>
	</xsl:template>
	
	<xsl:template name="FormatAddress">
		<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="HouseOrBuildingName"/></xsl:attribute>
		<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="HouseOrBuildingNumber"/></xsl:attribute>
		<xsl:attribute name="FLATNUMBER"><xsl:value-of select="FlatNameOrNumber"/></xsl:attribute>
		<xsl:attribute name="STREET"><xsl:value-of select="Street"/></xsl:attribute>
		<xsl:attribute name="DISTRICT"><xsl:value-of select="District"/></xsl:attribute>
		<xsl:attribute name="TOWN"><xsl:value-of select="TownOrCity"/></xsl:attribute>
		<xsl:attribute name="COUNTY"><xsl:value-of select="County"/></xsl:attribute>
		<xsl:attribute name="COUNTRY"><xsl:value-of select="CountryCode"/></xsl:attribute>
		<xsl:attribute name="POSTCODE"><xsl:value-of select="PostCode"/></xsl:attribute>
	</xsl:template>
	
	<xsl:template match="PhoneDetail">
		<xsl:element name="TELEPHONEDETAILS">
			<xsl:attribute name="USAGE"><xsl:value-of select="PhoneType"/></xsl:attribute>
			<xsl:attribute name="AREACODE"><xsl:value-of select="AreaCode"/></xsl:attribute>
			<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="LocalNumber"/></xsl:attribute>
			<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="Extension"/></xsl:attribute>
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

