<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: RunXMLCreditCheck

History:

Version Author		Date       Description
01.00	LDM			27/02/2006	EP6			New layout for for Epsom call.
01.01	LDM			05/04/2006	EP16		Restricts the length of telephone numbers, and current employer name
01.02   LDM 		13/04/2006  EP344	 	Restrict length of forename surname and the rest of the name block
01.03	LDM			25/05/2006  EP618		Force the address details by truncating, to be the right size for experian + email address
================================================================================-->

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

<xsl:template match="/">
<EXPERIAN>
	<xsl:apply-templates select="EXPERIAN/ESERIES" />    
	<xsl:apply-templates select="EXPERIAN/VER1" />    
	<xsl:apply-templates select="EXPERIAN/CTRL" />    
	<xsl:apply-templates select="EXPERIAN/AP09" />    
	<xsl:apply-templates select="EXPERIAN/NAME" />    
	<xsl:apply-templates select="EXPERIAN/ADDR" />
	<xsl:call-template name="CreateRESY"/>
	<xsl:apply-templates select="EXPERIAN/CTZI" />    
	<xsl:apply-templates select="EXPERIAN/INS1" />    
	<xsl:call-template name="CreateAUT1"/>
</EXPERIAN>
</xsl:template>

<xsl:template name="CreateAUT1">
	
	<xsl:element name="AUT1">
		
		<xsl:element name="OCCUPANCYSTATUS">
		<xsl:variable name="AddressType" select="EXPERIAN/AUT1INFO/AddressTypeVal"/>
		<xsl:variable name="NatureOfOccup" select="EXPERIAN/AUT1INFO/NatureOfOccupancyVal"/>
		<xsl:variable name="TenancyType" select="EXPERIAN/AUT1INFO/TenancyTypeVal"/>
		
		<xsl:choose>
		<xsl:when test="$AddressType='H'">
			<xsl:choose>
				<xsl:when test="$NatureOfOccup='O'">O</xsl:when>
				<xsl:when test="$NatureOfOccup='LP'">P</xsl:when>
				<xsl:when test="$NatureOfOccup='R'">X</xsl:when>
				<xsl:when test="$NatureOfOccup='RV'">
					<xsl:choose>
						<xsl:when test="$TenancyType">
							<xsl:choose>
								<xsl:when test="$TenancyType='L'">C</xsl:when>
								<xsl:when test="$TenancyType='P'">T</xsl:when>
								<xsl:otherwise>X</xsl:otherwise>
							</xsl:choose>				
						</xsl:when>
						<xsl:otherwise>T</xsl:otherwise>
					</xsl:choose>				
				</xsl:when>
				<xsl:when test="$NatureOfOccup='RB'">
					<xsl:choose>
						<xsl:when test="$TenancyType">
							<xsl:choose>
								<xsl:when test="$TenancyType='L'">C</xsl:when>
								<xsl:when test="$TenancyType='P'">T</xsl:when>
								<xsl:otherwise>X</xsl:otherwise>
							</xsl:choose>				
						</xsl:when>
						<xsl:otherwise>C</xsl:otherwise>
					</xsl:choose>				
				</xsl:when>
				<xsl:when test="$NatureOfOccup='X'">X</xsl:when>
				<xsl:otherwise>Z</xsl:otherwise>
			</xsl:choose>				
		</xsl:when>
		<xsl:otherwise>Z</xsl:otherwise>
		</xsl:choose>
		</xsl:element>
		
		<xsl:element name="HOMETELSTD">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/HomeAreaCode"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/HomeAreaCode, 1, 6)"/></xsl:when>
				<xsl:otherwise>N</xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		<xsl:element name="HOMETELLOCAL">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/HomeTelNo"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/HomeTelNo, 1, 10)"/></xsl:when>
				<xsl:otherwise> </xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		
		<xsl:element name="MOBILEPHONESTD">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/MobileAreaCode"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/MobileAreaCode, 1, 6)"/></xsl:when>
				<xsl:otherwise>N</xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		<xsl:element name="MOBILEPHONELOCAL">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/MobileTelNo"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/MobileTelNo, 1, 10)"/></xsl:when>
				<xsl:otherwise>N</xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		
		<xsl:element name="SORTCODE"/>
		<xsl:element name="ACCOUNTNUMBER"/>
		<xsl:element name="MOTHERMAIDENNAME"/>
		<xsl:element name="BIRTHSURNAME"/>
		<xsl:element name="PLACEOFBIRTH"/>
		<xsl:element name="BIRTHCOUNTRY"/>
		
		<xsl:element name="MARITALSTATUS">
			<xsl:variable name="MaritalStatus" select="EXPERIAN/AUT1INFO/MaritalStatusVal"/>
			<xsl:choose>
				<xsl:when test="$MaritalStatus='M'">M</xsl:when>
				<xsl:when test="$MaritalStatus='S'">S</xsl:when>
				<xsl:when test="$MaritalStatus='D'">D</xsl:when>
				<xsl:when test="$MaritalStatus='W'">W</xsl:when>
				<xsl:when test="$MaritalStatus='E'">E</xsl:when>
				<xsl:when test="$MaritalStatus='C'">C</xsl:when>
				<xsl:when test="$MaritalStatus='X'">X</xsl:when>
				<xsl:when test="$MaritalStatus='O'">O</xsl:when>
				<xsl:otherwise>Z</xsl:otherwise>
			</xsl:choose>				
		
		</xsl:element>
		
		<xsl:element name="DEPENDANTS">
		<xsl:variable name="NumberOfDependants" select="EXPERIAN/AUT1INFO/NoOfDependants"/>
			<xsl:choose>
				<xsl:when test="$NumberOfDependants">
					<xsl:choose >
						<xsl:when test="(number($NumberOfDependants) &gt; -1) and (number($NumberOfDependants) &lt; 8)"><xsl:value-of select="$NumberOfDependants"/></xsl:when>
						<xsl:when test="number($NumberOfDependants) &gt; 7"><xsl:value-of select="7"/></xsl:when>
						<xsl:otherwise>8</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>8</xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
			
		<xsl:element name="VRM"/>
		<xsl:element name="NINO"/>
		<xsl:element name="DVLANO"/>
		<xsl:element name="PASSPORTNO"/>
		<xsl:element name="TIMEINCURREMP"><xsl:value-of select="EXPERIAN/AUT1INFO/TimeInEmp"/></xsl:element>
		<xsl:element name="CURREMPLOYERNAME"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/CompanyName, 1, 30)"/></xsl:element>		
			
		<xsl:element name="WORKSTD">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/WorkAreaCode"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/WorkAreaCode, 1, 6)"/></xsl:when>
				<xsl:otherwise>N</xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		<xsl:element name="WORKLOCAL">
			<xsl:choose>
				<xsl:when test="EXPERIAN/AUT1INFO/WorkTelNo"><xsl:value-of select="substring(EXPERIAN/AUT1INFO/WorkTelNo, 1, 10)"/></xsl:when>
				<xsl:otherwise> </xsl:otherwise>
			</xsl:choose>		
		</xsl:element>
		
	</xsl:element>
</xsl:template>



<xsl:template name="CreateRESY">
	<xsl:variable name="NameArray1" select="EXPERIAN/NamesArr[ADDRBLOCK=1]"/>
	<xsl:variable name="NameArray2" select="EXPERIAN/NamesArr[ADDRBLOCK=2]"/>
	
	<xsl:choose>
	<xsl:when test="$NameArray1">
		<xsl:element name="RESY">
			<xsl:element name="NAMESCOUNT">
			<xsl:value-of select="$NameArray1/NAMESCOUNT"/>
			</xsl:element>
			
			<xsl:for-each select="$NameArray1">
				<xsl:element name="NAMESEQUENCE"><xsl:value-of select="NAMESEQUENCE"/></xsl:element>
				<xsl:element name="DATEFROM_DD"><xsl:value-of select="DATEFROM_DD"/></xsl:element>
				<xsl:element name="DATEFROM_MM"><xsl:value-of select="DATEFROM_MM"/></xsl:element>
				<xsl:element name="DATEFROM_CCYY"><xsl:value-of select="DATEFROM_CCYY"/></xsl:element>
				<xsl:element name="DATETO_DD"><xsl:value-of select="DATETO_DD"/></xsl:element>
				<xsl:element name="DATETO_MM"><xsl:value-of select="DATETO_MM"/></xsl:element>
				<xsl:element name="DATETO_CCYY"><xsl:value-of select="DATETO_CCYY"/></xsl:element>
			</xsl:for-each>
			
		</xsl:element>
	</xsl:when>
	</xsl:choose>

	<xsl:choose>
	<xsl:when test="$NameArray2">
		<xsl:element name="RESY">
			<xsl:element name="NAMESCOUNT">
			<xsl:value-of select="$NameArray2/NAMESCOUNT"/>
			</xsl:element>
			
			<xsl:for-each select="$NameArray2">
				<xsl:element name="NAMESEQUENCE"><xsl:value-of select="NAMESEQUENCE"/></xsl:element>
				<xsl:element name="DATEFROM_DD"><xsl:value-of select="DATEFROM_DD"/></xsl:element>
				<xsl:element name="DATEFROM_MM"><xsl:value-of select="DATEFROM_MM"/></xsl:element>
				<xsl:element name="DATEFROM_CCYY"><xsl:value-of select="DATEFROM_CCYY"/></xsl:element>
				<xsl:element name="DATETO_DD"><xsl:value-of select="DATETO_DD"/></xsl:element>
				<xsl:element name="DATETO_MM"><xsl:value-of select="DATETO_MM"/></xsl:element>
				<xsl:element name="DATETO_CCYY"><xsl:value-of select="DATETO_CCYY"/></xsl:element>
			</xsl:for-each>
			
		</xsl:element>
	</xsl:when>
	</xsl:choose>
	
</xsl:template>
	


<xsl:template match="ESERIES">
	<xsl:copy-of select="." />
</xsl:template>
<xsl:template match="VER1">
	<xsl:copy-of select="." />
</xsl:template>
<xsl:template match="CTRL">
	<xsl:copy-of select="." />
</xsl:template>
<xsl:template match="AP09">
	<xsl:copy-of select="." />
</xsl:template>

<xsl:template match="NAME">
	<xsl:element name="NAME">
		<xsl:element name="CODE"/>
		<xsl:element name="TITLE"><xsl:value-of select="substring(TITLE,1,10)"/></xsl:element>
		<xsl:element name="FORENAME"><xsl:value-of select="substring(FORENAME,1,15)"/></xsl:element>
		<xsl:element name="INITIALS"><xsl:value-of select="substring(INITIALS,1,15)"/></xsl:element>
		<xsl:element name="SURNAME"><xsl:value-of select="substring(SURNAME,1,30)"/></xsl:element>
		<xsl:element name="SUFFIX"/>
		<xsl:element name="DATEOFBIRTH_DD"><xsl:value-of select="DATEOFBIRTH_DD"/></xsl:element>
		<xsl:element name="DATEOFBIRTH_MM"><xsl:value-of select="DATEOFBIRTH_MM"/></xsl:element>
		<xsl:element name="DATEOFBIRTH_CCYY"><xsl:value-of select="DATEOFBIRTH_CCYY"/></xsl:element>
		<xsl:element name="SEX"><xsl:value-of select="SEX"/></xsl:element>
	</xsl:element>
</xsl:template>

<xsl:template match="ADDR">
	<xsl:element name="ADDR">
		<xsl:element name="FLAT"><xsl:value-of select="substring(FLAT,1,16)"/></xsl:element>
		<xsl:element name="HOUSENAME"><xsl:value-of select="substring(HOUSENAME,1,26)"/></xsl:element>
		<xsl:element name="HOUSENUMBER"><xsl:value-of select="substring(HOUSENUMBER,1,10)"/></xsl:element>
		<xsl:element name="STREET"><xsl:value-of select="substring(STREET,1,40)"/></xsl:element>
		<xsl:element name="DISTRICT"><xsl:value-of select="substring(DISTRICT,1,30)"/></xsl:element>
		<xsl:element name="TOWN"><xsl:value-of select="substring(TOWN,1,20)"/></xsl:element>
		<xsl:element name="COUNTY"><xsl:value-of select="substring(COUNTY,1,20)"/></xsl:element>
		<xsl:element name="POSTCODE"><xsl:value-of select="substring(POSTCODE,1,8)"/></xsl:element>
	</xsl:element>
</xsl:template>

<xsl:template match="CTZI">
	<xsl:copy-of select="." />
</xsl:template>
<xsl:template match="INS1">
	<xsl:element name="INS1">
		<xsl:element name="EMAILADDRESS"><xsl:value-of select="substring(EMAILADDRESS,1,60)"/></xsl:element>
	</xsl:element>
</xsl:template>
	
</xsl:stylesheet>
