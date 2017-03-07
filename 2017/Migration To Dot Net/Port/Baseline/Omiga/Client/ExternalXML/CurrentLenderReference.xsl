<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		04/07/2006	EP543			Incorporate 'Other' title, e.g. Lord, Baron, Reverend etc...
	PE		15/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		24/01/2007	EP2_805		Updated as per new spec
	PB		06/02/2007	EP2_805		Some details required modifying to match changes to imported files
	DS	 13/02/2007	 EP2_1360	 Formatted date
	PB		22/02/2007	EP2_805		Additional changes to cope with cases with an unusual amount of mortgageaccount records
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/CURRENTLENDERREFERENCE/APPLICATIONFACTFIND/APPLICATIONMORTGAGEACCOUNTS/MORTGAGEACCOUNT[not (@REDEMPTIONSTATUS_TYPE_A)]"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="MORTGAGEACCOUNT">
			
					<!-- We might be passed a customer id, or we might need to get all customers for the application -->
					<!-- We might also be passed an account guid, or we might just need to print letters for all -->
					
		<!-- If we have been passed a customer number, give details for that customer only -->
		<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER">
			<!-- If APPLICANT...  -->
			<xsl:if test=".//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]">
				<xsl:element name="LETTER">
					<xsl:element name="APPLICANT">
						<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TEXT"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
						
						<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:FormatNames(	string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERVERSION[last()]/@SURNAME ),
																																								string( '' ),
																																								string( '' )
																																							)"/></xsl:attribute>
						<!-- If there is a correspondence address use it -->
						<xsl:if test=".//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']">
							<!-- Correspondence address -->
							<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@FLATNUMBER ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINDORHOUSENAME ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@STREET ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@DISTRICT ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@TOWN ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@COUNTY ),
																																															string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@POSTCODE )
																																														)"/></xsl:attribute>
						</xsl:if>
						<!-- If there is no correspondence address... -->
						<xsl:if test="not(.//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C'])">
							<!-- ...then use the current address -->
							<xsl:if test=".//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']">
								<!-- Current address -->
								<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINDORHOUSENAME ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																																string( .//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																															)"/></xsl:attribute>
								<!-- Do we have a specific lender? -->
								<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
									<!-- Yes - only use the one requested -->
									<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
								</xsl:if>
								<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
									<!-- No specific lender so use any -->
									<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
								</xsl:if>
							</xsl:if>
						</xsl:if>
					</xsl:element>
					
					<xsl:element name="LENDERDETAILS">
						<xsl:attribute name="CURRENTDATE">{<xsl:value-of select="msg:GetDate()"/>}</xsl:attribute>
						
						<!-- Do we have a specific lender? -->
						<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
							<!-- Yes - only use the one requested -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
						<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
							<!-- No specific lender so use first one -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
					</xsl:element>
					
				</xsl:element>
			</xsl:if>
			
			<!-- If GUARANTOR... -->
			<xsl:if test=".//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G'][@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]">
				<xsl:element name="LETTER">
					<!-- Pagebreak if we need to -->
					<xsl:if test="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']"><xsl:element name="PAGEBREAK"/></xsl:if>
										
					<xsl:call-template name="GUARANTOR">
						<xsl:with-param name="RESPONSE" select="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G'][@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]"/>
					</xsl:call-template>

					<xsl:element name="LENDERDETAILS">
						<xsl:attribute name="CURRENTDATE">{<xsl:value-of select="msg:GetDate()"/>}</xsl:attribute>
						
						<!-- Do we have a specific lender? -->
						<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
							<!-- Yes - only use the one requested -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
						<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
							<!-- No specific lender so usefirst one -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
					</xsl:element>
				</xsl:element>
			</xsl:if>
		</xsl:if>
		
		
		<!-- If we have NOT been passed a customer number, give details for all customers -->
		<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER)">
			<!-- If APPLICANT...  -->
			<xsl:if test=".//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']">
				<xsl:element name="LETTER">
					<xsl:element name="APPLICANT">
						<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TEXT"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
						
						<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:FormatNames(	string( .//CUSTOMERROLE/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//CUSTOMERROLE/CUSTOMERVERSION[last()]/@SURNAME ),
																																								string( '' ),
																																								string( '' )
																																							)"/></xsl:attribute>
						<!-- If there is a correspondence address use it -->
						<xsl:if test=".//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']">
							<!-- Correspondence address -->
							<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@FLATNUMBER ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINDORHOUSENAME ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@STREET ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@DISTRICT ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@TOWN ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@COUNTY ),
																																															string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@POSTCODE )
																																														)"/></xsl:attribute>
						</xsl:if>
						<!-- If there is no correspondence address... -->
						<xsl:if test="not(.//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C'])">
							<!-- ...then use the current address -->
							<xsl:if test=".//CUSTOMERROLE[@CUSTOMERNUMBER=/RESPONSE/CURRENTLENDERREFERENCE/@CUSTOMERNUMBER]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']">
								<!-- Current address -->
								<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINDORHOUSENAME ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																																string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																															)"/></xsl:attribute>
								<!-- Do we have a specific lender? -->
								<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
									<!-- Yes - only use the one requested -->
									<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
								</xsl:if>
								<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
									<!-- No specific lender so use any -->
									<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
								</xsl:if>
							</xsl:if>
						</xsl:if>
					</xsl:element>
					
					<xsl:element name="LENDERDETAILS">
						<xsl:attribute name="CURRENTDATE">{<xsl:value-of select="msg:GetDate()"/>}</xsl:attribute>
						
						<!-- Do we have a specific lender? -->
						<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
							<!-- Yes - only use the one requested -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
						<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
							<!-- No specific lender so use any -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
					</xsl:element>
					
				</xsl:element>
			</xsl:if>
			
			<!-- If GUARANTOR... -->
			<xsl:if test=".//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']">
				<xsl:element name="LETTER">
					<!-- Pagebreak if we need to -->
					<xsl:if test="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']"><xsl:element name="PAGEBREAK"/></xsl:if>
										
					<xsl:call-template name="GUARANTOR">
						<xsl:with-param name="RESPONSE" select="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']"/>
					</xsl:call-template>
						
					<xsl:element name="LENDERDETAILS">
						<xsl:attribute name="CURRENTDATE">{<xsl:value-of select="msg:GetDate()"/>}</xsl:attribute>
						
						<!-- Do we have a specific lender? -->
						<xsl:if test="/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID">
							<!-- Yes - only use the one requested -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[@ACCOUNTGUID=/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
						<xsl:if test="not(/RESPONSE/CURRENTLENDERREFERENCE/@ACCOUNTGUID)">
							<!-- No specific lender so use first one -->
							<xsl:attribute name="LENDERREFERENCE">{<xsl:value-of select="//APPLICATIONMORTGAGEACCOUNTS[1]/ACCOUNT[1]/@ACCOUNTNUMBER"/>}</xsl:attribute>
							<xsl:apply-templates select="ACCOUNT[1]/THIRDPARTY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[1]/THIRDPARTY/CONTACTDETAILS"/>
							<xsl:apply-templates select="ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
							<xsl:apply-templates select="ACCOUNT[1]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
						</xsl:if>
					</xsl:element>
				</xsl:element>
			</xsl:if>
			
		</xsl:if>
		
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
