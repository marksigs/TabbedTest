<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
<xsl:import href="document-functions.xsl"/>
<xsl:import href="document-functions-applicant.xsl"/>
<!-- PB 13/06/2006 EP721	Not using correspondence address -->
<!-- PB 08/02/2007 EP2_729 Various fixes -->
<!-- DS 13/02/2007 EP2_1360 Formatted date -->
<!-- PB 14/02/2007 EP2_729 Cosmetic changes -->
<!-- OS 14/02/2007 EP2_813 Updated OfferDate according to new specs-->
<!-- PB	27/02/2007 EP2_729 Added missing 'string()' function to javascript parameters -->

	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="INFORMATION">
				<xsl:element name="CUSTOMERDETAILS">
					<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
					<xsl:if test="//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE">
						<xsl:attribute name="OFFERDATE"><xsl:value-of select="concat('{',msg:GetDate(string(//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE)),'}')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="NAMES"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																							)"/>
					</xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:FormatAddress(	string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY),
																																				string(.//CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE))"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//CUSTOMERROLE/@APPLICATIONNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:GetSingleOrJointShortSalutation(	string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																												string( .//CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																											)"/></xsl:attribute>
					<xsl:attribute name="MORTGAGETYPE"><xsl:value-of select="msg:DealWithTypeOfApplication( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
					<!-- Check if different mortgagetype or if additional text required -->
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'FT' )" >
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'HM' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'R' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'RTB' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'ABO' )">
						<xsl:element name="ADDBRWGORPSW">
							<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
							<xsl:attribute name="NAMES"><xsl:value-of select="msg:FormatNames(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME ) )"/></xsl:attribute>
							<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
							<xsl:if test="//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE">
								<xsl:attribute name="OFFERDATE"><xsl:value-of select="concat('{',msg:GetDate(string(//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE)),'}')"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="count(//CUSTOMERROLE)=1">
								<xsl:attribute name="IWE">I</xsl:attribute>
								<xsl:attribute name="IWELC">I</xsl:attribute>
							</xsl:if>
							<xsl:if test="count(//CUSTOMERROLE)=2">
								<xsl:attribute name="IWE">We</xsl:attribute>
								<xsl:attribute name="IWELC">we</xsl:attribute>
							</xsl:if>
							<xsl:if test="//CUSTOMERROLE[@CUSTOMERORDER=1]">
								<xsl:element name="APPLICANT1">
									<xsl:attribute name="NAME"><xsl:value-of select="concat( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME,' ', //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME )"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
							<xsl:if test="//CUSTOMERROLE[@CUSTOMERORDER=2]">
								<xsl:element name="APPLICANT2">
									<xsl:attribute name="NAME"><xsl:value-of select="concat( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME, ' ', //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
						</xsl:element>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'TOE' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'ABTOE' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'NP' )">
						<xsl:element name="MAINADVORTOE"/>
					</xsl:if>
					<xsl:if test="msg:CheckForType( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ), 'PSW' )">
						<xsl:element name="ADDBRWGORPSW">
							<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
							<xsl:attribute name="NAMES"><xsl:value-of select="msg:FormatNames(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																				string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME ))"/></xsl:attribute>
							<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
							<xsl:if test="//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE">
								<xsl:attribute name="OFFERDATE"><xsl:value-of select="concat('{',msg:GetDate(string(//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE)),'}')"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="count(//CUSTOMERROLE)=1">
								<xsl:attribute name="IWE">I</xsl:attribute>
								<xsl:attribute name="IWELC">I</xsl:attribute>
							</xsl:if>
							<xsl:if test="count(//CUSTOMERROLE)=2">
								<xsl:attribute name="IWE">We</xsl:attribute>
								<xsl:attribute name="IWELC">we</xsl:attribute>
							</xsl:if>
							<xsl:if test="//CUSTOMERROLE[@CUSTOMERORDER=1]">
								<xsl:element name="APPLICANT1">
									<xsl:attribute name="NAME"><xsl:value-of select="concat( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME,' ', //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME )"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
							<xsl:if test="//CUSTOMERROLE[@CUSTOMERORDER=2]">
								<xsl:element name="APPLICANT2">
									<xsl:attribute name="NAME"><xsl:value-of select="concat( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME, ' ', //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
						</xsl:element>
					</xsl:if>
					<xsl:if test="//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE">
						<xsl:attribute name="OFFERDATE"><xsl:value-of select="concat('{',msg:GetDate(string(//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE)),'}')"/></xsl:attribute>
					</xsl:if>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>