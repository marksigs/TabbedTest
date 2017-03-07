<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		04/07/2006	EP543			Incorporate 'Other' title, e.g. Lord, Baron, Reverend etc...
	PE		15/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		24/01/2007	EP2_805		Updated as per new spec
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
		<xsl:element name="LETTER">
			<xsl:element name="LENDERDETAILS">
				<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
				<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(	'','','','',
																																				string( .//THIRDPARTY/@COMPANYNAME ),
																																				string( .//THIRDPARTY/ADDRESS/@FLATNUMBER ),
																																				string( .//THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME ),
																																				string( .//THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																				string( .//THIRDPARTY/ADDRESS/@STREET ),
																																				string( .//THIRDPARTY/ADDRESS/@DISTRICT ),
																																				string( .//THIRDPARTY/ADDRESS/@TOWN ),
																																				string( .//THIRDPARTY/ADDRESS/@COUNTY ),
																																				string( .//THIRDPARTY/ADDRESS/@POSTCODE )
																																				)"/></xsl:attribute>
				<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( .//THIRDPARTY/CONTACTDETAILS/@TITLE_TEXT ),
																																									string( .//THIRDPARTY/CONTACTDETAILS/@TITLE_OTHER ),
																																									string( .//THIRDPARTY/CONTACTDETAILS/@FIRSTFORENAME ),
																																									string( .//THIRDPARTY/CONTACTDETAILS/@SURNAME )
																																									)"/></xsl:attribute>
				<xsl:if test="not(//THIRDPARTY/CONTACTDETAILS/@SURNAME)">
					<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
				</xsl:if>
				<xsl:if test=".//THIRDPARTY/CONTACTDETAILS/@SURNAME">
					<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
				</xsl:if>
			</xsl:element>
			<xsl:if test=".//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='A'">
				<xsl:element name="APPLICANT">
					<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																							)"/></xsl:attribute>
					<xsl:attribute name="REFERENCE">{<xsl:value-of select=".//ACCOUNT/@ACCOUNTNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
			<xsl:if test=".//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='G'">
				<xsl:element name="GUARANTOR">
					<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES )"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERROLE[CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																							)"/></xsl:attribute>
					<xsl:attribute name="REFERENCE">{<xsl:value-of select=".//ACCOUNT/@ACCOUNTNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																								string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
