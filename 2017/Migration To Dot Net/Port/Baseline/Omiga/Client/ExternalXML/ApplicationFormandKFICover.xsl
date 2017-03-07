<?xml version="1.0" encoding="UTF-8"?>
<!--	PRG		DATE				REF				DESCRIPTION
		PB		26/01/2007		EP2_734		Original version
		DS 13/02/2007 EP2_1360 Formatted date 



-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:import href="TypeOfMortgage.xslt"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="INFORMATION">
				<xsl:element name="CUSTOMERDETAILS">
					<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
					<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:DealWithAddressAndSalutation(
										msg:GetSingleOrJointSalutation( string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@TITLE_TEXT), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@TITLEOTHER),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@FIRSTFORENAME),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@SURNAME), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@TITLE_TEXT), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@TITLEOTHER),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@FIRSTFORENAME), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@SURNAME)
																		),
										string(.//ADDRESS/@FLATNUMBER),
										string(.//ADDRESS/@BUILDINGORHOUSENAME),
										string(.//ADDRESS/@BUILDINGORHOUSENUMBER),
										string(.//ADDRESS/@STREET),
										string(.//ADDRESS/@DISTRICT),
										string(.//ADDRESS/@TOWN),
										string(.//ADDRESS/@COUNTY),
										string(.//ADDRESS/@POSTCODE)
										)"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//APPLICATIONFACTFIND/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
					<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:GetSingleOrJointShortSalutation( string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@TITLE_TEXT), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@TITLEOTHER),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@FIRSTFORENAME),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='1']/CUSTOMERVERSION/@SURNAME), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@TITLE_TEXT), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@TITLEOTHER),
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@FIRSTFORENAME), 
																		string( .//CUSTOMERROLE[@CUSTOMERORDER='2']/CUSTOMERVERSION/@SURNAME)
																		)"/></xsl:attribute>
					<!-- Set default mortgagetype -->
					<xsl:attribute name="MORTGAGETYPE">Mortgage</xsl:attribute>
					<xsl:apply-templates select=".//COMBOVALIDATION"/>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>