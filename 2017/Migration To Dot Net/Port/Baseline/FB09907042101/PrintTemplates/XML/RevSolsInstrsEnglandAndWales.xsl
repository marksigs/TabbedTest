<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		04/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		27/01/2007	EP2_1000	Updated as per new spec. Could probably be made neater with abstraction
	DS	 13/02/2007	 EP2_1360	 Formatted date
	PB		15/02/2007	EP2_1000	Several minor fixes
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:import href="document-functions-solicitor.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="LETTER">
				<xsl:element name="SOLICITORDETAILS">
					<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
					<xsl:if test="//APPLICATIONLEGALREP/CONTACTDETAILS">
						<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( //APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																								string( //APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE_OTHER ),
																																								string( //APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTFORENAME ),
																																								string( //APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTSURNAME )
																																							)"/></xsl:attribute>
						<xsl:if test="string(//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTSURNAME)=''">
							<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
						</xsl:if>
						<xsl:if test="string(//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTSURNAME)!=''">
							<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
						</xsl:if>
						<xsl:if test="//APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS">
							<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(	string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTFORENAME ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/@COMPANYNAME ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@FLATNUMBER ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENAME ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@STREET ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@TOWN ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTY ),
																																							string( //APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE )
																																						)"/></xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:element>
				<xsl:element name="APPLICANT">
					<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONLEGALREP/@APPLICATIONNUMBER"/>}</xsl:attribute>
					<xsl:attribute name="TYPEOFMORTGAGE"><xsl:value-of select="msg:DealWithTypeOfApplication( string( //APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																						)"/></xsl:attribute>
				</xsl:element>
				<!-- Check if in England and Wales -->
				<xsl:if test="msg:CheckForType( string( //NEWPROPERTY/@PROPERTYLOCATION_TYPES ),'E')">
					<xsl:element name="ENGWAL">
						<xsl:attribute name="TCYEAR">{<xsl:value-of select="//GLOBALPARAMETER[@NAME='TemplateTCEditionYearEngWales']/@STRING"/>}</xsl:attribute>
					</xsl:element>
					<xsl:element name="ENGWALNI"/>
					<!-- Check if Leasehold -->
					<xsl:if test="msg:CheckForType( string( //NEWPROPERTY/@TENURETYPE_TYPES ),'L')">
						<xsl:element name="EWNILEASEHOLD"/>
					</xsl:if>
					<!-- Check for additional occupants, and list them as individual EWNIADDOCC elements -->
					<xsl:for-each select="//OTHERRESIDENT/PERSON">
						<xsl:element name="EWNIADDOCC">
							<xsl:attribute name="NAME"><xsl:value-of select="concat( ./@FIRSTFORENAME, ' ', ./@SURNAME )"/></xsl:attribute>
						</xsl:element>
					</xsl:for-each>
				</xsl:if>
				<!-- Check if in Scotland -->
				<xsl:if test="msg:CheckForType( string( //NEWPROPERTY/@PROPERTYLOCATION_TYPES ),'S')">
					<xsl:element name="SCOT">
						<xsl:attribute name="TCYEAR">{<xsl:value-of select="//GLOBALPARAMETER[@NAME='TemplateTCEditionYearScotland']/@STRING"/>}</xsl:attribute>
					</xsl:element>
				</xsl:if>
				<!-- Check if in Northern Ireland -->
				<xsl:if test="msg:CheckForType( string( //NEWPROPERTY/@PROPERTYLOCATION_TYPES ),'I')">
					<xsl:element name="NIRE">
						<xsl:attribute name="TCYEAR">{<xsl:value-of select="//GLOBALPARAMETER[@NAME='TemplateTCEditionYearNIreland']/@STRING"/>}</xsl:attribute>
					</xsl:element>
					<xsl:element name="ENGWALNI"/>
					<!-- Check if Leasehold -->
					<xsl:if test="msg:CheckForType( string( //NEWPROPERTY/@TENURETYPE_TYPES ),'L')">
						<xsl:element name="EWNILEASEHOLD"/>
					</xsl:if>
					<!-- Check for additional occupants, and list them as individual EWNIADDOCC elements -->
					<xsl:for-each select="//OTHERRESIDENT/PERSON">
						<xsl:element name="EWNIADDOCC">
							<!-- No title is held for additional occupants -->
							<xsl:attribute name="NAME"><xsl:value-of select="concat( ./@FIRSTFORENAME, ' ', ./@SURNAME )"/></xsl:attribute>
						</xsl:element>
					</xsl:for-each>
				</xsl:if>
				<xsl:if test="//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='G'">
					<xsl:element name="GUARANTOR">
						<xsl:attribute name="NAME"><xsl:value-of select="msg:DealWithSalutation(	string( //CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																					string( //CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																					string( //CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																					string( //CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='G']/CUSTOMERVERSION[last()]/@SURNAME )
																																				)"/></xsl:attribute>
					</xsl:element>
				</xsl:if>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
