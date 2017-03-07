<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		12/07/2006	EP543			Added 'TitleOther' to name detail
	PE		02/08/2006	EP1031		Added Sincereley/Faithfully attribute
	PE		08/08/2006	EP1031		Reworked. Abstracted javascript functions
	PB		24/01/2007	EP2_807		Modified as per new spec
	PB		14/02/2007	EP2_807		Fixes to previous changes
	PB		25/02/2007	EP2_807		Corrected typo
	OS		28/02/2007	EP2_807		Removed numbers from salutation text
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-employer.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>	
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="EMPLOYERDETAILS">
				<xsl:if test="//EMPLOYMENT/THIRDPARTY/ADDRESS/@POSTCODE">
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(	string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																					string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																					string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																					string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME ),
																																					string( //EMPLOYMENT/THIRDPARTY/@COMPANYNAME ),
																																					string(//EMPLOYMENT/THIRDPARTY/ADDRESS/@FLATNUMBER ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@STREET ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@DISTRICT ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@TOWN ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@COUNTY ),
																																					string( //EMPLOYMENT/THIRDPARTY/ADDRESS/@POSTCODE )
																																				)"/></xsl:attribute>
					<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																							string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_OTHER ),
																																							string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																							string( //EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME )
																																						)"/></xsl:attribute>
				<xsl:if test="string( .//EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME )=''">
					<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
				</xsl:if>
				<xsl:if test="string( .//EMPLOYMENT/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME )!=''">
					<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
				</xsl:if>
				</xsl:if>

				<xsl:if test="//EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE">
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(	string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTFORENAME ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/@COMPANYNAME ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@FLATNUMBER ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@STREET ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@TOWN ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTY ),
																																					string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE )
																																				)"/></xsl:attribute>
					<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																							string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_OTHER ),
																																							string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTFORENAME ),
																																							string( //EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME )
																																						)"/></xsl:attribute>
				<xsl:if test="string( .//EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME )=''">
					<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
				</xsl:if>
				<xsl:if test="string( .//EMPLOYMENT/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME )!=''">
					<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
				</xsl:if>
				</xsl:if>
				<xsl:attribute name="CURRENTDATE">{<xsl:value-of select="msg:GetDate()"/>}</xsl:attribute>
				<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select=".//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
				<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																											string( //CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																											string( //CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																											string( //CUSTOMERVERSION[last()]/@SURNAME ),
																																											'','','',''
																																										)"/></xsl:attribute>
				<xsl:if test=".//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='A'">
					<xsl:element name="APPLICANT">
						<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication( string(.//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select=".//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
						<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( //CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																								string( //CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( //CUSTOMERVERSION[last()]/@SURNAME ),
																																								'','','',''
																																								)"/></xsl:attribute>
					</xsl:element>
				</xsl:if>
				<xsl:if test=".//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='G'">
					<xsl:element name="GUARANTOR">
						<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication( string(.//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select=".//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
						<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( //CUSTOMERVERSION[last()]/@TITLE_OTHER ),
																																								string( //CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( //CUSTOMERVERSION[last()]/@SURNAME ),
																																								'','','',''
																																								)"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																									string( //CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																									)"/></xsl:attribute>
					</xsl:element>
				</xsl:if>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
