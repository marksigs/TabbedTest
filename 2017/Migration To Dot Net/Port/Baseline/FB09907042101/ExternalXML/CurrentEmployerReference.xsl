<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		11/07/2006	EP543			Added 'TitleOther' text for third-party contact details
	PE		01/08/2006	EP103			Salutations incorrect
	PE		03/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		09/02/2007	EP2_806		Modified as per new spec
	DS 		13/02/2007 	EP2_1360 	Formatted date
	PB		15/02/2007	EP2_806		Corrected FAITHFULLY/SINCERELY and address justification
	DS		27/03/2007	EP2_2018		Formatted  customer's / Guarantor's address
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-employer.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="LETTER">
				<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
				<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
				<xsl:attribute name="TYPEOFLOAN"><xsl:value-of select="msg:DealWithTypeOfApplication( string( .//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
				<xsl:attribute name="APPLICANTORGUARANTOR"><xsl:value-of select="msg:ApplicantOrGuarantor( string( .//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES ) )"/></xsl:attribute>
				<xsl:attribute name="NAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLE_TEXT ),
																																							string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLEOTHER ),
																																							string( .//CUSTOMERROLE/CUSTOMERVERSION/@FIRSTFORENAME ),
																																							string( .//CUSTOMERROLE/CUSTOMERVERSION/@SURNAME	)
																																						)"/></xsl:attribute>
				<xsl:element name="EMPLOYER">
					<xsl:if test="//CUSTOMERROLE/EMPLOYMENT/THIRDPARTY/ADDRESS">
						<xsl:attribute name="COMPANYCONTACT"><xsl:value-of select="msg:GetSingleOrJointSalutation2(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME ),
																																														'','','',''
																																													 )"/></xsl:attribute>
						<xsl:attribute name="COMPANYNAME"><xsl:value-of select="string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/@COMPANYNAME )"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:FormatAddress(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@FLATNUMBER ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@STREET ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@DISTRICT ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@TOWN ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@COUNTY ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/ADDRESS/@POSTCODE )
																																				)"/></xsl:attribute>
						<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME )
																																							)"/></xsl:attribute>
						<xsl:if test="//CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME = ''">
							<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
						</xsl:if>
						<xsl:if test="//CUSTOMERROLE/EMPLOYMENT[last()]/THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME != ''">
							<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
						</xsl:if>
					</xsl:if>
					<xsl:if test="//CUSTOMERROLE/EMPLOYMENT/NAMEANDADDRESSDIRECTORY/ADDRESS">
						<xsl:attribute name="COMPANYCONTACT"><xsl:value-of select="msg:GetSingleOrJointSalutation2(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTFORENAME ),
																																														string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME ),
																																														'','','',''
																																													 )"/></xsl:attribute>
						<xsl:attribute name="COMPANYNAME"><xsl:value-of select="string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/@COMPANYNAME )"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:FormatAddress(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@FLATNUMBER ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@STREET ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@TOWN ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTY ),
																																					string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE )
																																				)"/></xsl:attribute>
						<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTFORENAME ),
																																								string( //CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME )
																																							)"/></xsl:attribute>
						<xsl:if test="//CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME = ''">
							<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
						</xsl:if>
						<xsl:if test="//CUSTOMERROLE/EMPLOYMENT[last()]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/@CONTACTSURNAME != ''">
							<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
						</xsl:if>
					</xsl:if>
				</xsl:element>
				<xsl:if test="//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='A'">
					<xsl:element name="APPLICANT">
						<xsl:attribute name="NAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLE_TEXT ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLEOTHER ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@FIRSTFORENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@SURNAME )
																																								)"/></xsl:attribute>
					<xsl:if test=".//CUSTOMERROLE/CUSTOMERADDRESS/@ADDRESSTYPE_TYPES='C'">
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@FLATNUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@STREET ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@DISTRICT ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@TOWN ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@COUNTY ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
					</xsl:if>
					<xsl:if test=".//CUSTOMERROLE/CUSTOMERADDRESS/@ADDRESSTYPE_TYPES='H'">
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
					</xsl:if>
					</xsl:element>
				</xsl:if>
				<xsl:if test=".//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES='G'">
					<xsl:element name="GUARANTOR">
						<xsl:attribute name="NAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLE_TEXT ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@TITLEOTHER ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@FIRSTFORENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERVERSION/@SURNAME )
																																								)"/></xsl:attribute>
					<xsl:if test=".//CUSTOMERROLE/CUSTOMERADDRESS/@ADDRESSTYPE_TYPES='C'">
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@FLATNUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@STREET ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@DISTRICT ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@TOWN ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@COUNTY ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='C']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
					</xsl:if>
					<xsl:if test=".//CUSTOMERROLE/CUSTOMERADDRESS/@ADDRESSTYPE_TYPES='H'">
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY ),
																																									string( .//CUSTOMERROLE/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE )
																																								)"/></xsl:attribute>
					</xsl:if>
					</xsl:element>
				</xsl:if>
			</xsl:element>

			<!--<xsl:apply-templates select="/RESPONSE/CURRENTEMPLOYERREFERENCE/CUSTOMERROLE/EMPLOYMENT[last()]"/>
			<xsl:apply-templates select="/RESPONSE/CURRENTEMPLOYERREFERENCE/CUSTOMERROLE/CUSTOMERVERSION[last()]"/>-->
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
