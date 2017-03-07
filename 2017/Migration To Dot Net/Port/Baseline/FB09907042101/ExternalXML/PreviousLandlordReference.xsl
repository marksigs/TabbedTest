<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		12/07/2006	EP543			Modified to use 'Other' title if selected
	PE		16/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		25/01/2007	EP2_808		Modified as per new spec
	DS		13/02/2007	EP2_1360	Formatted date
	OS		12/03/2007	EP2_808		Modified to accommodate individual customer selection
	DS		16/03/2007	EP2_808		Corrected faithfully/sincerely logic

	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:apply-templates select="/RESPONSE/PREVIOUSLANDLORDREFERENCE/APPLICATIONFACTFIND/APPLICATIONTENANCY[@DATEMOVEDOUT and TENANCYAPPLICANTS/CUSTOMERVERSION/@CUSTOMERNUMBER]"/>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="APPLICATIONTENANCY">
		<xsl:element name="LETTER">
			
			<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="./@APPLICATIONNUMBER"/>}</xsl:attribute>
			<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication( string( ..//APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES ))"/></xsl:attribute>
			
			<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( //CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																											string( //CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																											string( //CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																											string( //CUSTOMERVERSION[last()]/@SURNAME )
																																											)"/></xsl:attribute>
			<xsl:if test="msg:NotFirstPage()='Y'">
				<xsl:element name="PAGEBREAK"/>
			</xsl:if>
			<xsl:element name="LANDLORDDETAILS">
				<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(	string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																				string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																				string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																				string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME ),
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
				<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
				<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(	string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE_TEXT ),
																																						string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTTITLEOTHER ),
																																						string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME ),
																																						string( .//THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME )
																																						)"/></xsl:attribute>
				<xsl:if test=".//THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME">
					<xsl:attribute name="FAITHFULLYSINCERELY">sincerely</xsl:attribute>
				</xsl:if>
				<xsl:if test="not(.//THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME)">
					<xsl:attribute name="FAITHFULLYSINCERELY">faithfully</xsl:attribute>
				</xsl:if>
			</xsl:element>
			
			<xsl:attribute name="TEST"><xsl:value-of select="//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES"/></xsl:attribute>
			<xsl:if test="msg:CheckForType( string( .//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES), 'A' )">
				<xsl:element name="APPLICANT">
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@SURNAME )
																																								)"/></xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@FLATNUMBER	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@STREET	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@DISTRICT	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@TOWN	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@COUNTY	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@POSTCODE	)
																																								)"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
			
			<xsl:if test="msg:CheckForType( string( .//CUSTOMERROLE/@CUSTOMERROLETYPE_TYPES), 'G' )">
				<xsl:element name="GUARANTOR">
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(	string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																								string( .//TENANCYAPPLICANTS/CUSTOMERVERSION[last()]/@SURNAME )
																																								)"/></xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@FLATNUMBER	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@STREET	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@DISTRICT	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@TOWN	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@COUNTY	),
																																								string(	.//TENANCYAPPLICANTS/CUSTOMERADDRESS/ADDRESS/@POSTCODE	)
																																								)"/></xsl:attribute>
				</xsl:element>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<!--============================================================================================================-->
</xsl:stylesheet>
