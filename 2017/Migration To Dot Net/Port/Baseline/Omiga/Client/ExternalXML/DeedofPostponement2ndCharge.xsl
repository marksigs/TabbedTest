<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		30/01/2007	EP2_821		New document 
	DS	 13/02/2007	 EP2_1360	 Formatted date
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="LETTER">
				<xsl:attribute name="DATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
				<xsl:element name="SECONDCHARGECOMPANY">
					<xsl:if test="//VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY">
						<xsl:attribute name="NAME"><xsl:value-of select="//THIRDPARTY/@COMPANYNAME"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:FormatAddress(	string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@FLATNUMBER ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@STREET ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@DISTRICT ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@TOWN ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@COUNTY ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/THIRDPARTY/ADDRESS/@POSTCODE )
																																					)"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="//VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY">
						<xsl:attribute name="NAME"><xsl:value-of select="//VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/@COMPANYNAME"/></xsl:attribute>
						<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:FormatAddress(	string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@FLATNUMBER ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENAME ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENUMBER ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@STREET ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@TOWN ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTY ),
																																					string( //VALIDATEDACCOUNTS[@SECONDCHARGEINDICATOR=1]/ACCOUNT/NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE )
																																					)"/>
						</xsl:attribute>
					</xsl:if>
				</xsl:element>
				<xsl:element name="APPLICANT">
					<xsl:attribute name="NAME"><xsl:value-of select="msg:GetSingleOrJointSalutation(	string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION[last()]/@SURNAME ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLE_TEXT ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@TITLEOTHER ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@FIRSTFORENAME ),
																																							string( //APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION[last()]/@SURNAME )
																																							)"/>
					</xsl:attribute>
					<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(	string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@FLATNUMBER	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENAME	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@BUILDINGORHOUSENUMBER	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@STREET	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@DISTRICT	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@TOWN	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@COUNTY	),
																																								string(	//APPLICATIONFACTFIND/CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERADDRESS[@ADDRESSTYPE_TYPES='H']/ADDRESS/@POSTCODE	)
																																								)"/>
					</xsl:attribute>
				</xsl:element>
				
				
				<xsl:element name="APPLICATION">
					<xsl:attribute name="TYPEOFLOAN"><xsl:value-of select="msg:DealWithTypeOfApplication(	string( //APPLICATIONFACTFIND/@TYPEOFMORTGAGE_TYPES ) )"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="//APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>}</xsl:attribute>
				</xsl:element>
				<xsl:element name="EXISTINGLOAN">
					<xsl:attribute name="AMOUNT"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEBORROWEDAMOUNT"/></xsl:attribute>
					<xsl:attribute name="STARTDATE"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGESTARTDATE"/></xsl:attribute>
					<xsl:attribute name="REMAININGTERM"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGETERM"/></xsl:attribute>
					<xsl:attribute name="BALANCE"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEBALANCE"/></xsl:attribute>
					<xsl:attribute name="PAYMENT"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEMONTHLYPAYMENTAMOUNT"/></xsl:attribute>
					<xsl:attribute name="REPAYMENTTYPE"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEREPAYMENTTYPE_TEXT"/></xsl:attribute>
					<xsl:attribute name="VALUATIONAMOUNT"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEVALUATION"/></xsl:attribute>
					<xsl:attribute name="VALUATIONDATE"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='0']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEVALUATIONDATE"/></xsl:attribute>
				</xsl:element>
				<xsl:element name="ADDITIONALLOAN">
					<xsl:attribute name="AMOUNT"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='1']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEBORROWEDAMOUNT"/></xsl:attribute>
					<xsl:attribute name="TERM"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='1']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGETERM"/></xsl:attribute>
					<xsl:attribute name="PAYMENT"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='1']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEMONTHLYPAYMENTAMOUNT"/></xsl:attribute>
					<xsl:attribute name="REPAYMENTTYPE"><xsl:value-of select="//MORTGAGEACCOUNT[@SECONDCHARGEINDICATOR='1']/MORTGAGEACCOUNTLOANSUMMARY/@EXISTINGMORTGAGEREPAYMENTTYPE_TEXT"/></xsl:attribute>
				</xsl:element>
				
				
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
