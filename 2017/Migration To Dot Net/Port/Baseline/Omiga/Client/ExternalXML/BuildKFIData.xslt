<?xml version="1.0" encoding="UTF-8"?>
<!--  
InterfaceBO_BuildCalcMortCalcRequest.xsl is used by the omInterface component to format a request to
the LTVBO.CalcMortCalcLTV method.

Modification history
_________________________________________________________________________________________________ 
 Date			Author		What has changed?													 
_________________________________________________________________________________________________ 
31/10/2005	GHun		Copied from BBGCode
06/02/2006  	HMA     MAR1190 Add Start Date to MortgageProductLanguage	
10/02/2006	PE	MAR1190 Added GLOBALDATAITEM/@PERCENTAGE and  MORTGAGEPRODUCTLANGUAGE/@LANGUAGE
02/03/2006	BC	MAR1347	Changes for Disposable KFI
06/03/2006	BC	MAR1367	Changes to LoanComponentRedemptionFees for Disposable KFI
22/03/2006	BC	MAR1503	Changes to InterestOnlyElement for Disposable KFI
05/04/2006	bc	mar1571	Add PURCHASEPRICEORESTIMATEDVALUE to MORTGAGESUBQUOTe
08/05/2006	GHun	MAR??? Changed TotalAmountPayable, APR and AmountPerUnitBorrowed for LoanComponent
15/05/2006	BC	MAR1788	Populate MotgageSubQuote.AmountperUnitBorrowed from AlphaPlus element
16/05/2006	BC	MAR1726 Populate TypeofBuyer from input Application node
25/05/2006	BC	MAR1836	Include CAPITALANDINTERESTELEMENT
07/06/2006	BC	MAR1858 Include CAPINTMONTHLYCOST and INTONLYMONTHLYCOST in LOANCOMPONENTPAYMENTSCHEDULE
12/12/2006	GHun	EP2_445 Only get one copy of OneOffCosts
20/02/2007	SR		EP2_1177 - Add NATUREOFLOAN, PRODUCTSCHEME to APPLICATIONFACTFIND and PURPOSEOFLOAN,
								REPAYMENTVEHICLE and REPAYMENTVEHICLEMONTHLYCOST to LOANCOMPONENT
20/03/2007	INR	EP2__1977 Add REGULATIONINDICATOR to APPLICATIONFACTFIND
26/03/2007  INR   EP2_1980 ADDEDTOLOAN should be ADDTOLOAN
29/03/2007	PSC	EP2_1941 Include OTHERNAMES as SECONDFORENAME so they come out on the KFI
06/04/2007  INR   EP2_1994 DIRECTINDIRECTINDICATOR should be indirect for Web
07/04/2007	SR	EP2_1994  Add processing APPLICATIONINTRODUCER
10/04/2007	INR	EP2_1994  Add processing for ProcFee Refund
13/04/2007	INR	EP2_1994  Add processing for ProcFee  Refund
26/03/2007  INR   EP2_2395 Add Fees to loancomponent
16/04/2007 INR	EP2_2448 Add FREELEGALFEES to MortgageProduct
18/04/2007 INR	EP2_2448 Add SPECIALSCHEME to AppFactFind
19/04/2007 INR	EP2_2478 use PRINCIPALFIRMNETWORKID for NETWORK
_________________________________________________________________________________________________ 
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<xsl:variable name="Now" select="/@NOW"/>
	<!-- This variable is put in by the interface -->
	<xsl:variable name="App" select="//APPLICATION"/>
	<xsl:variable name="AppNum" select="$App/@APPLICATIONNUMBER"/>
	<xsl:variable name="AppFactFind" select="$App/APPLICATIONFACTFIND"/>
	<xsl:variable name="QuoteNum" select="$App/APPLICATIONFACTFIND/@INCREMENT"/>
	<xsl:variable name="MSQ" select="$AppFactFind/QUOTATION/MORTGAGESUBQUOTE"/>
	<xsl:variable name="Calcs" select="//CALCS"/>
	<xsl:variable name="AlphaPlus" select="//ALPHAPLUS"/>
	<xsl:variable name="PackageIndicator" select="count(/REQUEST/INTERMEDIARYPACKAGER)"/>
	<xsl:template match="/">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="COMBOLOOKUP">YES</xsl:attribute>
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:attribute name="STATUS"/>
			<!-- APPLICATION Node -->
			<xsl:element name="APPLICATION">
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
				<xsl:attribute name="TYPEOFBUYER_TEXT"/>
				<xsl:attribute name="TYPEOFBUYER"><xsl:value-of select="$App/@TYPEOFBUYER"/></xsl:attribute>
				<!-- APPLICATIONFACTFIND Node -->
				<xsl:element name="APPLICATIONFACTFIND">
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
					<xsl:attribute name="ACCEPTEDQUOTENUMBER"/>
					<xsl:attribute name="ACTIVEDQUOTENUMBER"/>
					<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE"><xsl:value-of select="$AppFactFind/@PURCHASEPRICE"/></xsl:attribute>
					<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="$AppFactFind/@TYPEOFAPPLICATION"/></xsl:attribute>
					<xsl:attribute name="TYPEOFAPPLICATION_TEXT"/>
					<xsl:attribute name="DIRECTINDIRECTBUSINESS">I,</xsl:attribute>
					<xsl:attribute name="DIRECTINDIRECTBUSINESS_TEXT">Web</xsl:attribute>
					<xsl:attribute name="APPLICATIONPACKAGEINDICATOR"><xsl:value-of select="$PackageIndicator"/></xsl:attribute>
					<xsl:attribute name="LEVELOFADVICE"><xsl:value-of select="$AppFactFind/@LEVELOFADVICE"/></xsl:attribute>
					<xsl:attribute name="NATUREOFLOAN"><xsl:value-of select="$AppFactFind/@NATUREOFLOAN"/></xsl:attribute>
					<xsl:attribute name="PRODUCTSCHEME"><xsl:value-of select="$AppFactFind/@PRODUCTSCHEME"/></xsl:attribute>
					<xsl:attribute name="REGULATIONINDICATOR"><xsl:value-of select="$AppFactFind/@REGULATIONINDICATOR"/></xsl:attribute>
					<xsl:attribute name="BROKERPROCFEEREFUND"><xsl:value-of select="$App/@PROCFEETOCUSTAMOUNT"/></xsl:attribute>
					<xsl:attribute name="BROKERPROCTOCUST"><xsl:value-of select="$App/@PROCFEETOCUST"/></xsl:attribute>
					<xsl:attribute name="SPECIALSCHEME"><xsl:value-of select="$AppFactFind/@SPECIALSCHEME"/></xsl:attribute>
					<!--<xsl:attribute name="INTRODUCERFEE"><xsl:value-of select="/REQUEST/QUOTE/@BROKERFEE"/></xsl:attribute>
					<xsl:attribute name="BROKERFEEDETAILS"><xsl:value-of select="$App/@BROKERFEEDETAILS"/></xsl:attribute>
					<xsl:attribute name="INDIRECTTHIRDPARTYPAYMENT"><xsl:value-of select="/REQUEST/APPLICATION/@INDIRECTTHIRDPARTYPAYMENT"/></xsl:attribute>
					<xsl:attribute name="INDIRECTTHIRDPARTY1"><xsl:value-of select="/REQUEST/APPLICATION/@INDIRECTTHIRDPARTY1"/></xsl:attribute>
					<xsl:attribute name="INDIRECTTHIRDPARTY2"><xsl:value-of select="/REQUEST/APPLICATION/@INDIRECTTHIRDPARTY2"/></xsl:attribute>
					<xsl:attribute name="INDIRECTTHIRDPARTY3"><xsl:value-of select="/REQUEST/APPLICATION/@INDIRECTTHIRDPARTY3"/></xsl:attribute>
-->
					<!--APPLICATIONINTRODUCER -->
					<xsl:for-each select=".//APPLICATION/APPLICATIONFACTFIND/APPLICATIONINTRODUCERLIST/APPLICATIONINTRODUCER">
						<xsl:element name="APPLICATIONINTRODUCER">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
							<xsl:attribute name="APPLICATIONINTRODUCERSEQNO"><xsl:value-of select="@APPLICATIONINTRODUCERSEQNO"/></xsl:attribute>
							<xsl:attribute name="PRINCIPALFIRMID"><xsl:value-of select="@PRINCIPALFIRMID"/></xsl:attribute>
							<xsl:attribute name="CLUBNETWORKASSOCID"><xsl:value-of select="@CLUBNETWORKASSOCID"/></xsl:attribute>
							<xsl:attribute name="ARFIRMID"><xsl:value-of select="@ARFIRMID"/></xsl:attribute>
							<xsl:attribute name="INTRODUCERID"><xsl:value-of select="@INTRODUCERID"/></xsl:attribute>
							<xsl:for-each select="PACKAGER">
								<xsl:element name="PRINCIPALFIRM">
									<xsl:attribute name="PRINCIPALFIRMID"><xsl:value-of select="@PACKAGERID"/></xsl:attribute>
									<xsl:attribute name="FSAREF"><xsl:value-of select="@FSAREF"/></xsl:attribute>
									<xsl:choose>
										<xsl:when test="INTRODUCERFIRM/@TRADINGAS">
											<!-- Firmname exists.  -->
											<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="INTRODUCERFIRM/@TRADINGAS"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<!--Firmname does not exist -->
											<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="@PRINCIPALFIRMNAME"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:attribute name="PACKAGERINDICATOR"><xsl:value-of select="@PACKAGERINDICATOR"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE1"><xsl:value-of select="@ADDRESSLINE1"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE2"><xsl:value-of select="@ADDRESSLINE2"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE3"><xsl:value-of select="@ADDRESSLINE3"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE4"><xsl:value-of select="@ADDRESSLINE4"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE5"><xsl:value-of select="@ADDRESSLINE5"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE6"><xsl:value-of select="@ADDRESSLINE6"/></xsl:attribute>
									<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONEAREACODE"><xsl:value-of select="@TELEPHONEAREACODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="@TELEPHONENUMBER"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
							<xsl:for-each select="NETWORK">
								<xsl:element name="PRINCIPALFIRM">
									<xsl:attribute name="PRINCIPALFIRMID"><xsl:value-of select="@PRINCIPALFIRMNETWORKID"/></xsl:attribute>
									<xsl:attribute name="FSAREF"><xsl:value-of select="@FSAREF"/></xsl:attribute>
									<xsl:choose>
										<xsl:when test="INTRODUCERFIRM/@TRADINGAS">
											<!-- Firmname exists.  -->
											<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="INTRODUCERFIRM/@TRADINGAS"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<!--Firmname does not exist -->
											<xsl:attribute name="PRINCIPALFIRMNAME"><xsl:value-of select="@PRINCIPALFIRMNAME"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:attribute name="PACKAGERINDICATOR"><xsl:value-of select="@PACKAGERINDICATOR"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE1"><xsl:value-of select="@ADDRESSLINE1"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE2"><xsl:value-of select="@ADDRESSLINE2"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE3"><xsl:value-of select="@ADDRESSLINE3"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE4"><xsl:value-of select="@ADDRESSLINE4"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE5"><xsl:value-of select="@ADDRESSLINE5"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE6"><xsl:value-of select="@ADDRESSLINE6"/></xsl:attribute>
									<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONEAREACODE"><xsl:value-of select="@TELEPHONEAREACODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="@TELEPHONENUMBER"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
							<xsl:for-each select="ARFIRM">
								<xsl:element name="ARFIRM">
									<xsl:attribute name="ARFIRMID"><xsl:value-of select="@ARFIRMID"/></xsl:attribute>
									<xsl:attribute name="FSAARFIRMREF"><xsl:value-of select="@FSAARFIRMREF"/></xsl:attribute>
									<xsl:choose>
										<xsl:when test="INTRODUCERFIRM/@TRADINGAS">
											<!-- Firmname exists.  -->
											<xsl:attribute name="ARFIRMNAME"><xsl:value-of select="INTRODUCERFIRM/@TRADINGAS"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<!--Firmname does not exist -->
											<xsl:attribute name="ARFIRMNAME"><xsl:value-of select="@ARFIRMNAME"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:attribute name="ADDRESSLINE1"><xsl:value-of select="@ADDRESSLINE1"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE2"><xsl:value-of select="@ADDRESSLINE2"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE3"><xsl:value-of select="@ADDRESSLINE3"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE4"><xsl:value-of select="@ADDRESSLINE4"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE5"><xsl:value-of select="@ADDRESSLINE5"/></xsl:attribute>
									<xsl:attribute name="ADDRESSLINE6"><xsl:value-of select="@ADDRESSLINE6"/></xsl:attribute>
									<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONEAREACODE"><xsl:value-of select="@TELEPHONEAREACODE"/></xsl:attribute>
									<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="@TELEPHONENUMBER"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
							<xsl:for-each select="MORTGAGECLUBNETWORKASSOCIATION">
								<xsl:element name="MORTGAGECLUBNETWORKASSOCIATION">
									<xsl:attribute name="CLUBNETWORKASSOCIATIONID"><xsl:value-of select="@CLUBNETWORKASSOCIATIONID"/></xsl:attribute>
									<xsl:attribute name="MORTGAGECLUBNETWORKASSOCNAME"><xsl:value-of select="@MORTGAGECLUBNETWORKASSOCNAME"/></xsl:attribute>
									<xsl:attribute name="PACKAGERINDICATOR"><xsl:value-of select="@PACKAGERINDICATOR"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
							<xsl:for-each select="INTRODUCER">
								<xsl:element name="INTRODUCER">
									<xsl:attribute name="INTRODUCERID"><xsl:value-of select="@INTRODUCERID"/></xsl:attribute>
									<xsl:attribute name="USERID"><xsl:value-of select="@USERID"/></xsl:attribute>
									<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="@BUILDINGORHOUSENUMBER"/></xsl:attribute>
									<xsl:attribute name="STREET"><xsl:value-of select="@STREET"/></xsl:attribute>
									<xsl:attribute name="DISTRICT"><xsl:value-of select="@DISTRICT"/></xsl:attribute>
									<xsl:attribute name="TOWN"><xsl:value-of select="@TOWN"/></xsl:attribute>
									<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
									<xsl:for-each select="ORGANISATIONUSER">
										<xsl:element name="ORGANISATIONUSER">
											<xsl:attribute name="USERTITLE"><xsl:value-of select="@USERTITLE"/></xsl:attribute>
											<xsl:attribute name="USERFORENAME"><xsl:value-of select="@USERFORENAME"/></xsl:attribute>
											<xsl:attribute name="USERSURNAME"><xsl:value-of select="@USERSURNAME"/></xsl:attribute>
											<xsl:attribute name="USERID"><xsl:value-of select="@USERID"/></xsl:attribute>
											<xsl:for-each select="ORGANISATIONUSERCONTACTDETAILS">
												<xsl:element name="ORGANISATIONUSERCONTACTDETAILS">
													<xsl:attribute name="USERID"><xsl:value-of select="@USERID"/></xsl:attribute>
													<xsl:for-each select="CONTACTTELEPHONEDETAILS">
														<xsl:element name="CONTACTTELEPHONEDETAILS">
															<xsl:attribute name="TELEPHONESEQNUM"><xsl:value-of select="@TELEPHONESEQNUM"/></xsl:attribute>
															<xsl:attribute name="USAGE"><xsl:value-of select="@USAGE"/></xsl:attribute>
															<xsl:attribute name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:value-of select="@TELENUMBER"/></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:attribute>
														</xsl:element>
													</xsl:for-each>
												</xsl:element>
											</xsl:for-each>
										</xsl:element>
									</xsl:for-each>
								</xsl:element>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
					<!-- QUOTATION Node -->
					<xsl:element name="QUOTATION">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
						<xsl:attribute name="MORTGAGESUBQUOTENUMBER"/>
						<xsl:attribute name="QUOTATIONNUMBER"><xsl:value-of select="$QuoteNum"/></xsl:attribute>
						<xsl:attribute name="DATEANDTIMEGENERATED"><xsl:value-of select="$Now"/></xsl:attribute>
						<!-- MORTGAGESUBQUOTE Node -->
						<xsl:element name="MORTGAGESUBQUOTE">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
							<xsl:attribute name="MORTGAGESUBQUOTENUMBER"/>
							<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="$MSQ/@AMOUNTREQUESTED"/></xsl:attribute>
							<xsl:attribute name="TOTALLOANAMOUNT"><xsl:value-of select="/REQUEST/QUOTE/@TOTALLOANAMOUNT"/></xsl:attribute>
							<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE"><xsl:value-of select="$Calcs/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/PURCHASEPRICEORESTIMATEDVALUE"/></xsl:attribute>
							<xsl:attribute name="LTV"><xsl:value-of select="$Calcs/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LTV"/></xsl:attribute>
							<xsl:attribute name="APR"><xsl:value-of select="$AlphaPlus/APR"/></xsl:attribute>
							<xsl:attribute name="AMOUNTPERUNITBORROWED"><xsl:value-of select="$AlphaPlus/AMOUNTPERUNITBORROWED"/></xsl:attribute>
							<xsl:attribute name="DRAWDOWN"/>
							<xsl:for-each select="/ONEOFFCOSTLIST/ONEOFFCOST">
								<xsl:element name="MORTGAGEONEOFFCOST">
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
									<xsl:attribute name="APPLICATIONFACTFINDNUMBER"/>
									<xsl:attribute name="MORTGAGESUBQUOTENUMBER"/>
									<xsl:attribute name="MORTGAGEONEOFFCOSTTYPE_TEXT"/>
									<xsl:attribute name="MORTGAGEONEOFFCOSTTYPE"><xsl:value-of select="IDENTIFIER"/></xsl:attribute>
									<xsl:attribute name="AMOUNT"><xsl:value-of select="AMOUNT"/></xsl:attribute>
									<xsl:if test="ADDTOLOAN='1'">
										<xsl:attribute name="ADDTOLOAN"><xsl:value-of select="ADDTOLOAN"/></xsl:attribute>
									</xsl:if>
								</xsl:element>
							</xsl:for-each>
							<xsl:for-each select="//CALCS/MORTGAGEINTRODUCERFEELIST/MORTGAGEINTRODUCERFEE">
								<xsl:element name="MORTGAGEINTRODUCERFEE">
									<xsl:attribute name="INTRODUCERFEETYPE"><xsl:value-of select="@INTRODUCERFEETYPE"/></xsl:attribute>
									<xsl:attribute name="INTRODUCERFEETYPE_TEXT"><xsl:value-of select="@INTRODUCERFEETYPE_TEXT"/></xsl:attribute>
									<xsl:attribute name="INTRODUCERFEETYPE_VALIDID"><xsl:value-of select="@INTRODUCERFEETYPE_VALIDID"/></xsl:attribute>
									<xsl:attribute name="FEEAMOUNT"><xsl:value-of select="@FEEAMOUNT"/></xsl:attribute>
									<xsl:attribute name="RECIPIENTID"><xsl:value-of select="@RECIPIENTID"/></xsl:attribute>
								</xsl:element>
							</xsl:for-each>
							<!-- LOANCOMPONENT Nodes -->
							<xsl:for-each select="$Calcs/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT">
								<xsl:variable name="LCPosition" select="position()"/>
								<xsl:variable name="dblLTV" select="$Calcs/APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LTV"/>
								<xsl:element name="LOANCOMPONENT">
									<xsl:attribute name="RESOLVEDRATE"><xsl:value-of select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCPosition]/RESOLVEDRATE"/></xsl:attribute>
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
									<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
									<xsl:attribute name="MORTGAGESUBQUOTENUMBER"/>
									<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="$LCPosition"/></xsl:attribute>
									<xsl:variable name="PackageIndicator" select="count(/REQUEST/INTERMEDIARYPACKAGER)"/>
									<xsl:attribute name="APR"><xsl:value-of select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCPosition]/APR"/></xsl:attribute>
									<xsl:attribute name="CAPITALANDINTERESTELEMENT"><xsl:value-of select="CAPITALANDINTERESTELEMENT"/></xsl:attribute>
									<xsl:attribute name="GROSSMONTHLYCOST"><xsl:value-of select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCPosition]/GROSSMONTHLYCOST"/></xsl:attribute>
									<xsl:attribute name="INTERESTONLYELEMENT"><xsl:value-of select="INTERESTONLYELEMENT"/></xsl:attribute>
									<xsl:attribute name="LOANAMOUNT"><xsl:value-of select="TOTALLOANCOMPONENTAMOUNT"/></xsl:attribute>
									<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="MORTGAGEPRODUCTCODE"/></xsl:attribute>
									<xsl:attribute name="REPAYMENTMETHOD_TEXT"/>
									<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
									<xsl:attribute name="TERMINMONTHS"><xsl:value-of select="TERMINMONTHS"/></xsl:attribute>
									<xsl:attribute name="TERMINYEARS"><xsl:value-of select="TERMINYEARS"/></xsl:attribute>
									<xsl:attribute name="TOTALAMOUNTPAYABLE"><xsl:value-of select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCPosition]/TOTALAMOUNTPAYABLE"/></xsl:attribute>
									<xsl:choose>
										<xsl:when test="$LCPosition='1'">
											<!-- first LC.  Add on any Fees-->
											<xsl:variable name="FeeAmountAddToLoan" select="sum(//ONEOFFCOSTLIST/ONEOFFCOST[@ADDTOLOAN='1']/@AMOUNT)"/>
											<xsl:variable name="LoanAmount" select="TOTALLOANCOMPONENTAMOUNT"/>
											<xsl:attribute name="TOTALLOANCOMPONENTAMOUNT"><xsl:value-of select="$FeeAmountAddToLoan + $LoanAmount"/></xsl:attribute>
										</xsl:when>
										<xsl:otherwise>
											<xsl:attribute name="TOTALLOANCOMPONENTAMOUNT"><xsl:value-of select="TOTALLOANCOMPONENTAMOUNT"/></xsl:attribute>
										</xsl:otherwise>
									</xsl:choose>
									<xsl:attribute name="INCREASEDMONTHLYCOST"/>
									<xsl:attribute name="INCREASEDMONTHLYCOSTDIFFERENCE"/>
									<xsl:attribute name="REPAYMENTMETHOD"><xsl:value-of select="REPAYMENTMETHOD"/></xsl:attribute>
									<xsl:attribute name="FINALBALANCEPLUSONEPERCENT"/>
									<xsl:attribute name="AMOUNTPERUNITBORROWED"><xsl:value-of select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENT[position()=$LCPosition]/AMOUNTPERUNITBORROWED"/></xsl:attribute>
									<xsl:attribute name="PURPOSEOFLOAN"><xsl:value-of select="$AppFactFind/@TYPEOFAPPLICATION"/></xsl:attribute>
									<xsl:attribute name="REPAYMENTVEHICLE"><xsl:value-of select="REPAYMENTVEHICLE"/></xsl:attribute>
									<xsl:attribute name="REPAYMENTVEHICLEMONTHLYCOST"><xsl:value-of select="REPAYMENTVEHICLEMONTHLYCOST"/></xsl:attribute>
									<!-- LOANCOMPONENTBALANCESCHEDULE Node -->
									<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTBALANCESCHEDULE[LOANCOMPONENTSEQUENCENUMBER=$LCPosition]">
										<xsl:element name="LOANCOMPONENTBALANCESCHEDULE">
											<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="APPLICATIONNUMBER"/></xsl:attribute>
											<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="APPLICATIONFACTFINDNUMBER"/></xsl:attribute>
											<xsl:attribute name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="MORTGAGESUBQUOTENUMBER"/></xsl:attribute>
											<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="LOANCOMPONENTSEQUENCENUMBER"/></xsl:attribute>
											<xsl:attribute name="BALANCE"><xsl:value-of select="BALANCE"/></xsl:attribute>
											<xsl:attribute name="INTONLYBALANCE"><xsl:value-of select="INTONLYBALANCE"/></xsl:attribute>
											<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
										</xsl:element>
									</xsl:for-each>
									<!-- LOANCOMPONENTPAYMENTSCHEDULE Node -->
									<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTPAYMENTSCHEDULE[LOANCOMPONENTSEQUENCENUMBER=$LCPosition]">
										<xsl:element name="LOANCOMPONENTPAYMENTSCHEDULE">
											<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="APPLICATIONNUMBER"/></xsl:attribute>
											<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="APPLICATIONFACTFINDNUMBER"/></xsl:attribute>
											<xsl:attribute name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="MORTGAGESUBQUOTENUMBER"/></xsl:attribute>
											<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="LOANCOMPONENTSEQUENCENUMBER"/></xsl:attribute>
											<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
											<xsl:attribute name="MONTHLYCOST"><xsl:value-of select="MONTHLYCOST"/></xsl:attribute>
											<xsl:attribute name="INCREASEDMONTHLYCOSTDIFFERENCE"><xsl:value-of select="INCREASEDMONTHLYCOSTDIFFERENCE"/></xsl:attribute>
											<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
											<xsl:attribute name="INTERESTRATE"><xsl:value-of select="INTERESTRATE"/></xsl:attribute>
											<xsl:attribute name="CAPINTMONTHLYCOST"><xsl:value-of select="CAPINTMONTHLYCOST"/></xsl:attribute>
											<xsl:attribute name="INTONLYMONTHLYCOST"><xsl:value-of select="INTONLYMONTHLYCOST"/></xsl:attribute>
										</xsl:element>
									</xsl:for-each>
									<!-- LOANCOMPONENTREDEMPTIONFEE Node -->
									<xsl:for-each select="$AlphaPlus/LOANCOMPONENTLIST/LOANCOMPONENTREDEMPTIONFEE[LOANCOMPONENTSEQUENCENUMBER=$LCPosition]">
										<xsl:element name="LOANCOMPONENTREDEMPTIONFEE">
											<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="APPLICATIONNUMBER"/></xsl:attribute>
											<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="APPLICATIONFACTFINDNUMBER"/></xsl:attribute>
											<xsl:attribute name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="MORTGAGESUBQUOTENUMBER"/></xsl:attribute>
											<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER"><xsl:value-of select="LOANCOMPONENTSEQUENCENUMBER"/></xsl:attribute>
											<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
											<xsl:attribute name="REDEMPTIONFEESTEPNUMBER"><xsl:value-of select="REDEMPTIONFEESTEPNUMBER"/></xsl:attribute>
											<xsl:attribute name="REDEMPTIONFEEAMOUNT"><xsl:value-of select="REDEMPTIONFEEAMOUNT"/></xsl:attribute>
											<xsl:attribute name="REDEMPTIONFEEPERIOD"><xsl:value-of select="REDEMPTIONFEEPERIOD"/></xsl:attribute>
										</xsl:element>
									</xsl:for-each>
									<!-- MORTGAGEPRODUCT Node -->
									<xsl:element name="MORTGAGEPRODUCT">
										<xsl:variable name="MP" select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT"/>
										<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="$MP/MORTGAGEPRODUCTCODE"/></xsl:attribute>
										<xsl:attribute name="STARTDATE"><xsl:value-of select="$MP/STARTDATE"/></xsl:attribute>
										<xsl:attribute name="CATIND"><xsl:value-of select="$MP/CATIND"/></xsl:attribute>
										<xsl:attribute name="FLEXIBLEMORTGAGEPRODUCT"><xsl:value-of select="$MP/FLEXIBLEMORTGAGEPRODUCT"/></xsl:attribute>
										<xsl:attribute name="IMPAIREDCREDIT"><xsl:value-of select="$MP/IMPAIREDCREDIT"/></xsl:attribute>
										<xsl:attribute name="MAXIMUMLTV"><xsl:value-of select="$MP/MAXIMUMLTV"/></xsl:attribute>
										<xsl:attribute name="MORTGAGEPRODUCTPORTABLEIND"><xsl:value-of select="$MP/MORTGAGEPRODUCTPORTABLEIND"/></xsl:attribute>
										<xsl:attribute name="ORGANISATIONID"><xsl:value-of select="MORTGAGEPRODUCTDETAILS/MORTGAGELENDER/ORGANISATIONID"/></xsl:attribute>
										<xsl:attribute name="FREELEGALFEES"><xsl:value-of select="$MP/FREELEGALFEES"/></xsl:attribute>
										<!--<xsl:attribute name="FLEXIBLEOPTION"><xsl:value-of select="$MP/FLEXIBLEOPTION"/></xsl:attribute>-->
										<!-- MORTGAGEPRODUCTEMPLOYMENT Node -->
										<xsl:for-each select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCTEMPLOYMENTLIST/MORTGAGEPRODUCTEMPLOYMENT">
											<xsl:element name="MORTGAGEPRODUCTEMPLOYMENT">
												<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="MORTGAGEPRODUCTCODE"/></xsl:attribute>
												<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
												<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="EMPLOYMENTSTATUS"/></xsl:attribute>
												<xsl:attribute name="EMPLOYMENTSTATUS_TEXT"/>
											</xsl:element>
										</xsl:for-each>
										<!-- MORTGAGEPRODUCTLANGUAGE Node -->
										<xsl:element name="MORTGAGEPRODUCTLANGUAGE">
											<xsl:variable name="MPL" select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCTLANGUAGE"/>
											<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="$MPL/MORTGAGEPRODUCTCODE"/></xsl:attribute>
											<xsl:attribute name="STARTDATE"><xsl:value-of select="$MPL/STARTDATE"/></xsl:attribute>
											<xsl:attribute name="LANGUAGE"><xsl:value-of select="$MPL/LANGUAGE"/></xsl:attribute>
											<xsl:attribute name="PRODUCTNAME"><xsl:value-of select="$MPL/PRODUCTNAME"/></xsl:attribute>
											<xsl:attribute name="PRODUCTTEXTDETAILS"><xsl:value-of select="$MPL/PRODUCTTEXTDETAILS"/></xsl:attribute>
										</xsl:element>
										<!-- MORTGAGEPRODUCTPARAMETERS Node -->
										<xsl:element name="MORTGAGEPRODUCTPARAMETERS">
											<xsl:variable name="MPP" select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCTPARAMETERS"/>
											<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="$MPP/MORTGAGEPRODUCTCODE"/></xsl:attribute>
											<xsl:attribute name="COMPULSORYBC"><xsl:value-of select="$MPP/COMPULSORYBC"/></xsl:attribute>
											<xsl:attribute name="COMPULSORYPP"><xsl:value-of select="$MPP/COMPULSORYPP"/></xsl:attribute>
										</xsl:element>
										<!-- MORTGAGEPRODUCTPROPLOCATION Node -->
										<xsl:for-each select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCTPROPLOCATIONLIST/MORTGAGEPRODUCTPROPLOCATION">
											<xsl:element name="MORTGAGEPRODUCTPROPLOCATION">
												<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="MORTGAGEPRODUCTCODE"/></xsl:attribute>
												<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
												<xsl:attribute name="MORTGAGEPRODUCTLOCATIONTYPE"><xsl:value-of select="MORTGAGEPRODUCTLOCATIONTYPE"/></xsl:attribute>
												<xsl:attribute name="MORTGAGEPRODUCTLOCATIONTYPE_TEXT"/>
											</xsl:element>
										</xsl:for-each>
										<!-- MORTGAGEPRODUCTTYPEOFBUYER Node -->
										<xsl:for-each select="MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCTTYPEOFBUYERLIST/MORTGAGEPRODUCTTYPEOFBUYER">
											<xsl:element name="MORTGAGEPRODUCTTYPEOFBUYER">
												<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="MORTGAGEPRODUCTCODE"/></xsl:attribute>
												<xsl:attribute name="STARTDATE"><xsl:value-of select="STARTDATE"/></xsl:attribute>
												<xsl:attribute name="TYPEOFBUYER"><xsl:value-of select="TYPEOFBUYER"/></xsl:attribute>
												<xsl:attribute name="TYPEOFBUYER_TEXT"/>
											</xsl:element>
										</xsl:for-each>
										<!-- MORTGAGELENDER Node -->
										<xsl:element name="MORTGAGELENDER">
											<xsl:variable name="ML" select="MORTGAGEPRODUCTDETAILS/MORTGAGELENDER"/>
											<xsl:attribute name="ORGANISATIONID"><xsl:value-of select="$ML/ORGANISATIONID"/></xsl:attribute>
											<xsl:attribute name="LENDERNAME"><xsl:value-of select="$ML/LENDERNAME"/></xsl:attribute>
											<!-- MORTGAGELENDERDIRECTORY Node -->
											<xsl:element name="MORTGAGELENDERDIRECTORY">
												<xsl:attribute name="ORGANISATIONID"><xsl:value-of select="//MORTGAGELENDERDIRECTORY/ORGANISATIONID"/></xsl:attribute>
												<xsl:attribute name="DIRECTORYGUID"><xsl:value-of select="//MORTGAGELENDERDIRECTORY/DIRECTORYGUID"/></xsl:attribute>
												<!-- NAMEANDADDRESSDIRECTORY Node -->
												<xsl:element name="NAMEANDADDRESSDIRECTORY">
													<xsl:attribute name="DIRECTORYGUID"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/DIRECTORYGUID"/></xsl:attribute>
													<xsl:attribute name="COMPANYNAME"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/COMPANYNAME"/></xsl:attribute>
													<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/DIRECTORYGUID"/></xsl:attribute>
													<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESSGUID"/></xsl:attribute>
													<!-- CONTACTDETAILS Node -->
													<xsl:element name="CONTACTDETAILS">
														<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTDETAILSGUID"/></xsl:attribute>
														<xsl:attribute name="CONTACTFORENAME"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTFORENAME"/></xsl:attribute>
														<xsl:attribute name="CONTACTSURNAME"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTSURNAME"/></xsl:attribute>
														<xsl:attribute name="CONTACTTITLE"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/CONTACTDETAILS/CONTACTTITLE"/></xsl:attribute>
														<xsl:attribute name="CONTACTTITLE_TEXT"/>
													</xsl:element>
													<xsl:element name="ADDRESS">
														<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/ADDRESSGUID"/></xsl:attribute>
														<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/BUILDINGORHOUSENUMBER"/></xsl:attribute>
														<xsl:attribute name="STREET"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/STREET"/></xsl:attribute>
														<xsl:attribute name="DISTRICT"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT"/></xsl:attribute>
														<xsl:attribute name="TOWN"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/TOWN"/></xsl:attribute>
														<xsl:attribute name="COUNTY"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/COUNTY"/></xsl:attribute>
														<xsl:attribute name="COUNTRY"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/COUNTRY"/></xsl:attribute>
														<xsl:attribute name="COUNTRY_TEXT"/>
														<xsl:attribute name="POSTCODE"><xsl:value-of select="//NAMEANDADDRESSDIRECTORY/ADDRESS/POSTCODE"/></xsl:attribute>
													</xsl:element>
													<xsl:for-each select="//NAMEANDADDRESSDIRECTORY/CONTACTTELEPHONEDETAILS">
														<xsl:element name="CONTACTTELEPHONEDETAILS">
															<xsl:attribute name="USAGE"><xsl:value-of select="USAGE"/></xsl:attribute>
															<xsl:attribute name="COUNTRYCODE"><xsl:value-of select="COUNTRYCODE"/></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:value-of select="AREACODE"/></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:value-of select="TELENUMBER"/></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="EXTENSIONNUMBER"/></xsl:attribute>
														</xsl:element>
														<!--CONTACTTELEPHONEDETAILS -->
													</xsl:for-each>
												</xsl:element>
												<!-- NAMEANDADDRESSDIRECTORY Node -->
											</xsl:element>
											<!-- MORTGAGELENDERDIRECTORY Node -->
										</xsl:element>
										<!-- MORTGAGELENDER Node -->
										<!-- INTERESTRATETYPE Node -->
										<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE">
											<xsl:element name="INTERESTRATETYPE">
												<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="MORTGAGEPRODUCTCODE"/></xsl:attribute>
												<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
												<xsl:attribute name="INTERESTRATEPERIOD"><xsl:value-of select="INTERESTRATEPERIOD"/></xsl:attribute>
												<xsl:attribute name="INTERESTRATEENDDATE"><xsl:value-of select="INTERESTRATEENDDATE"/></xsl:attribute>
												<xsl:choose>
													<xsl:when test="RATETYPE ='B'">
														<xsl:attribute name="RATE"><xsl:value-of select="//BASERATESET/BASERATE/BASEINTERESTRATE"/></xsl:attribute>
													</xsl:when>
													<xsl:otherwise>
														<xsl:attribute name="RATE"><xsl:value-of select="RATE"/></xsl:attribute>
													</xsl:otherwise>
												</xsl:choose>
												<xsl:attribute name="RATETYPE"><xsl:value-of select="RATETYPE"/></xsl:attribute>
												<xsl:attribute name="BASERATESET"><xsl:value-of select="BASERATESET"/></xsl:attribute>
												<!-- BASERATESETDATA Node -->
												<xsl:element name="BASERATESETDATA">
													<xsl:attribute name="BASERATESET"><xsl:value-of select="BASERATEBAND/BASERATESET"/></xsl:attribute>
													<xsl:attribute name="RATEID"><xsl:value-of select="BASERATEBAND/RATEID"/></xsl:attribute>
													<xsl:attribute name="BASERATESETDESCRIPTION"><xsl:value-of select="BASERATEBAND/BASERATESETDESCRIPTION"/></xsl:attribute>
													<!-- BASERATE Node -->
													<xsl:element name="BASERATE">
														<xsl:attribute name="RATEID"><xsl:value-of select="BASERATEBAND/BASERATERATEID"/></xsl:attribute>
														<xsl:attribute name="BASERATESTARTDATE"><xsl:value-of select="BASERATEBAND/BASERATESTARTDATE"/></xsl:attribute>
														<!--	<xsl:attribute name="BASEINTERESTRATE"><xsl:value-of select="BASERATEBAND/RATE/@RAW"/></xsl:attribute> -->
														<xsl:attribute name="BASEINTERESTRATE"><xsl:value-of select="format-number(number(BASERATEBAND/BASERATEBASEINTERESTRATEID),'#.##')"/></xsl:attribute>
														<xsl:attribute name="RATETYPE"><xsl:value-of select="BASERATEBAND/BASERATERATETYPE"/></xsl:attribute>
														<xsl:attribute name="RATEDESCRIPTION"><xsl:value-of select="BASERATEBAND/RATEDESCRIPTION"/></xsl:attribute>
														<!-- <xsl:variable name="dtBRBSDate" select="BASERATEBANDSTARTDATE"/> -->
													</xsl:element>
													<xsl:for-each select="BASERATEBAND">
														<xsl:element name="BASERATEBAND">
															<!-- <xsl:variable name="dtBRBSDate" select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPE/BASERATEBANDSTARTDATE"/> -->
															<xsl:attribute name="BASERATESET"><xsl:value-of select="BASERATESET"/></xsl:attribute>
															<!-- <xsl:attribute name="BASERATEBANDSTARTDATE"><xsl:value-of select="$dtBRBSDate"/></xsl:attribute>-->
															<xsl:attribute name="BASERATEBANDSTARTDATE"><xsl:value-of select="BASERATEBANDSTARTDATE"/></xsl:attribute>
															<xsl:attribute name="MAXIMUMLOANAMOUNT"><xsl:value-of select="MAXIMUMLOANAMOUNT"/></xsl:attribute>
															<xsl:attribute name="MAXIMUMLTV"><xsl:value-of select="MAXIMUMLTV"/></xsl:attribute>
															<xsl:attribute name="RATEDIFFERENCE"><xsl:value-of select="format-number(number(RATEDIFFERENCE),'#.##')"/></xsl:attribute>
														</xsl:element>
													</xsl:for-each>
												</xsl:element>
											</xsl:element>
										</xsl:for-each>
										<!-- INTERESTRATETYPE Node -->
										<xsl:for-each select="MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND">
											<xsl:element name="REDEMPTIONFEEBAND">
												<xsl:attribute name="REDEMPTIONFEESET"><xsl:value-of select="MREDEMPTIONFEESET"/></xsl:attribute>
												<xsl:attribute name="REDEMPTIONFEESTARTDATE"><xsl:value-of select="REDEMPTIONFEESTARTDATE"/></xsl:attribute>
												<xsl:attribute name="REDEMPTIONFEESTEPNUMBER"><xsl:value-of select="REDEMPTIONFEESTEPNUMBER"/></xsl:attribute>
												<xsl:attribute name="FEEMONTHSINTEREST"><xsl:value-of select="FEEMONTHSINTEREST"/></xsl:attribute>
												<xsl:attribute name="PERIOD"><xsl:value-of select="PERIOD"/></xsl:attribute>
												<xsl:attribute name="PERIODENDDATE"><xsl:value-of select="PERIODENDDATE"/></xsl:attribute>
												<xsl:attribute name="FEEPERCENTAGE"><xsl:value-of select="FEEPERCENTAGE"/></xsl:attribute>
												<xsl:attribute name="LASTCHANGEDATE"><xsl:value-of select="LASTCHANGEDATE"/></xsl:attribute>
											</xsl:element>
										</xsl:for-each>
									</xsl:element>
									<!-- MORTGAGEPRODUCT Node -->
								</xsl:element>
							</xsl:for-each>
							<!-- LOANCOMPONENT Nodes -->
						</xsl:element>
						<!-- MORTGAGESUBQUOTE Node -->
					</xsl:element>
					<!-- QUOTATION Node -->
					<xsl:for-each select="//APPLICATIONINTERMEDIARY">
						<xsl:element name="APPLICATIONINTERMEDIARY">
							<xsl:attribute name="INTERMEDIARYTYPE"><xsl:value-of select="INTERMEDIARY/@INTERMEDIARYTYPE"/></xsl:attribute>
							<xsl:choose>
								<xsl:when test="@COMPANYNAME">
									<!-- Company name exists. So INTERMEDIARYORGANISATION -->
									<xsl:element name="INTERMEDIARY">
										<xsl:element name="ADDRESS">
											<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="INTERMEDIARY/ADDRESS/@ADDRESSGUID"/></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="INTERMEDIARY/ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="INTERMEDIARY/ADDRESS/@BUILDINGORHOUSENAME"/></xsl:attribute>
											<xsl:attribute name="STREET"><xsl:value-of select="INTERMEDIARY/ADDRESS/@STREET"/></xsl:attribute>
											<xsl:attribute name="DISTRICT"><xsl:value-of select="INTERMEDIARY/ADDRESS/@DISTRICT"/></xsl:attribute>
											<xsl:attribute name="TOWN"><xsl:value-of select="INTERMEDIARY/ADDRESS/@TOWN"/></xsl:attribute>
											<xsl:attribute name="COUNTY"><xsl:value-of select="INTERMEDIARY/ADDRESS/@COUNTY"/></xsl:attribute>
											<xsl:attribute name="COUNTRY"><xsl:value-of select="INTERMEDIARY/ADDRESS/@COUNTRY"/></xsl:attribute>
											<xsl:attribute name="COUNTRY_TEXT"/>
											<xsl:attribute name="POSTCODE"><xsl:value-of select="INTERMEDIARY/ADDRESS/@POSTCODE"/></xsl:attribute>
										</xsl:element>
										<xsl:element name="INTERMEDIARYORGANISATION">
											<xsl:attribute name="NAME"><xsl:value-of select="@COMPANYNAME"/></xsl:attribute>
											<xsl:element name="INTERMEDIARYCONTACT">
												<xsl:for-each select="INTERMEDIARY/ADDRESS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS">
													<xsl:element name="CONTACTTELEPHONEDETAILS">
														<xsl:attribute name="USAGE"><xsl:value-of select="@USAGE"/></xsl:attribute>
														<xsl:attribute name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:attribute>
														<xsl:attribute name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:attribute>
														<xsl:attribute name="TELENUMBER"><xsl:value-of select="@TELENUMBER"/></xsl:attribute>
														<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:attribute>
													</xsl:element>
												</xsl:for-each>
											</xsl:element>
										</xsl:element>
									</xsl:element>
								</xsl:when>
								<xsl:otherwise>
									<!-- Company name does not exist, so INTERMEDIARYINDIVIDUAL -->
									<xsl:element name="INTERMEDIARY">
										<xsl:element name="ADDRESS">
											<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="INTERMEDIARY/ADDRESS/@ADDRESSGUID"/></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="INTERMEDIARY/ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="INTERMEDIARY/ADDRESS/@BUILDINGORHOUSENAME"/></xsl:attribute>
											<xsl:attribute name="STREET"><xsl:value-of select="INTERMEDIARY/ADDRESS/@STREET"/></xsl:attribute>
											<xsl:attribute name="DISTRICT"><xsl:value-of select="INTERMEDIARY/ADDRESS/@DISTRICT"/></xsl:attribute>
											<xsl:attribute name="TOWN"><xsl:value-of select="INTERMEDIARY/ADDRESS/@TOWN"/></xsl:attribute>
											<xsl:attribute name="COUNTY"><xsl:value-of select="INTERMEDIARY/ADDRESS/@COUNTY"/></xsl:attribute>
											<xsl:attribute name="COUNTRY"><xsl:value-of select="INTERMEDIARY/ADDRESS/@COUNTRY"/></xsl:attribute>
											<xsl:attribute name="COUNTRY_TEXT"/>
											<xsl:attribute name="POSTCODE"><xsl:value-of select="INTERMEDIARY/ADDRESS/@POSTCODE"/></xsl:attribute>
										</xsl:element>
										<xsl:element name="INTERMEDIARYINDIVIDUAL">
											<xsl:attribute name="FORENAME"><xsl:value-of select="@FORENAME"/></xsl:attribute>
											<xsl:attribute name="SURNAME"><xsl:value-of select="@SURNAME"/></xsl:attribute>
											<xsl:attribute name="TITLE"><xsl:value-of select="@TITLE"/></xsl:attribute>
											<xsl:attribute name="TITLE_TEXT"/>
											<xsl:element name="INTERMEDIARYCONTACT">
												<xsl:for-each select="INTERMEDIARY/ADDRESS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS">
													<xsl:element name="CONTACTTELEPHONEDETAILS">
														<xsl:attribute name="USAGE"><xsl:value-of select="@USAGE"/></xsl:attribute>
														<xsl:attribute name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:attribute>
														<xsl:attribute name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:attribute>
														<xsl:attribute name="TELENUMBER"><xsl:value-of select="@TELENUMBER"/></xsl:attribute>
														<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:attribute>
													</xsl:element>
												</xsl:for-each>
											</xsl:element>
										</xsl:element>
									</xsl:element>
								</xsl:otherwise>
								<!-- Company name does not exist -->
							</xsl:choose>
						</xsl:element>
						<!-- APPLICATIONINTERMEDIARY Node -->
					</xsl:for-each>
					<!-- CUSTOMERROLE Node -->
					<xsl:for-each select="//CUSTOMER">
						<xsl:element name="CUSTOMERROLE">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNum"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER">1</xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"/>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER">1</xsl:attribute>
							<xsl:attribute name="CUSTOMERROLETYPE">1</xsl:attribute>
							<xsl:attribute name="CUSTOMERROLETYPE_TEXT">Applicant</xsl:attribute>
							<!-- CUSTOMERVERSION Node -->
							<xsl:element name="CUSTOMERVERSION">
								<xsl:attribute name="CUSTOMERNUMBER"/>
								<xsl:attribute name="CUSTOMERVERSIONNUMBER">1</xsl:attribute>
								<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="@FIRSTNAME"/></xsl:attribute>
								<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="@OTHERNAMES"/></xsl:attribute>
								<xsl:attribute name="SURNAME"><xsl:value-of select="@SURNAME"/></xsl:attribute>
								<xsl:attribute name="TITLE"><xsl:value-of select="@TITLE"/></xsl:attribute>
								<xsl:attribute name="TITLE_TEXT"/>
								<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="@DATEOFBIRTH"/></xsl:attribute>
								<xsl:variable name="dtDOB" select="@DATEOFBIRTH"/>
								<xsl:variable name="d" select="substring($dtDOB, 1,2)"/>
								<xsl:variable name="m" select="substring($dtDOB,4,2)"/>
								<xsl:variable name="y" select="substring($dtDOB,7,4)"/>
								<xsl:variable name="nDOB" select="(($y * 10000) + ($m * 100) + $d)"/>
								<xsl:variable name="curd" select="substring($Now, 1,2)"/>
								<xsl:variable name="curm" select="substring($Now,4,2)"/>
								<xsl:variable name="cury" select="substring($Now,7,4)"/>
								<xsl:variable name="nNow" select="(($cury * 10000) + ($curm * 100) + $curd)"/>
								<xsl:variable name="age" select="(($nNow - $nDOB) div 10000)"/>
								<xsl:attribute name="AGE"><xsl:value-of select="format-number(floor($age), '###')"/></xsl:attribute>
							</xsl:element>
						</xsl:element>
					</xsl:for-each>
					<xsl:element name="NEWPROPERTY">
						<xsl:attribute name="VALUATIONTYPE"/>
					</xsl:element>
				</xsl:element>
				<!-- APPLICATIONFACTFIND Node -->
			</xsl:element>
			<!-- APPLICATION Node -->
			<xsl:for-each select="//GLOBALPARAMETERLIST/GLOBALPARAMETER">
				<xsl:element name="GLOBALDATAITEM">
					<xsl:attribute name="NAME"><xsl:value-of select="NAME"/></xsl:attribute>
					<xsl:attribute name="GLOBALPARAMETERSTARTDATE"><xsl:value-of select="GLOBALPARAMETERSTARTDATE"/></xsl:attribute>
					<xsl:attribute name="DESCRIPTION"><xsl:value-of select="DESCRIPTION"/></xsl:attribute>
					<xsl:attribute name="AMOUNT"><xsl:value-of select="AMOUNT"/></xsl:attribute>
					<xsl:attribute name="STRING"><xsl:value-of select="STRING"/></xsl:attribute>
					<xsl:attribute name="BOOLEAN"><xsl:value-of select="BOOLEAN"/></xsl:attribute>
					<xsl:attribute name="PERCENTAGE"><xsl:value-of select="PERCENTAGE"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
		<!-- RESPONSE -->
	</xsl:template>
</xsl:stylesheet>
