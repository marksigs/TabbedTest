<?xml version="1.0" encoding="UTF-8"?>
<!-- Created by I. Kemp & D. Crossley May 2006-->
<!-- Updated D. Crossley 6 June 2006 -->
<!-- P. Edney - 30/06/2006 - EP926 - term in months and years not coming accross -->
<!-- P. Edney - 23/08/2006 - EP1090 - Changed ComponentPayment to ComponentPaymentOptions -->
<!-- P. Edney - 23/08/2006 - EP1090 - Removed Advance Details section -->
<!-- L.Hudson  24/08/2006 - EP1090 - Changed BrokerrWorkPhoneNumber to BrokerWorkPhoneNumber so that broker number is retrieved-->
<!-- L.Hudson  24/08/2006 - EP1090 - Add Omiga Product Code to Loan Details - COMPONENT UPDATE Miscellaneous Fields-->
<!-- AW	       12/10/2006 - EP1224 - Amended InitialAdvanceDate -->
<!-- DRC          05/03/2007 - EP2_956 Changes for Epsom 2 Mainly Multiple Loan Components-->
<!-- AW  28/03/2007 - EP2_2125 DBM178	Changes for CashBack-->
<!-- AW  29/03/2007 - EP2_956 Added OriginatorBranch, Introducer, IntroducerFee, address formatting, removed foreign insurance section-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="html" version="4" encoding="UTF-8" indent="yes"/>
	<!-- Define Styles-->
	<xsl:template match="/">
		<html>
			<head>
				<title/>
				<style>
				body {margin-left:20px;font-size: 10pt; font-family: 'trebuchet ms'}
				.head1{font-size:12pt;font-weight:bold}
				.prompt {width:10%;font-weight:bold}
				.box {border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid;padding-bottom:4px}
				.boxHead {background-color: gainsboro;border-bottom: thin solid;font-weight:bold;padding-left:4px}
				.boxPrompt {width:20%;font-weight:bold;padding-left:4px}
				.boxPromptM {font-weight:bold;padding-left:4px;padding-right:8px}
			</style>
			</head>
			<body>
				<xsl:apply-templates/>
			</body>
		</html>
	</xsl:template>
	<!-- Main Routine-->
	<xsl:template match="RESPONSE/CompletionsHandover">
			<xsl:for-each select="CustomerDetails">
			<xsl:call-template name="StdHeader"/>
			<div class="head1" align="center">Borrower details</div>
			<xsl:call-template name="borrowerDets"/>
			<div style="page-break-before:always"/>
		</xsl:for-each>
		
		<xsl:call-template name="StdHeader"/>
		<div class="head1" align="center">Bank account details</div>
		<xsl:for-each select="ApplicantBankAccountDetails">
			<xsl:call-template name="bankDets"/>
		</xsl:for-each>
		<br></br>
		<div class="head1" align="center">Solicitor details</div>
		<xsl:for-each select="SolicitorDetails">
			<xsl:call-template name="solicitorDets"/>
		</xsl:for-each>
		<div style="page-break-before:always"/>
		<xsl:call-template name="StdHeader"/>
		<div class="head1" align="center">Broker details</div>
		<xsl:for-each select="BrokerDetails">
			<xsl:call-template name="brokerDets"/>
		</xsl:for-each>
		<br></br>
		<div class="head1" align="center">Valuer details</div>
		
		<xsl:for-each select="ValuerDetails">
			<xsl:call-template name="valuerDets"/>
		</xsl:for-each>
		<div style="page-break-before:always"/>
		<xsl:call-template name="StdHeader"/>
		<div class="head1" align="center">Loan details</div>
		<xsl:for-each select="PropertyDetails">
			<xsl:call-template name="propertyDets"/>
		</xsl:for-each>
		<xsl:for-each select="ChargeDetails">
			<xsl:call-template name="chargeDets"/>
		</xsl:for-each>
		<div style="page-break-before:always"/>
		<xsl:call-template name="StdHeader"/>
		<div class="head1" align="center">Loan details cont.</div>
		<xsl:for-each select="LoanDetails">
			<xsl:call-template name="componentDets"/>
		</xsl:for-each>
		<br></br>
		<xsl:for-each select="PeripheralSecurity">
			<xsl:call-template name="PeripheralSecDets"/>
		</xsl:for-each>
		<br></br>
		<xsl:for-each select="Cashback">
		<div class="head1" align="center">Cashback details</div>
			<xsl:call-template name="CashbackDets"/>
		<br></br>
		</xsl:for-each>
		<xsl:for-each select="RetentionDetails">
		<div class="head1" align="center">Retention details</div>
			<xsl:call-template name="retentionDets"/>
		</xsl:for-each>
		<br></br>
		<div class="head1" align="center">Fee details</div>
		<xsl:for-each select="MortgageFeeDetails">
			<xsl:call-template name="feeDets"/>
		</xsl:for-each>
		<br></br>
		<div class="head1" align="center">Advance details</div>
		<xsl:for-each select="MortgageFunding">
			<xsl:call-template name="advanceDets"/>
		</xsl:for-each>
		<br></br>
	</xsl:template>
	<!-- Standard Page Header -->
	<xsl:template name="StdHeader">
		<div class="head1" align="center">db Handover Document</div>
		<div>
			<span class="boxPrompTM">Loan No.:</span>
			<span>&#160;
			    <xsl:value-of select="//RESPONSE/CompletionsHandover/LoanDetails/ComponentOrigination/ExternalAccountNumber"/>
			</span>
		</div>
		<div>
			<span class="boxPromptM">Submitted:</span>
			<span>
				<xsl:value-of select="//RESPONSE/CompletionsHandover/LoanDetails/ComponentOrigination/ApplicationDate"/>
			</span>
		</div>
		<div>
			<span class="boxPromptM">Introducer:</span>
			<span>
				<xsl:value-of select="//RESPONSE/CompletionsHandover/HeaderDetails/Introducer"/>
			</span>
		</div>
		<div>
			<span class="boxPromptM">Introducer Fee:</span>
			<span>
				<xsl:value-of select="//RESPONSE/CompletionsHandover/HeaderDetails/IntroducerFee"/>
			</span>
		</div>
	</xsl:template>
	<!-- Advance Details-->
	<xsl:template name="advanceDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">MORTGAGE FUNDING Gross Advance, Source of Funds</div>
			<div style="margin-top:4px">
			    <span class="boxPromptM">Gross advance amount</span>
				<span><xsl:value-of select="GrossAdvanceAmount"/></span>
			</div>	
		</div>
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">MORTGAGE FUNDING</div>
			<div style="margin-top:4px">
			    <span class="boxPromptM">Net advance amount</span>
				<span><xsl:value-of select="NetAdvanceAmount"/></span>
			</div>	
		</div>
	</xsl:template>
	<!-- Fees -->
	<xsl:template name="feeDets">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">MORTGAGE FEES RECEIVEABLE Directory</div>
			<xsl:for-each select="Fee">
			<div style="margin-top:4px">
			    <span style="width:33%">
					<span class="boxPromptM">Fee type</span>
					<span><xsl:value-of select="FeeType"/></span>
				</span>
				<span style="width:33%">
					<span class="boxPromptM">Fee amount</span>
					<span><xsl:value-of select="FeeAmount"/></span>
				</span>
				<span class="boxPromptM">Status</span>
				<span><xsl:value-of select="FeeStatus"/></span>
			</div>	
			</xsl:for-each>
		</div>
	</xsl:template>
	<!-- Retention -->
	<xsl:template name="retentionDets">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">MORTGAGOR RETENTIONS Directory</div>
			<div style="margin-top:4px">
			    <span class="boxPromptM">Retention</span>
				<span><xsl:value-of select="RetentionAmount"/></span>
			</div>	
			
		</div>
	</xsl:template>
	<xsl:template match="APPLICATIONFACTFIND">
		<div>
			<span class="prompt">purchase price / estimated value:</span>
			<span>&#163;<xsl:value-of select="@PURCHASEPRICEORESTIMATEDVALUE"/>
			</span>
		</div>
	</xsl:template>
	<!--Foreign  Insurance-->
	<xsl:template name="PeripheralSecDets">
		<!-- Credit Rating-->
		<div style="page-break-before:always"/>
		<xsl:call-template name="StdHeader"/>
		<div class="head1" align="center">Credit rating details</div>
		<xsl:for-each select="CreditRating">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">PERIPHERAL SECURITY UPDATE Details</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Income method used</span>
					<span><xsl:value-of select="IncomeMethodUsed"/></span>
				</span>
				<span class="boxPromptM">Bankruptcy / IVA? </span>
				<span><xsl:value-of select="BankruptcyOrIVA"/></span>
			</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Prev Arrears?</span>
					<span><xsl:value-of select="PreviousArrears"/></span>
				</span>
				<span class="boxPromptM">Income for Affordability</span>
				<span><xsl:value-of select="IncomeForAffordability"/></span>
			</div>
			<div style="margin-top:4px">
			    <span class="boxPromptM">Debt to income ratio</span>
				<span><xsl:value-of select="DebtToIncomeRatio"/></span>
			</div>
			
		</div>
			
		</xsl:for-each>
	</xsl:template>
	
	
	<!-- Loan Details-->
	<xsl:template name="componentDets">
		
		<xsl:for-each select="ComponentOrigination">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">COMPONENT UPDATE Origination</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Date of loan application</span>
					<span><xsl:value-of select="ApplicationDate"/></span>
				</span>
				<span class="boxPromptM">Date commitment issued </span>
				<span><xsl:value-of select="CommitmentIssuedDate"/></span>
			</div>	
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Commitment expiry date</span>
					<span><xsl:value-of select="CommitmentExpiryDate"/></span>
				</span>
				<span class="boxPromptM">Commitment accepted date</span>
				<span><xsl:value-of select="CommitmentAcceptedDate"/></span>
			</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Committed amount</span>
					<span><xsl:value-of select="CommittedAmount"/></span>
				</span>
				<span class="boxPromptM">Initial advance date</span>
				<span><xsl:value-of select="InitialAdvanceDate"/></span>
			</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Loan Component Amount</span>
					<span><xsl:value-of select="LoanComponentAmount"/></span>
				</span>
				<span class="boxPromptM">Original LTV</span>
				<span><xsl:value-of select="OriginalLTV"/></span>
			</div>
			<div style="margin-top:4px">
				<span style="width:50%">
			    		<span class="boxPromptM">Name of packager</span>
					<span><xsl:value-of select="PackagerName"/></span>
				</span>
				<span class="boxPromptM">Origination branch</span>
				<span><xsl:value-of select="OriginatorBranch"/></span>
			</div>
			<div style="margin-top:4px">
			   <span style="width:50%">
			    <span class="boxPromptM">Repayment Vehicle</span>
				<span><xsl:value-of select="RepaymentVehicle"/></span>
				</span>
			    	<span class="boxPromptM">External Account Number</span>
				<span><xsl:value-of select="ExternalAccountNumber"/></span>
			</div>
			  
			
			</div>
			</xsl:for-each>
		<!-- Loan Components-->	
		<xsl:for-each select="ComponentRateAndTerm">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">COMPONENT UPDATE Rate and Term</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Interest Rate Code</span>
					<span><xsl:value-of select="InterestRateCode"/></span>
				</span>
				<span class="boxPromptM">Rate or Rate Increment</span>
				<span><xsl:value-of select="RateOrRateIncrement"/></span>
			</div>	
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Date to set interest rate  </span>
					<span><xsl:value-of select="DateToSetInterestRate"/></span>
				</span>
				<span class="boxPromptM">Rate change frequency  </span>
				<span><xsl:value-of select="RateChangeFrequency"/></span>
			</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Term of loan </span>					
					<span><xsl:value-of select="TermOfLoan"/></span>
				</span>
				<span class="boxPromptM">Product code</span>
				<span><xsl:value-of select="ProductCode"/></span>
			</div>
			</div>
		</xsl:for-each>
		<xsl:for-each select="ComponentPaymentOptions">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">COMPONENT UPDATE Payment Options</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Repayment method </span>
					<span><xsl:value-of select="RepaymentMethod"/></span>
				</span>
				<span class="boxPromptM">Primary payment amount  </span>
				<span><xsl:value-of select="PrimaryPaymentAmount"/></span>
			</div>	
			
			</div>
		</xsl:for-each>
		<!-- Loan Components Misc-->
		<xsl:for-each select="ComponentMiscellaneous">
			<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">COMPONENT UPDATE Miscellaneous Fields</div>
			<div style="margin-top:4px">
			    <span class="boxPromptM">Product Scheme   </span>
				<span><xsl:value-of select="ProductScheme"/></span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPromptM">LTV Banding </span>
				<span><xsl:value-of select="LTVBanding"/></span>
			</div>
			<div style="margin-top:4px">				
				<span class="boxPromptM">Regulation Status </span>
				<span><xsl:value-of select="RegulationStatus"/></span>
			</div>	
			<div style="margin-top:4px">				
				<span class="boxPromptM">Omiga Product Number </span>
				<span><xsl:value-of select="OmigaProductCode"/></span>
			</div>
			</div>
		</xsl:for-each>
		<!-- Loan Components Investor -->
		<xsl:for-each select="InvestorDetails">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">INVESTOR UPDATE Investor Detail</div>
			<div style="margin-top:4px">
			    <span style="width:50%">
					<span class="boxPromptM">Investor</span>
					<span><xsl:value-of select="Investor"/></span>
				</span>
				<span class="boxPromptM">Committed amount</span>
				<span><xsl:value-of select="CommittedAmount"/></span>
			</div>	
			</div>
		</xsl:for-each>
	</xsl:template>
	<!-- Charge Info-->
	<xsl:template name="chargeDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CHARGE UPDATE Product Information</div>
			<xsl:for-each select="ProductInformation">
			<div style="margin-top:4px">
			   	<span class="boxPromptM">Line of Business</span>
				<span><xsl:value-of select="LineOfBusiness"/></span>
			</div>	
			<div style="margin-top:4px">
			   <span style="width:50%">
					<span class="boxPromptM">Mortgage type</span>
					<span><xsl:value-of select="MortgageType"/></span>
				</span>
				<span class="boxPromptM">Mortgage purpose type</span>
				<span><xsl:value-of select="MortgagePurposeType"/></span>
			</div>
			</xsl:for-each>
			</div>
		<div class="box" style="width:95%;margin-top:16px;">	
			<div class="boxHead">CHARGE UPDATE Payment Terms</div>
			<xsl:for-each select="PaymentTerms">
				<div style="margin-top:4px">
					<span class="boxPromptM">First Payment Date</span>
					<span><xsl:value-of select="FirstPaymentDate"/></span>
				</div>	
			</xsl:for-each>
		</div>
		<div class="box" style="width:95%;margin-top:16px;">	
			<div class="boxHead">CHARGE UPDATE Miscellaneous Charge Fields</div>
			<xsl:for-each select="MiscellaneousChargeData">
			  <xsl:for-each select="CCJDATA">
				<div style="margin-top:4px">
					<span class="boxPromptM">Last CCJ before Date of Completion-yrs (Securitisation definition) </span>
					<span><xsl:value-of select="LastCCJBeforeDataOfCompletionYear"/></span>
				</div>	
				<div style="margin-top:4px">
					<span class="boxPromptM">Last satisfied CCJ before Date of Completion -yrs (Securitisation definition)  </span>
					<span><xsl:value-of select="LastSatisfiedCCJBeforeDateOfCompletionYears"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">Advised Sale? </span>
					<span><xsl:value-of select="AdvisedSale"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">Total number of CCJs    (FSA definition) </span>
					<span><xsl:value-of select="TotalNumberOfCCJsLessThan3Years"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">Total value of CCJs    (FSA definition) </span>
					<span><xsl:value-of select="TotalValueOfCCJsLessThan3Years"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">Value of all unsatisfied CCJs (Securitisation definition)  </span>
					<span><xsl:value-of select="ValueOfAllUnsatisfiedCCJs"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">No of all unsatisfied CCJs (Securitisation definition)  </span>
					<span><xsl:value-of select="NoOfAllUnsatisfiedCCJs"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">Value of all satisfied CCJs (Securitisation definition)  </span>
					<span><xsl:value-of select="ValueOfAllSatisfiedCCJs"/></span>
				</div>
				<div style="margin-top:4px">
					<span class="boxPromptM">No of all satisfied CCJs (Securitisation definition)  </span>
					<span><xsl:value-of select="NoOfAllSatisfiedCCJs"/></span>
				</div>
				
				</xsl:for-each>
				
			</xsl:for-each>
		</div>
	</xsl:template>
	<!-- Security Property-->
	<xsl:template name="propertyDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">SECURITY UPDATE Address Entry</div>
			<div style="margin-top:4px">
				<xsl:for-each select="PropertyAddress">
				<xsl:choose>
						<xsl:when test="string-length(FlatNameOrNumber) != 0  or string-length(HouseOrBuildingName) != 0">
							<xsl:call-template name="addrOne"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name="addrTwo"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">SECURITY UPDATE Property Description</div>
			<div style="margin-top:4px">
				<xsl:for-each select="PropertyDetailsDescription">
					<div style="margin-top:4px">
						<span class="boxPromptM">Title Number</span>
						<span>
							<xsl:if test="string-length(TitleNoOne) != 0"><xsl:value-of select="TitleNoOne"/>&#160;</xsl:if>
							<xsl:if test="string-length(TitleNoTwo) != 0"><xsl:value-of select="TitleNoTwo"/>&#160;</xsl:if>
							<xsl:if test="string-length(TitleNoThree) != 0"><xsl:value-of select="TitleNoThree"/></xsl:if>
						</span>
					</div>
					<div style="margin-top:4px">
						<span class="boxPromptM">Property tenure</span>
						<xsl:value-of select="PropertyTenure"/>
					</div>
					<div style="margin-top:4px">
						<span style="width:50%">
							 <span class="boxPromptM">Property type</span>
							 <xsl:value-of select="PropertyType"/>
						 </span>
						 <span class="boxPromptM">Property Sub type</span>
						 <xsl:value-of select="PropertySubType"/>
					 
					</div>
					<div style="margin-top:4px">
						<span style="width:50%">
								 <span class="boxPromptM">Property Description</span>
								<xsl:value-of select="PropertyDescription"/>
						 </span>
						 <span class="boxPromptM">Freehold title?</span>
							 <xsl:value-of select="FreeholdTitle"/>
					</div>
					<div style="margin-top:4px">
						<span class="boxPromptM">Property Let?</span>
						<xsl:value-of select="PropertyLet"/>
					</div>
				</xsl:for-each>
				</div>
		</div>
		<!-- Valuation -->
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">SECURITY UPDATE Valuation</div>
			<div style="margin-top:4px">
				<xsl:for-each select="PropertyValuationDetails">
					<div style="margin-top:4px">
					   <span style="width:50%">
							<span class="boxPromptM">Date of current valuation</span>
							<xsl:value-of select="LastValuationDate"/>
					  </span>
					  		<span class="boxPromptM">Property value</span>
						 <xsl:value-of select="PropertyValue"/>
					</div>
					<div style="margin-top:4px">
					    <span style="width:50%">
							<span class="boxPromptM">Total Valuation</span>
							<xsl:value-of select="TotalValuation"/>
						</span>	
						<span class="boxPromptM">Valuation Type</span>
						 <xsl:value-of select="ValuationType"/>
					</div>
					<div style="margin-top:4px">
						<span class="boxPromptM">Purchase price </span>
							 <xsl:value-of select="PropertyPurchasePrice"/>
					</div>
					<div style="margin-top:4px">
						 <span style="width:50%">
								 <span class="boxPromptM">Servicing branch</span>
								<xsl:value-of select="ServicingBranch"/>
						 </span>
						 <span class="boxPromptM">Underwriter</span>
							 <xsl:value-of select="Underwriter"/>
					</div>
					</xsl:for-each>
				</div>
		</div>
		<!-- Security Misc-->
		<div class="box" style="width:95%;margin-top:16px;">
				<div class="boxHead">SECURITY UPDATE Miscellaneous Security Fields</div>
				<div style="margin-top:4px">
					<xsl:for-each select="MiscellaneousSecurityData">
						<div style="margin-top:4px">
							<span style="width:50%">
								<span class="boxPromptM">Property Purchase date </span>
								<xsl:value-of select="PropertyPurchaseDate"/>
							</span>
							<span class="boxPromptM">Original LTV </span>
							<xsl:value-of select="LoanToValue"/>
						</div>
					</xsl:for-each>
				</div>
			</div>
	</xsl:template>
	<!-- Valuer-->
	<xsl:template name="valuerDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CLIENT INFORMATION UPDATE Name And Address Entry</div>
			<div style="margin-top:4px">
			   	<span class="boxPrompt">Valuer number </span>
				<span><xsl:value-of select="ValuerNumber"/></span>
			</div>	
			<div style="margin-top:4px">
				<span class="boxPrompt">Company name </span>
				<span><xsl:value-of select="CompanyName"/></span>
			</div>
		</div>
	</xsl:template>
	<!-- Broker -->
	<xsl:template name="brokerDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CLIENT INFORMATION UPDATE Name And Address Entry</div>
			<div style="margin-top:4px">
			   	<span class="boxPrompt">Broker number </span>
				<span><xsl:value-of select="BrokerNumber"/></span>
			</div>	
			<div style="margin-top:4px">
				<span class="boxPrompt">Company name </span>
				<span><xsl:value-of select="CompanyName"/></span>
			</div>
			<div style="margin-top:4px">
				<xsl:for-each select="BrokerAddress">
				 <xsl:choose>
						<xsl:when test="string-length(FlatNameOrNumber) != 0  or string-length(HouseOrBuildingName) != 0">
							<xsl:call-template name="addrOne"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name="addrTwo"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Telephones</div>
			
			<div style="margin-top:4px">
			
				<span class="boxPrompt">Work Phone Number</span>
				<span>
					<xsl:if test="string-length(BrokerWorkPhoneNumber/AreaCode) != 0"><xsl:value-of select="BrokerWorkPhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(BrokerWorkPhoneNumber/LocalNumber) != 0"><xsl:value-of select="BrokerWorkPhoneNumber/LocalNumber"/>&#160;</xsl:if>
					<xsl:if test="string-length(BrokerWorkPhoneNumber/Extension) != 0">extn. <xsl:value-of select="BrokerWorkPhoneNumber/Extension"/></xsl:if>
				</span>
			</div>
			<div style="margin-top:4px">
			
				<span class="boxPrompt">Other phone number</span>
				<span>
					<xsl:if test="string-length(BrokerOtherPhoneNumber/AreaCode) != 0"><xsl:value-of select="BrokerOtherPhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(BrokerOtherPhoneNumber/LocalNumber) != 0"><xsl:value-of select="BrokerOtherPhoneNumber/LocalNumber"/>&#160;</xsl:if>
					<xsl:if test="string-length(BrokerOtherPhoneNumber/Extension) != 0">extn. <xsl:value-of select="BrokerOtherPhoneNumber/Extension"/></xsl:if>
				</span>
			</div>
		</div>
	</xsl:template>
	<!-- Solicitor-->
	<xsl:template name="solicitorDets">
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CLIENT INFORMATION UPDATE Name And Address Entry</div>
			<div style="margin-top:4px">
			   	<span class="boxPrompt">Solicitor number</span>
				<span><xsl:value-of select="SolicitorNumber"/></span>
			</div>	
			<div style="margin-top:4px">
				<span class="boxPrompt">Company name </span>
				<span><xsl:value-of select="CompanyName"/></span>
			</div>
			<div style="margin-top:4px">
				<xsl:for-each select="SolicitorAddress">
				<xsl:choose>
						<xsl:when test="string-length(FlatNameOrNumber) != 0  or string-length(HouseOrBuildingName) != 0">
							<xsl:call-template name="addrOne"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:call-template name="addrTwo"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Telephones</div>
			
			<div style="margin-top:4px">
			
				<span class="boxPrompt">Work Phone Number</span>
				<span>
					<xsl:if test="string-length(SolicitorWorkPhoneNumber/AreaCode) != 0"><xsl:value-of select="SolicitorWorkPhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(SolicitorWorkPhoneNumber/LocalNumber) != 0"><xsl:value-of select="SolicitorWorkPhoneNumber/LocalNumber"/>&#160;</xsl:if>
					<xsl:if test="string-length(SolicitorWorkPhoneNumber/Extension) != 0">extn. <xsl:value-of select="SolicitorWorkPhoneNumber/Extension"/></xsl:if>
				</span>
			</div>
			<div style="margin-top:4px">
			
				<span class="boxPrompt">Other phone number</span>
				<span>
					<xsl:if test="string-length(SolicitorOtherPhoneNumber/AreaCode) != 0"><xsl:value-of select="SolicitorOtherPhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(SolicitorOtherPhoneNumber/LocalNumber) != 0"><xsl:value-of select="SolicitorOtherPhoneNumber/LocalNumber"/>&#160;</xsl:if>
					<xsl:if test="string-length(SolicitorOtherPhoneNumber/Extension) != 0">extn. <xsl:value-of select="SolicitorOtherPhoneNumber/Extension"/></xsl:if>
				</span>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px">
			<div class="boxHead">CUSTOMER DX ADDRESS UPDATE</div>
			
			<div style="margin-top:4px">
			
				<span class="boxPrompt">Exchange service/number</span>
				<span>
					<xsl:value-of select="DXNumber"/>
					
				</span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Exchange city</span>
				<span>
					<xsl:value-of select="DXLocation"/>
					
				</span>
			</div>
		</div>
	</xsl:template>
	<!--Customer (aka Borower) -->
	<xsl:template name="borrowerDets">
		<div class="box" style="width:95%">
			<div class="boxHead">CLIENT INFORMATION UPDATE Name and Address Entry</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Customer number:</span>
				<span>
					<xsl:value-of select="CustomerNumber"/>
				</span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Name:</span>
				<xsl:for-each select="CustomerName">
					<span>
						<xsl:if test="string-length(Title) != 0"><xsl:value-of select="Title"/>&#160;</xsl:if>
						<xsl:if test="string-length(ForeName) != 0"><xsl:value-of select="ForeName"/>&#160;</xsl:if>
						<xsl:if test="string-length(SecondForeName) != 0"><xsl:value-of select="SecondForeName"/>&#160;</xsl:if>
						<xsl:if test="string-length(OtherForeNames) != 0"><xsl:value-of select="OtherForeNames"/>&#160;</xsl:if>
						<xsl:if test="string-length(SurName) != 0"><xsl:value-of select="SurName"/>&#160;</xsl:if>
					</span>
				</xsl:for-each>
			</div>
			<xsl:for-each select="CurrentAddressDetails/AddressDetails">
				<xsl:choose>
					<xsl:when test="string-length(FlatNameOrNumber) != 0  or string-length(HouseOrBuildingName) != 0">
						<xsl:call-template name="addrOne"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="addrTwo"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</div>
		<div class="box" style="width:95%;margin-top:16px">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Identification</div>
			<div style="margin-top:4px">
				<span style="width:50%">
					<span class="boxPromptM">Birth Date</span>
					<span>
						<xsl:value-of select="DateOfBirth"/>
					</span>
				</span>
				<span>
					<span class="boxPromptM">Gender</span>
					<span>
						<xsl:value-of select="Gender"/>
					</span>
				</span>
			</div>
			<div style="margin-top:4px">
				<span style="width:50%">
					<span class="boxPromptM">Marital Status</span>
					<span>
						<xsl:value-of select="MaritalStatus"/>
					</span>
				</span>
				<span>
					<span class="boxPromptM">National Insurance Number</span>
					<span>
						<xsl:value-of select="NationalInsuranceNumber"/>
					</span>
				</span>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Telephones</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Home Phone Number</span>
				<span>
					<xsl:if test="string-length(HomePhoneNumber/AreaCode) != 0"><xsl:value-of select="HomePhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(HomePhoneNumber/LocalNumber) != 0"><xsl:value-of select="HomePhoneNumber/LocalNumber"/></xsl:if>
				</span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Work Phone Number</span>
				<span>
					<xsl:if test="string-length(WorkPhoneNumber/AreaCode) != 0"><xsl:value-of select="WorkPhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(WorkPhoneNumber/LocalNumber) != 0"><xsl:value-of select="WorkPhoneNumber/LocalNumber"/>&#160;</xsl:if>
					<xsl:if test="string-length(WorkPhoneNumber/Extension) != 0">extn. <xsl:value-of select="WorkPhoneNumber/Extension"/></xsl:if>
				</span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Mobile Phone Number</span>
				<span>
					<xsl:if test="string-length(MobilePhoneNumber/AreaCode) != 0"><xsl:value-of select="MobilePhoneNumber/AreaCode"/>&#160;</xsl:if>
					<xsl:if test="string-length(MobilePhoneNumber/LocalNumber) != 0"><xsl:value-of select="MobilePhoneNumber/LocalNumber"/></xsl:if>
				</span>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Demographic</div>
			<div style="margin-top:4px">
				<span class="boxPromptM">Lending Branch</span>
				<span><xsl:value-of select="LendingBranch"/></span>
			</div>
			<div style="margin-top:4px">
				<span style="width:50%">
					<span class="boxPromptM">Occupation / Business Type</span>
					<span><xsl:value-of select="OccupationBusinessType"/></span>
				</span>
				<span class="boxPromptM">Industry Sector</span>
				<span><xsl:value-of select="IndustrySector"/></span>
			</div>
		</div>
		<div class="box" style="width:95%;margin-top:16px;">
			<div class="boxHead">CLIENT INFORMATION UPDATE Client Detail - Miscellaneous</div>
			<div style="margin-top:4px">
			   <span style="width:50%">
					<span class="boxPromptM">Proof of income </span>
					<span><xsl:value-of select="ProofOfIncome"/></span>
				 </span>
				<span class="boxPromptM">Gross income</span>
				<span><xsl:value-of select="GrossIncome"/></span>
			</div>
		</div>
	</xsl:template>
	<!-- Bank Account Details-->
	<xsl:template name="bankDets">
	    
		<div class="box" style="width:95%;margin-top:20px">
			<div class="boxHead">BANKING INFORMATION UPDATE</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Payer name</span>
				<span>
				    <xsl:value-of select="AccountName"/>
				</span>
			</div>	
			<div style="margin-top:4px">
				<span class="boxPrompt">Bank sort code</span>
				<span><xsl:value-of select="SortCode"/></span>
			</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Bank account number</span>
				<span><xsl:value-of select="AccountNumber"/></span>
			</div>
		</div>
	</xsl:template>
	<!-- Address details-->
	<xsl:template name="addrOne">
		<div style="margin-top:4px">
			
		    <xsl:if test="string-length(FlatNameOrNumber) != 0  or string-length(HouseOrBuildingName) != 0">
			<span class="boxPrompt">Address:</span>
							<span>
								<xsl:if test="string-length(FlatNameOrNumber) != 0"><xsl:value-of select="FlatNameOrNumber"/>&#160;</xsl:if>
								<xsl:if test="string-length(HouseOrBuildingName) != 0"><xsl:value-of select="HouseOrBuildingName"/></xsl:if>
							</span>
			</xsl:if>
			 
			<xsl:call-template name="addrFour"/>
		</div>
	</xsl:template>
	<xsl:template name="addrTwo">
		<div>
			<span class="boxPrompt">Address:</span>
			<span>
				<xsl:if test="string-length(HouseOrBuildingNumber) != 0"><xsl:value-of select="HouseOrBuildingNumber"/>&#160;</xsl:if>
				<xsl:if test="string-length(Street) != 0"><xsl:value-of select="Street"/></xsl:if>
			</span>
			<xsl:call-template name="addrThree"/>
		</div>
	</xsl:template>
	<xsl:template name="addrFour">
		<div>
			<span class="boxPrompt"></span>
			<span>
				<xsl:if test="string-length(HouseOrBuildingNumber) != 0"><xsl:value-of select="HouseOrBuildingNumber"/>&#160;</xsl:if>
				<xsl:if test="string-length(Street) != 0"><xsl:value-of select="Street"/></xsl:if>
			</span>
			<xsl:call-template name="addrThree"/>
		</div>
	</xsl:template>
	<xsl:template name="addrThree">
		<xsl:if test="string-length(District) != 0">
			<div>
				<span class="boxPrompt">&#160;</span>
				<span>
					<xsl:value-of select="District"/>
				</span>
			</div>
		</xsl:if>
		<xsl:if test="string-length(TownOrCity) != 0">
			<div>
				<span class="boxPrompt">&#160;</span>
				<span>
					<xsl:value-of select="TownOrCity"/>
				</span>
			</div>
		</xsl:if>
		<xsl:if test="string-length(County) != 0">
			<div>
				<span class="boxPrompt">&#160;</span>
				<span>
					<xsl:value-of select="County"/>
				</span>
			</div>
		</xsl:if>
		<xsl:if test="string-length(PostCode) != 0">
			<div>
				<span class="boxPrompt">&#160;</span>
				<span>
					<xsl:value-of select="PostCode"/>
				</span>
			</div>
		</xsl:if>
		<xsl:if test="string-length(CountryCode) != 0">
			<div>
				<span class="boxPrompt">&#160;</span>
				<span>
					<xsl:value-of select="CountryCode"/>
				</span>
			</div>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CashbackDets">
		<div class="box" style="width:95%;margin-top:20px">
			<div class="boxHead">CASHBACK Details</div>
			<div style="margin-top:4px">
				<span class="boxPrompt">Cashback</span>
				<span>
				    <xsl:value-of select="CashbackAmount"/>
				</span>
			</div>	
		</div>
	</xsl:template>
</xsl:stylesheet>
