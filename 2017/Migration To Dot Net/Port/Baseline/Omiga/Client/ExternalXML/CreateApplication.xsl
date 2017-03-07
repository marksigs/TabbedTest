<?xml version="1.0" encoding="UTF-8"?>
<!--  
CreateApplication.xsl is used by the omDataIngestion component of Omiga4 BMIDS.

Modification history
_________________________________________________________________________________________________ 
 Date		Author		What has changed?													 
_________________________________________________________________________________________________ 
24/07/2002	MO			Created		
02/10/2002  MO			Fixed various bugs, noted from testing in AQR BMIDS00542
16/10/2002  MO			Performed second 'phase' upgrades as per AQR BMIDS00462
29/10/2002  MO			Fixed various bugs, noted from testing in AQR BMIDS00750
15/11/2002	MO			Fixed various bugs, noted from UAT testing in AQR BMIDS00959, BM Ref 0473
19/11/2002	MO			Fixed bugs, noted from System testing in AQR BMIDS01004
25/11/2002	MO			Made changes to ingest activity details noted in AQR BMIDS01073
25/11/2002	MO			Made changes to correct the size of areacode in AQR BMIDS01067
27/11/20002 MO			Fixed bugs noted frmo BMIDS UAT testing in AQR BMIDS01097
_________________________________________________________________________________________________ 
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	
	<!-- Setup common variables -->
	<xsl:variable name="UserID" select="/REQUEST/@USERID"/>
	<xsl:variable name="UnitID" select="/REQUEST/@UNITID"/>
	<xsl:variable name="ChannelID" select="/REQUEST/@CHANNELID"/>
	<xsl:variable name="UserAuthorityLevel" select="/REQUEST/@USERAUTHORITYLEVEL"/>
	<xsl:variable name="Now" select="/REQUEST/@NOW"/> <!-- This variable is put in by the ingestion engine -->
	
	<!-- 
	Template - Main "/"
	Purpose - Starting point of the stylesheet
	Author - MO
	Date - 24/07/2002
	-->
	<xsl:template match="/">
		<xsl:element name="REQUEST">
			<xsl:attribute name="OPERATION"><xsl:value-of select="REQUEST/@OPERATION"/></xsl:attribute>
			<xsl:attribute name="USERID"><xsl:value-of select="$UserID"/></xsl:attribute>
			<xsl:attribute name="UNITID"><xsl:value-of select="$UnitID"/></xsl:attribute>
			<xsl:attribute name="CHANNELID"><xsl:value-of select="$ChannelID"/></xsl:attribute>
			<xsl:attribute name="USERAUTHORITYLEVEL"><xsl:value-of select="$UserAuthorityLevel"/></xsl:attribute>
			<!-- Find the Proposal Node -->
			<xsl:for-each select='REQUEST/ProposalRoot/Proposal'>
				<!-- APPLICATON Node -->
				<xsl:element name="APPLICATION">
					<xsl:attribute name="APPLICATIONNUMBER"></xsl:attribute>
					<xsl:attribute name="CORRESPONDENCESALUTATION"><xsl:call-template name="BuildApplicationCorrespondenceSalutation"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONDATE"><xsl:value-of select="$Now"/></xsl:attribute>
					<xsl:attribute name="TYPEOFBUYER">10</xsl:attribute> <!-- Not Applicable -->
					<xsl:attribute name="CHANNELID"><xsl:value-of select="$ChannelID"/></xsl:attribute>
					<xsl:attribute name="OTHERSYSTEMACCOUNTNUMBER"></xsl:attribute>
					<xsl:attribute name="INGESTIONAPPLICATIONNUMBER"><xsl:value-of select="Id"/></xsl:attribute>
					
					<!-- USERHISTORY Node -->
					<xsl:element name="USERHISTORY">
						<xsl:attribute name="UNITID"><xsl:value-of select="$UnitID"/></xsl:attribute>
						<xsl:attribute name="USERID"><xsl:value-of select="$UserID"/></xsl:attribute>
						<xsl:attribute name="USERHISTORYDATE"><xsl:value-of select="$Now"/></xsl:attribute>
					</xsl:element> <!-- USERHISTORY -->
					
					<!-- APPLICATIONPRIORITY -->
					<xsl:element name="APPLICATIONPRIORITY">
						<xsl:attribute name="APPLICATIONPRIORITYVALUE">20</xsl:attribute> <!-- Normal -->
						<xsl:attribute name="DATETIMEASSIGNED"><xsl:value-of select="$Now"/></xsl:attribute>
						<xsl:attribute name="USERID"><xsl:value-of select="$UserID"/></xsl:attribute>
						<xsl:attribute name="UNITID"><xsl:value-of select="$UnitID"/></xsl:attribute>
					</xsl:element> <!-- APPLICATIONPRIORITY -->

					<!-- APPLICATONFACTFIND -->
					<xsl:element name="APPLICATIONFACTFIND">
						<xsl:attribute name="ACTIVEQUOTENUMBER">1</xsl:attribute> <!-- There is only 1 quote -->
						<xsl:attribute name="ACCEPTEDQUOTENUMBER">1</xsl:attribute> <!-- There is only 1 quote -->
						<xsl:attribute name="APPLICATIONCURRENCY">10</xsl:attribute> <!-- Currency value for stirling -->
						<xsl:attribute name="ATTITUDETOBORROWINGSCORE"></xsl:attribute>
						<xsl:attribute name="CREDITSCORE"></xsl:attribute>
						<xsl:attribute name="DIRECTINDIRECTBUSINESS">6</xsl:attribute> <!-- Website -->
						<xsl:attribute name="LENDERCODE">10</xsl:attribute> <!-- BM's Lender code will have to go in here -->
						<xsl:attribute name="MARKETINGSOURCE"></xsl:attribute>
						<xsl:attribute name="NUMBEROFAPPLICANTS"><xsl:value-of select="NumberOfApplicants"/></xsl:attribute>
						<xsl:attribute name="NUMBEROFGUARANTORS">0</xsl:attribute> <!-- Guarantors are not supported in GIWS -->
						<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE"><xsl:value-of select="PropertyRoot/Property/PurchasePrice"/></xsl:attribute>
						<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="MortgageStatusTypeId"/></xsl:attribute> <!-- New Loan -->
						<xsl:attribute name="NATUREOFLOAN">1</xsl:attribute> <!-- Residential -->
						<xsl:attribute name="SPECIALSCHEME"><xsl:value-of select="MortgageTypeId"/></xsl:attribute> 
						<xsl:attribute name="APPLICATIONNOTES"></xsl:attribute>
						<!-- Introducer details -->
						<xsl:choose>
							<!-- With a mortgage club -->
							<xsl:when test="ProposalAgentRoot/ProposalAgent/MortgageClubYN = 1">
								<xsl:attribute name="INTRODUCERIDLEVEL1"><xsl:value-of select="MortgageClub/ExternalIntroducerId"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL1"><xsl:value-of select="MortgageClub/MCMCCBNo"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERIDLEVEL2"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/ExternalCompanyIntroducerId"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL2"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/AgentMCCBRegNo"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERIDLEVEL3"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/ExternalIndividualIntroducerId"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL3"></xsl:attribute>
								<!-- Is the individual the contact? -->
								<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL1">0</xsl:attribute>
								<xsl:choose>
									<xsl:when test="ProposalAgentRoot/ProposalAgent/ContactYN = 1">
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL2">0</xsl:attribute>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL3">1</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL2">1</xsl:attribute>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL3">0</xsl:attribute>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<!-- Without a mortgage club -->
							<xsl:otherwise>
								<xsl:attribute name="INTRODUCERIDLEVEL1"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/ExternalCompanyIntroducerId"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL1"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/AgentMCCBRegNo"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERIDLEVEL2"><xsl:value-of select="ProposalAgentRoot/ProposalAgent/ExternalIndividualIntroducerId"/></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL2"></xsl:attribute>
								<xsl:attribute name="INTRODUCERIDLEVEL3"></xsl:attribute>
								<xsl:attribute name="INTRODUCERMCCBLEVEL3"></xsl:attribute>
								<!-- Is the individual the contact? -->
								<xsl:choose>
									<xsl:when test="ProposalAgentRoot/ProposalAgent/ContactYN = 1">
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL1">0</xsl:attribute>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL2">1</xsl:attribute>
									</xsl:when>
									<xsl:otherwise>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL1">1</xsl:attribute>
										<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL2">0</xsl:attribute>
									</xsl:otherwise>
								</xsl:choose>
								<xsl:attribute name="INTRODUCERCORRESPONDINDLEVEL3"></xsl:attribute>
							</xsl:otherwise>
						</xsl:choose> <!-- Introducer details -->
									
						<!-- Select the applicant root -->
						<xsl:for-each select='ApplicantRoot/Applicant'>
							
							<!-- Note the applicant number -->
							<xsl:variable name="ApplicantNumber" select="position()"/>
							
							<!-- CUSTOMERNODE -->
							<xsl:element name="CUSTOMER">
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="OMIGAClientId"/></xsl:attribute>
								<xsl:attribute name="OTHERSYSTEMCUSTOMERNUMBER"><xsl:value-of select="CIFCustomerId"/></xsl:attribute>
								<xsl:attribute name="OTHERSYSTEMTYPE"></xsl:attribute>
								<xsl:attribute name="CHANNELID"><xsl:value-of select="$ChannelID"/></xsl:attribute>

								<!-- CUSTOMERVERSION -->
								<xsl:element name="CUSTOMERVERSION">
									<xsl:attribute name="CONTACTEMAILADDRESS"><xsl:value-of select="ApplicantAddressRoot/ApplicantAddress[AddressTypeId = 2]/ApplicantAddressEmail"/></xsl:attribute> <!-- Email address is stored against the applicants current address -->
									<xsl:attribute name="CORRESPONDENCESALUATION"></xsl:attribute>
									<xsl:attribute name="INTERMEDIARYGUID"></xsl:attribute>
									<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="DOB"/></xsl:attribute>
									<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="Forename1"/></xsl:attribute>
									<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="Forename2"/></xsl:attribute>
									<xsl:attribute name="GENDER"><xsl:value-of select="GenderTypeId"/></xsl:attribute>
									<xsl:attribute name="MARITALSTATUS"><xsl:value-of select="MaritalStatusTypeId"/></xsl:attribute>
									<xsl:attribute name="MEMBEROFSTAFF">0</xsl:attribute>
									<xsl:attribute name="NATIONALINSURANCENUMBER"><xsl:value-of select="NINo"/></xsl:attribute>
									<xsl:attribute name="INTERMEDIARYGUID"></xsl:attribute>
									<xsl:attribute name="SURNAME"><xsl:value-of select="Surname"/></xsl:attribute>
									<xsl:attribute name="TITLE"><xsl:value-of select="TitleTypeId"/></xsl:attribute>
									<xsl:attribute name="UKRESIDENTINDICATOR"><xsl:call-template name="DetermineApplicantUKResident"><xsl:with-param name="COUNTRY" select="ApplicantAddressRoot/ApplicantAddress[AddressTypeId = 2]/ApplicantAddressCountryTypeId" /></xsl:call-template></xsl:attribute>
									<xsl:attribute name="NATIONALITY"><xsl:value-of select="NationalityTypeId"/></xsl:attribute>
									<xsl:attribute name="NUMBEROFDEPENDANTS">0</xsl:attribute> <!-- This value is set in pre-processing for the first application, all others are 0 -->
									
									<!-- CUSTOMERROLE -->
									<xsl:element name="CUSTOMERROLE">
										<xsl:attribute name="CUSTOMERROLETYPE">1</xsl:attribute> <!-- Applicant role type -->
										<xsl:attribute name="CUSTOMERORDER"><xsl:value-of select="position()"/></xsl:attribute>
									</xsl:element> <!-- CUSTOMERROLE -->
									
									<!--  For each ApplicantAddress (Current, previous and portfolio addresses) -->
									<xsl:for-each select="ApplicantAddressRoot/ApplicantAddress">

										<!-- ADDRESS -->
										<xsl:element name="ADDRESS">
											<!-- Addresses is GIWS are stored weird, basically there is address Number, lines 1, 2, 3, Town, Country, Country on the GUI but they
											are stored as Number, Name, Street, District, Town, County, Country  - There is no where to store district but the business has approved this -->
											<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="ApplicantAddressBuildingNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="ApplicantAddressBuildingNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="FLATNUMBER"></xsl:attribute>
											<xsl:attribute name="STREET"><xsl:value-of select="ApplicantAddressBuildingName"/></xsl:attribute>
											<xsl:attribute name="DISTRICT"><xsl:value-of select="ApplicantAddressStreet"/></xsl:attribute>
											<xsl:attribute name="TOWN"><xsl:value-of select="ApplicantAddressTown"/></xsl:attribute>
											<xsl:attribute name="COUNTY"><xsl:value-of select="ApplicantAddressCounty"/></xsl:attribute>
											<xsl:attribute name="COUNTRY"><xsl:value-of select="ApplicantAddressCountryTypeId"/></xsl:attribute>
											<xsl:attribute name="POSTCODE"><xsl:value-of select="ApplicantAddressPostCode"/></xsl:attribute>

											<!-- CUSTOMERADDRESS -->
											<xsl:element name="CUSTOMERADDRESS">
												<xsl:attribute name="ADDRESSTYPE"><xsl:value-of select="AddressTypeId"/></xsl:attribute>
												<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="DateMovedIn"/></xsl:attribute>
												<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="DateMovedOut"/></xsl:attribute>
												<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="ResidentialStatusTypeId"/></xsl:attribute>
												
												<!-- If its the current property, save current property details -->
												<xsl:if test="AddressTypeId = 2">
													<xsl:element name="CURRENTPROPERTY">
														<xsl:attribute name="ESTIMATEDCURRENTVALUE"><xsl:value-of select="SellingPrice"/></xsl:attribute>
													</xsl:element>
												</xsl:if>
																							
												<!-- Is there a mortgage account associated with this address? -->
												<!-- Is it owner occupied, have a mortgage = yes and if its a current address then the sameasapplicantflag is different  -->
												<xsl:if test="ResidentialStatusTypeId = 2 and HaveMortgageYN = 1 and (AddressTypeId !=2 or (AddressTypeId = 2 and MortgageLenderSameAsApplicant = 99))"> 
	
													<!-- LENDERDETAILS -->
													<xsl:element name="LENDERLANDLORDDETAILS">
														<xsl:attribute name="CONTACTFORENAME"></xsl:attribute>
														<xsl:attribute name="CONTACTSURNAME"></xsl:attribute>
														<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
														<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
														<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="LenderAddressEmail"/></xsl:attribute>														
														
														<!-- LENDERLANDLORDTELEPHONEDETAILS -->
														<xsl:element name="LENDERLANDLORDTELEPHONEDETAILS">
															<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
															<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
														</xsl:element> <!-- LENDERLANDLORDTELEPHONEDETAILS -->
														
														<!-- LENDERLANDLORDTELEPHONEDETAILS -->
														<xsl:element name="LENDERLANDLORDTELEPHONEDETAILS">
															<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
															<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
														</xsl:element> <!-- LENDERLANDLORDTELEPHONEDETAILS -->
														
														<!-- LENDERADDRESS -->
														<xsl:element name="LENDERLANDLORDADDRESS">
															<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LenderAddressBuildingNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LenderAddressBuildingNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="FLATNUMBER"></xsl:attribute>
															<xsl:attribute name="STREET"><xsl:value-of select="LenderAddressBuildingName"/></xsl:attribute>
															<xsl:attribute name="DISTRICT"><xsl:value-of select="LenderAddressStreet"/></xsl:attribute>
															<xsl:attribute name="TOWN"><xsl:value-of select="LenderAddressTown"/></xsl:attribute>
															<xsl:attribute name="COUNTY"><xsl:value-of select="LenderAddressCounty"/></xsl:attribute>
															<xsl:attribute name="COUNTRY"><xsl:value-of select="LenderAddressCountryTypeId"/></xsl:attribute>
															<xsl:attribute name="POSTCODE"><xsl:value-of select="LenderAddressPostCode"/></xsl:attribute>

															<!-- LENDERCOMPANYDETAILS -->
															<xsl:element name="LENDERLANDLORDTHIRDPARTY">
																<xsl:attribute name="THIRDPARTYTYPE">3</xsl:attribute> <!-- Other Lender Thirdparty type -->
																<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="MortgageLenderName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
																
																<!-- ACCOUNT DETAILS -->
																<xsl:element name="ACCOUNT">
																	<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="MortgageAccountNo"/></xsl:attribute>
																	
																	<!-- MORTGAGEACCOUNTDETAILS -->
																	<xsl:element name="MORTGAGEACCOUNT">
																		<xsl:attribute name="SECONDCHARGEINDICATOR"><xsl:value-of select="MortgageSecondCharge"/></xsl:attribute>
																		<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="MonthlyIncomeRental"/></xsl:attribute>
																		<xsl:attribute name="ESTIMATEDCURRENTVALUE"><xsl:value-of select="SellingPrice"/></xsl:attribute>
																		<xsl:attribute name="ADDRESSGUID"><xsl:text>../../../../../@ADDRESSGUID</xsl:text></xsl:attribute> <!-- This key is inherited from an x path -->
																		
																		<!-- MORTGAGELOANS -->
																		<xsl:element name="MORTGAGELOAN">
																			<xsl:attribute name="LOANACCOUNTNUMBER"><xsl:value-of select="MortgageAccountNo"/></xsl:attribute>
																			<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MortgageMonthlyPayment"/></xsl:attribute> 
																			<xsl:attribute name="ORIGINALLOANAMOUNT"><xsl:value-of select="MortgageOriginalLoan"/></xsl:attribute> 
																			<xsl:attribute name="OUTSTANDINGBALANCE"><xsl:value-of select="MortgageOutstandingAmount"/></xsl:attribute> 
																			<xsl:attribute name="PURPOSEOFLOAN"><xsl:value-of select="MortgagePurposeTypeId"/></xsl:attribute> 
																			<xsl:attribute name="STARTDATE"><xsl:value-of select="MortgageStartDate"/></xsl:attribute> 
																			
																			<!-- Redemption Status -->
																			<xsl:choose>
																				<!-- If its a previous address set the redemption status and date -->
																				<xsl:when test="AddressTypeId = 3">
																					<xsl:attribute name="REDEMPTIONSTATUS">40</xsl:attribute> <!-- Already Redeemed -->
																					<xsl:attribute name="REDEMPTIONDATE"><xsl:value-of select="MortgageDateOfRedemption"/></xsl:attribute>  
																					
																				</xsl:when> <!-- Previous Address -->
																				
																				<!-- If its not a previous address, work it out! -->
																				<xsl:otherwise>
																				
																					<!-- Is this mortgage to be redeemed? -->
																					<xsl:if test="MortgageRepaidYN = 1"> <!-- Yes -->
																						<xsl:attribute name="REDEMPTIONSTATUS">10</xsl:attribute> <!-- To be Redeemed - This advance -->
																					</xsl:if> <!-- Mortgage repaid : Yes -->
																					<xsl:if test="MortgageRepaidYN = 0"> <!-- No -->
																						<!-- Is the property going to be let? -->
																						<xsl:if test="MortgageLet = 1"> <!-- Yes -->
																							<xsl:attribute name="REDEMPTIONSTATUS">60</xsl:attribute> <!-- Other BTL Lender -->
																						</xsl:if>
																						<xsl:if test="MortgageLet = 0"> <!-- No -->
																							<xsl:attribute name="REDEMPTIONSTATUS">50</xsl:attribute> <!-- Not to be Redeemed -->
																						</xsl:if>
																					</xsl:if> <!-- Mortgage repaid : No -->
																					
																				</xsl:otherwise> <!-- Current address -->
																				
																			</xsl:choose> <!-- Redemption Status -->
																			
																		</xsl:element> <!-- MORTGAGELOAN -->
																		
																	</xsl:element> <!-- MORTGAGEACCOUNT -->
																	
																	<!-- Only do complicated account relationships if its the current address (see after customers are created), otherwise assume its just the single customer -->
																	<xsl:if test="AddressTypeId != 2">
																		<!-- Assign the current applicant -->
																		<!-- ACCOUNT RELATIONSHIP -->
																		<xsl:element name="ACCOUNTRELATIONSHIP">
																			<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
																			<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
																			<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant role type -->
																			<xsl:attribute name="ACCOUNTGUID">../@ACCOUNTGUID</xsl:attribute>
																				
																		</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
																		
																	</xsl:if> <!-- Not the current address -->
																	
																	<xsl:element name="INDEMNITYINSURANCE"/>
																	
																</xsl:element> <!-- ACCOUNT -->
																
															</xsl:element> <!-- THIRDPARTY -->
															
														</xsl:element> <!-- ADDRESS -->
														
													</xsl:element> <!-- CONTACTDETAILS -->
													
													<!-- Is there a second charge on this mortgage -->
													<xsl:if test="MortgageSecondCharge = 1">
														<!-- LENDERDETAILS -->
														<xsl:element name="SECONDCHARGELENDERDETAILS">
															<xsl:attribute name="CONTACTFORENAME"></xsl:attribute>
															<xsl:attribute name="CONTACTSURNAME"></xsl:attribute>
															<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
															<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
															<xsl:attribute name="EMAILADDRESS"></xsl:attribute>
															
															<!-- SECONDCHARGELENDERTELEPHONEDETAILS -->
															<xsl:element name="SECONDCHARGELENDERTELEPHONEDETAILS">
																<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
																<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="SecondChgAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="SecondChgAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																
															</xsl:element> <!-- SECONDCHARGELENDERTELEPHONEDETAILS -->
														
															<!-- SECONDCHARGELENDERTELEPHONEDETAILS -->
															<xsl:element name="SECONDCHARGELENDERTELEPHONEDETAILS">
																<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
																<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="SecondChgAddressFaxNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="SecondChgAddressFaxNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																
															</xsl:element> <!-- SECONDCHARGELENDERTELEPHONEDETAILS -->
															
															<!-- LENDERADDRESS -->
															<xsl:element name="LENDERADDRESS">
																<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="SecondChgAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="SecondChgAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="FLATNUMBER"></xsl:attribute>
																<xsl:attribute name="STREET"><xsl:value-of select="SecondChgAddressBuildingName"/></xsl:attribute>
																<xsl:attribute name="DISTRICT"><xsl:value-of select="SecondChgAddressStreet"/></xsl:attribute>
																<xsl:attribute name="TOWN"><xsl:value-of select="SecondChgAddressTown"/></xsl:attribute>
																<xsl:attribute name="COUNTY"><xsl:value-of select="SecondChgAddressCounty"/></xsl:attribute>
																<xsl:attribute name="COUNTRY"><xsl:value-of select="SecondChgAddressCountryTypeId"/></xsl:attribute>
																<xsl:attribute name="POSTCODE"><xsl:value-of select="SecondChgAddressPostCode"/></xsl:attribute>
																
																<!-- LENDERTHIRDPARTY -->
																<xsl:element name="LENDERTHIRDPARTY">
																	<xsl:attribute name="THIRDPARTYTYPE">3</xsl:attribute> <!-- Other Lender Thirdparty type -->
																	<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="SecondChgLenderName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
																	
																	<!-- ACCOUNT DETAILS -->
																	<xsl:element name="ACCOUNT">
																		<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="SecondChgAccountNo"/></xsl:attribute>
																			
																		<!-- MORTGAGEACCOUNT -->
																		<xsl:element name="MORTGAGEACCOUNT">
																			<xsl:attribute name="MONTHLYRENTALINCOME"></xsl:attribute>
																			<xsl:attribute name="ADDRESSGUID"><xsl:text>../../../../../@ADDRESSGUID</xsl:text></xsl:attribute> <!-- This key is inherited from an x path -->
																				
																			<!-- MORTGAGELOANS -->
																			<xsl:element name="MORTGAGELOAN">
																				<xsl:attribute name="LOANACCOUNTNUMBER"><xsl:value-of select="SecondChgAccountNo"/></xsl:attribute>
																				<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="SecondChgMonthlyPayment"/></xsl:attribute> 
																				<xsl:attribute name="OUTSTANDINGBALANCE"><xsl:value-of select="SecondChgOutstandingAmount"/></xsl:attribute> 
																				<xsl:attribute name="PURPOSEOFLOAN"></xsl:attribute>
																			
																			</xsl:element> <!-- MORTGAGELOAN -->
																			
																		</xsl:element> <!-- MORTGAGEACCOUNT -->
																		
																		<!-- INDEMNITYINSURANCE -->
																		<xsl:element name="INDEMNITYINSURANCE"/>
																		
																	</xsl:element> <!-- ACCOUNT -->
																	
																</xsl:element> <!-- LENDERTHIRDPARTY -->
																
															</xsl:element> <!-- LENDERADDRESS -->
															
														</xsl:element> <!-- SECONDCHARGELENDERDETAILS -->
															
													</xsl:if> <!-- Is there a second charge on this mortgage -->
													
													<!-- Have there been any previous mortgages for this address in the last 12 months -->
													<xsl:if test="PrevMortgageYN = 1">
														
														<!-- Loop through the previous lenders -->
														<xsl:for-each select="PreviousLenderRoot/PreviousLender">
														
															<!-- LENDERDETAILS -->
															<xsl:element name="PREVIOUSLENDERDETAILS">
																<xsl:attribute name="CONTACTFORENAME"></xsl:attribute>
																<xsl:attribute name="CONTACTSURNAME"></xsl:attribute>
																<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
																<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
																<xsl:attribute name="EMAILADDRESS"></xsl:attribute>
																
																<!-- PREVIOUSLENDERTELEPHONEDETAILS -->
																<xsl:element name="PREVIOUSLENDERTELEPHONEDETAILS">
																	<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
																	<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																	<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																	
																</xsl:element> <!-- PREVIOUSLENDERTELEPHONEDETAILS -->
														
																<!-- PREVIOUSLENDERTELEPHONEDETAILS -->
																<xsl:element name="PREVIOUSLENDERTELEPHONEDETAILS">
																	<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
																	<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																	<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																	
																</xsl:element> <!-- PREVIOUSLENDERTELEPHONEDETAILS -->
																
																<!-- LENDERADDRESS -->
																<xsl:element name="LENDERADDRESS">
																	<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LenderAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LenderAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="FLATNUMBER"></xsl:attribute>
																	<xsl:attribute name="STREET"><xsl:value-of select="LenderAddressBuildingName"/></xsl:attribute>
																	<xsl:attribute name="DISTRICT"><xsl:value-of select="LenderAddressStreet"/></xsl:attribute>
																	<xsl:attribute name="TOWN"><xsl:value-of select="LenderAddressTown"/></xsl:attribute>
																	<xsl:attribute name="COUNTY"><xsl:value-of select="LenderAddressCounty"/></xsl:attribute>
																	<xsl:attribute name="COUNTRY"><xsl:value-of select="LenderAddressCountryTypeId"/></xsl:attribute>
																	<xsl:attribute name="POSTCODE"><xsl:value-of select="LenderAddressPostCode"/></xsl:attribute>
																	
																	<!-- LENDERTHIRDPARTY -->
																	<xsl:element name="LENDERTHIRDPARTY">
																		<xsl:attribute name="THIRDPARTYTYPE">3</xsl:attribute> <!-- Other Lender Thirdparty type -->
																		<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="MortgageLenderName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
																		
																		<!-- ACCOUNT DETAILS -->
																		<xsl:element name="ACCOUNT">
																			<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="MortgageAccountNo"/></xsl:attribute>
																				
																			<!-- MORTGAGEACCOUNT -->
																			<xsl:element name="MORTGAGEACCOUNT">
																				<xsl:attribute name="MONTHLYRENTALINCOME"></xsl:attribute>
																				<xsl:attribute name="ADDRESSGUID"><xsl:text>../../../../../@ADDRESSGUID</xsl:text></xsl:attribute> <!-- This key is inherited from an x path -->
																					
																				<!-- MORTGAGELOANS -->
																				<xsl:element name="MORTGAGELOAN">
																					<xsl:attribute name="LOANACCOUNTNUMBER"><xsl:value-of select="MortgageAccountNo"/></xsl:attribute>
																					<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MortgageMonthlyPayment"/></xsl:attribute> 
																					<xsl:attribute name="OUTSTANDINGBALANCE"><xsl:value-of select="MortgageBalanceOnRedemption"/></xsl:attribute>
																					<xsl:attribute name="REDEMPTIONSTATUS">40</xsl:attribute>  <!-- Already redeemed -->
																					<xsl:attribute name="REDEMPTIONDATE"><xsl:value-of select="MortgageDateOfRedemption"/></xsl:attribute>
																					<xsl:attribute name="STARTDATE"><xsl:value-of select="MortgageStartDate"/></xsl:attribute>  
																					<xsl:attribute name="PURPOSEOFLOAN"></xsl:attribute>
																				
																				</xsl:element> <!-- MORTGAGELOAN -->
																				
																			</xsl:element> <!-- MORTGAGEACCOUNT -->
																			
																			<!-- INDEMNITYINSURANCE -->
																			<xsl:element name="INDEMNITYINSURANCE"/>
																			
																		</xsl:element> <!-- ACCOUNT -->
																		
																	</xsl:element> <!-- LENDERTHIRDPARTY -->
																	
																</xsl:element> <!-- LENDERADDRESS -->
																
															</xsl:element> <!-- PREVIOUSLENDERDETAILS -->
														
														</xsl:for-each> <!-- Loop through the previous lenders -->
														
													</xsl:if> <!-- Have there been any previous mortgages for this address in the last 12 months -->
													
												</xsl:if> <!-- Is there a mortgage account associated with this address? Is it owner occupied, have a mortgage = yes and if its a current address then the sameasapplicantflag is different  -->
	
												<!-- Is there a landlord/tenancy associated with this address? -->
												<xsl:if test="(ResidentialStatusTypeId = 3 or ResidentialStatusTypeId = 4)"> <!-- This is local authority renting and private renting  -->
	
													<!-- LANDLORDDETAILS -->
													<xsl:element name="LENDERLANDLORDDETAILS">
														<xsl:attribute name="CONTACTFORENAME"></xsl:attribute>
														<xsl:attribute name="CONTACTSURNAME"></xsl:attribute>
														<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
														<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
														<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="LandLordAddressEmail"/></xsl:attribute>
														
														<!-- LENDERLANDLORDTELEPHONEDETAILS -->
														<xsl:element name="LENDERLANDLORDTELEPHONEDETAILS">
															<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
															<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
														</xsl:element> <!-- LENDERLANDLORDTELEPHONEDETAILS -->
														
														<!-- LENDERLANDLORDTELEPHONEDETAILS -->
														<xsl:element name="LENDERLANDLORDTELEPHONEDETAILS">
															<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
															<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
															<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="LenderAddressFaxNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
														</xsl:element> <!-- LENDERLANDLORDTELEPHONEDETAILS -->
														
														<!-- LANDLORDADDRESS -->
														<xsl:element name="LENDERLANDLORDADDRESS">
															<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LandLordAddressBuildingNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="LandLordAddressBuildingNo" /></xsl:call-template></xsl:attribute>
															<xsl:attribute name="FLATNUMBER"></xsl:attribute>
															<xsl:attribute name="STREET"><xsl:value-of select="LandLordAddressBuildingName"/></xsl:attribute>
															<xsl:attribute name="DISTRICT"><xsl:value-of select="LandLordAddressStreet"/></xsl:attribute>
															<xsl:attribute name="TOWN"><xsl:value-of select="LandLordAddressTown"/></xsl:attribute>
															<xsl:attribute name="COUNTY"><xsl:value-of select="LandLordAddressCounty"/></xsl:attribute>
															<xsl:attribute name="COUNTRY"><xsl:value-of select="LandLordAddressCountryTypeId"/></xsl:attribute>
															<xsl:attribute name="POSTCODE"><xsl:value-of select="LandLordAddressPostCode"/></xsl:attribute>
														
															<!-- LANDLORDCOMPANYDETAILS -->
															<xsl:element name="LENDERLANDLORDTHIRDPARTY">
																<xsl:attribute name="THIRDPARTYTYPE">8</xsl:attribute> <!-- Landlord Thirdparty type -->
																<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="LandLordName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
																
																<!-- TENANCYDETAILS -->
																<xsl:element name="TENANCY">
																	<xsl:attribute name="TENANCYTYPE"><xsl:value-of select="ResidentialStatusTypeId"/></xsl:attribute>
																	<xsl:attribute name="MONTHLYRENTAMOUNT"><xsl:value-of select="LandLordMonthlyRent"/></xsl:attribute>
																	<xsl:attribute name="ACCOUNTNUMBER"></xsl:attribute>
																</xsl:element> <!-- TENANCY -->
																
															</xsl:element> <!-- LENDERLANDLORDTHIRDPARTY -->
															
														</xsl:element> <!-- LENDERLANDLORDADDRESS -->
														
													</xsl:element> <!-- LENDERLANDLORDDETAILS -->
													
												</xsl:if> <!-- Is there a loandlord associated with this request -->
												
											</xsl:element> <!-- CUSTOMERADDRESS -->
											
										</xsl:element> <!-- ADDRESS -->
										
									</xsl:for-each> <!-- ApplicantRoot/Applicant -->
	
									<!--  Correspondence address -->
									<xsl:if test="CorrespondenceAddressRequiredYN=1">
	
										<!-- CORRESPONDENCEADDRESS -->
										<xsl:element name="ADDRESS">
											<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="CorrespondenceAddressBuildingNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="CorrespondenceAddressBuildingNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="FLATNUMBER"></xsl:attribute>
											<xsl:attribute name="STREET"><xsl:value-of select="CorrespondenceAddressBuildingName"/></xsl:attribute>
											<xsl:attribute name="DISTRICT"><xsl:value-of select="CorrespondenceAddressStreet"/></xsl:attribute>
											<xsl:attribute name="TOWN"><xsl:value-of select="CorrespondenceAddressTown"/></xsl:attribute>
											<xsl:attribute name="COUNTY"><xsl:value-of select="CorrespondenceAddressCounty"/></xsl:attribute>
											<xsl:attribute name="COUNTRY"><xsl:value-of select="CorrespondenceAddressCountryTypeId"/></xsl:attribute>
											<xsl:attribute name="POSTCODE"><xsl:value-of select="CorrespondenceAddressPostCode"/></xsl:attribute>
											
											<!-- CUSTOMERADDRESS -->
											<xsl:element name="CUSTOMERADDRESS">
												<xsl:attribute name="ADDRESSTYPE">99</xsl:attribute> <!-- Correspondence Address -->
												<xsl:attribute name="DATEMOVEDIN"></xsl:attribute>
												<xsl:attribute name="DATEMOVEDOUT"></xsl:attribute>
												<xsl:attribute name="NATUREOFOCCUPANCY"></xsl:attribute>
												
											</xsl:element> <!-- CUSTOMERADDRESS -->
											
										</xsl:element> <!-- ADDRESS -->
										
									</xsl:if>
									
									<!-- Telephone Numbers -->
									<!-- If they have a correspondence address, use these telephone numbers, else use the current address -->
									<xsl:choose>
										<!-- Correspondence address -->
										<xsl:when test="CorrespondenceAddressRequiredYN = 1">
											
											<!-- Home telephone number -->
											<xsl:if test="string-length(CorrespondenceAddressTelephoneNo) > 0">
												<!-- CUSTOMERTELEPHONENUMBER -->
												<xsl:element name="CUSTOMERTELEPHONENUMBER">
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="CorrespondenceAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="CorrespondenceAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="USAGE">1</xsl:attribute> <!-- Home -->
													<xsl:attribute name="PREFERREDMETHODOFCONTACT">1</xsl:attribute> <!-- Set home telephone number to preferred method -->
														
												</xsl:element> <!-- CUSTOMERTELEPHONENUMBER -->
												
											</xsl:if>
												
											<!-- Mobile telephone number -->
											<xsl:if test="string-length(CorrespondenceAddressMobileNo) > 0">
												<!-- CUSTOMERTELEPHONENUMBER -->
												<xsl:element name="CUSTOMERTELEPHONENUMBER">
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="CorrespondenceAddressMobileNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="CorrespondenceAddressMobileNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="USAGE">3</xsl:attribute> <!-- Mobile -->
													<xsl:attribute name="PREFERREDMETHODOFCONTACT">0</xsl:attribute> 

												</xsl:element> <!-- CUSTOMERTELEPHONENUMBER -->
												
											</xsl:if>
											
										</xsl:when> <!-- Correspondence address -->
										
										<!-- Current address -->
										<xsl:otherwise>
											
											<xsl:for-each select="ApplicantAddressRoot/ApplicantAddress[AddressTypeId=2]">  <!-- Current Address -->
												
												<!-- Home telephone number -->
												<xsl:if test="string-length(ApplicantAddressTelephoneNo) > 0">
													<!-- CUSTOMERTELEPHONENUMBER -->
													<xsl:element name="CUSTOMERTELEPHONENUMBER">
														<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
														<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
														<xsl:attribute name="TELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
														<xsl:attribute name="USAGE">1</xsl:attribute> <!-- Home -->
														<xsl:attribute name="PREFERREDMETHODOFCONTACT">1</xsl:attribute> <!-- Set home telephone number to preferred method -->
														
													</xsl:element> <!-- CUSTOMERTELEPHONENUMBER -->
												
												</xsl:if>
												
												<!-- Mobile telephone number -->
												<xsl:if test="string-length(ApplicantAddressMobileNo) > 0">
													<!-- CUSTOMERTELEPHONENUMBER -->
													<xsl:element name="CUSTOMERTELEPHONENUMBER">
														<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
														<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantAddressMobileNo" /></xsl:call-template></xsl:attribute>
														<xsl:attribute name="TELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantAddressMobileNo" /></xsl:call-template></xsl:attribute>
														<xsl:attribute name="USAGE">3</xsl:attribute> <!-- Mobile -->
														<xsl:attribute name="PREFERREDMETHODOFCONTACT">0</xsl:attribute> 
														
													</xsl:element> <!-- CUSTOMERTELEPHONENUMBER -->
												
												</xsl:if>
												
											</xsl:for-each> <!-- Select current address -->
										
										</xsl:otherwise> <!-- Current address -->
										
									</xsl:choose>  <!-- If they have a correspondence address, use these telephone numbers, else use the current address -->
									
									<!-- Work Telephone Number -->
									<!-- Take the current employments telephone number, if one exists -->
									<xsl:if test="string-length(ApplicantEmploymentRoot/ApplicantEmployment[1]/EmployedAddressTelephoneNo) > 0">
										<!-- CUSTOMERTELEPHONENUMBER -->
										<xsl:element name="CUSTOMERTELEPHONENUMBER">
											<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
											<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantEmploymentRoot/ApplicantEmployment[1]/EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="TELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ApplicantEmploymentRoot/ApplicantEmployment[1]/EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
											<xsl:attribute name="USAGE">2</xsl:attribute> <!-- Work -->
											<xsl:attribute name="PREFERREDMETHODOFCONTACT">0</xsl:attribute> 
													
										</xsl:element> <!-- CUSTOMERTELEPHONENUMBER -->
										
									</xsl:if>
									
									<!-- For each Employment or Self employment -->
									<xsl:for-each select="ApplicantEmploymentRoot/ApplicantEmployment[EmploymentTypeId=2 or EmploymentTypeId=3 or EmploymentTypeId=4]"> <!-- Employed, Self-employed or Unemployed -->
									
										<!-- No Accountant details, employed and not related to the employer, or self employed and no accountant -->
										<xsl:if test="(RelatedToEmployerYN != '1' and EmploymentTypeId='2') or (ConfirmationAddressYN != '1' and EmploymentTypeId='3')">
											
											<!-- EMPLOYMENTDETAILS -->
											<xsl:element name="EMPLOYMENTDETAILS">
												<xsl:attribute name="CONTACTFORENAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'FORENAME'" /><xsl:with-param name="FULLNAME" select="ContactName" /></xsl:call-template></xsl:attribute>
												<xsl:attribute name="CONTACTSURNAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'SURNAME'" /><xsl:with-param name="FULLNAME" select="ContactName" /></xsl:call-template></xsl:attribute>
												<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
												<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
												<xsl:attribute name="EMAILADDRESS"></xsl:attribute>
												<xsl:attribute name="FAXNUMBER"><xsl:value-of select="EmployedAddressFaxNo"/></xsl:attribute>
												<xsl:attribute name="TELEPHONEEXTENSIONNUMBER"></xsl:attribute>
												<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="EmployedAddressTelephoneNo"/></xsl:attribute>
												
												<!-- EMPLOYERTELEPHONEDETAILS -->
												<xsl:element name="EMPLOYERTELEPHONEDETAILS">
													<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
												</xsl:element> <!-- EMPLOYERTELEPHONEDETAILS -->
														
												<!-- EMPLOYERTELEPHONEDETAILS -->
												<xsl:element name="EMPLOYERTELEPHONEDETAILS">
													<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressFaxNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressFaxNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
												</xsl:element> <!-- EMPLOYERTELEPHONEDETAILS -->
												
												<!-- EMPLOYERCOMPANYADDRESS -->
												<xsl:element name="EMPLOYERADDRESS">
													
													<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="EmployedAddressBuildingNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="EmployedAddressBuildingNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="FLATNUMBER"></xsl:attribute>
													<xsl:attribute name="STREET"><xsl:value-of select="EmployedAddressBuildingName"/></xsl:attribute>
													<xsl:attribute name="DISTRICT"><xsl:value-of select="EmployedAddressStreet"/></xsl:attribute>
													<xsl:attribute name="TOWN"><xsl:value-of select="EmployedAddressTown"/></xsl:attribute>
													<xsl:attribute name="COUNTY"><xsl:value-of select="EmployedAddressCounty"/></xsl:attribute>
													<xsl:attribute name="COUNTRY"><xsl:value-of select="EmployedAddressCountryTypeId"/></xsl:attribute>
													<xsl:attribute name="POSTCODE"><xsl:value-of select="EmployedAddressPostCode"/></xsl:attribute>
													
													<!-- EMPLOYERCOMPANYDETAILS -->
													<xsl:element name="EMPLOYERTHIRDPARTY">
														<xsl:attribute name="THIRDPARTYTYPE">5</xsl:attribute> <!-- Employer Thirdparty type -->
														<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="EmployerName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
														
														<!-- EMPLOYMENT -->
														<xsl:element name="EMPLOYMENT">	
															<xsl:attribute name="DATELEFTORCEASEDTRADING"><xsl:value-of select="DateCompleted"/></xsl:attribute>
															<xsl:attribute name="DATESTARTEDORESTABLISHED"><xsl:value-of select="DateStarted"/></xsl:attribute>
															<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="EmploymentTypeId"/></xsl:attribute>
															<xsl:attribute name="EMPLOYMENTTYPE"><xsl:value-of select="EmployedTypeId"/></xsl:attribute>
															<xsl:attribute name="INDUSTRYTYPE"></xsl:attribute>
															<xsl:attribute name="JOBTITLE"><xsl:value-of select="Occupation"/></xsl:attribute>
															<xsl:attribute name="MAINSTATUS">
																<xsl:choose>
																	<!-- If its the first employment node, then its the main status -->
																	<xsl:when test="position() = 1"><xsl:text>1</xsl:text></xsl:when>
																	<xsl:otherwise><xsl:text>0</xsl:text></xsl:otherwise>
																</xsl:choose>
															</xsl:attribute>
															<xsl:attribute name="OCCUPATIONTYPE"></xsl:attribute>
															<xsl:attribute name="NETMONTHLYINCOME">1</xsl:attribute> <!-- This value needs to be removed after the income calcs have been done -->
															<xsl:attribute name="OTHEREMPLOYMENTSTATUS"></xsl:attribute>
															<xsl:attribute name="SHARESOWNEDINDICATOR"></xsl:attribute>
															
															<!-- Is Applicant Employed, i.e. Not self employed -->
															<xsl:if test="EmploymentTypeId='2'">
															
																<!-- EMPLOYEDDETAILS -->
																<xsl:element name="EMPLOYEDDETAILS">	
																	<xsl:attribute name="NOTICEPROBLEMINDICATOR"></xsl:attribute>
																	<xsl:attribute name="PAYROLLNUMBER"><xsl:value-of select="StaffNo"/></xsl:attribute>
																	<xsl:attribute name="PERCENTSHARESHELD"></xsl:attribute>
																	<xsl:attribute name="PROBATIONARYINDICATOR"></xsl:attribute>
																	<xsl:attribute name="P60SEENINDICATOR"></xsl:attribute>
																	<xsl:attribute name="WAGESLIPSSEENINDICATOR"></xsl:attribute>
																	<xsl:attribute name="EMPLOYMENTRELATIONSHIPIND">0</xsl:attribute> <!-- They arn't related to the employer -->
																	
																	<!-- DO INCOMES, BASIC, REGULAR, BONUS or only TOTAL -->
																	
																	<!-- Have Full income details or only TOTAL income details been captured? 
																	Full income details are captured unless its a Self Certification, SubPrime - Self Cert, 
																	or Buy to Let Mortgage (and if its BTL the applicant is not a first time buyer) -->
																	<xsl:choose>
																		
																		<!-- Only Total Income -->
																		<xsl:when test="/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 3 or /REQUEST/ProposalRoot/Proposal/MortgageTypeId = 8 or (/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 2 and ../../FirstTimeBuyerYN = 0)">
																		
																			<!-- EARNEDINCOME -->
																			<xsl:element name="EARNEDINCOME">
																				<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="TotalAnnualIncome"/></xsl:attribute>
																				<xsl:attribute name="EARNEDINCOMETYPE">1</xsl:attribute> <!-- 1 = Basic IncomeType -->
																				<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																			</xsl:element> <!-- EARNEDINCOME -->
																			
																		</xsl:when> <!-- Only Total Income -->
																		
																		<xsl:otherwise>
																		
																			<!-- Is there a basic income -->
																			<xsl:if test="number(BasicAnnualIncome) > 0">
																				<!-- EARNEDINCOME -->
																				<xsl:element name="EARNEDINCOME">
																				
																					<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:call-template name="CalculateBasicIncome"><xsl:with-param name="BasicAnnualIncome" select="BasicAnnualIncome" /><xsl:with-param name="OtherIncome" select="OtherIncome" /><xsl:with-param name="OtherIncomeTypeId" select="OtherIncomeTypeId" /></xsl:call-template></xsl:attribute>
																					<xsl:attribute name="EARNEDINCOMETYPE">1</xsl:attribute> <!-- 1 = Basic IncomeType -->
																					<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																				</xsl:element> <!-- EARNEDINCOME -->
																	
																			</xsl:if> <!-- Basic income -->
																	
																			<!-- Is there a regular overtime income -->
																			<xsl:if test="number(RegularOvertimeAmount) > 0">
																				<!-- EARNEDINCOME -->
																				<xsl:element name="EARNEDINCOME">
																					<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="RegularOvertimeAmount"/></xsl:attribute>
																					<xsl:attribute name="EARNEDINCOMETYPE">3</xsl:attribute> <!-- 3 = Regular overtime IncomeType -->
																					<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																				</xsl:element> <!-- EARNEDINCOME -->
																	
																			</xsl:if> <!-- Regular overtime income -->
																	
																			<!-- Is there a regular bonus or commission income -->
																			<xsl:if test="number(RegularBonusCommissionAmount) > 0">
																				<!-- EARNEDINCOME -->
																				<xsl:element name="EARNEDINCOME">
																					<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="RegularBonusCommissionAmount"/></xsl:attribute>
																					<xsl:attribute name="EARNEDINCOMETYPE">5</xsl:attribute> <!-- 5 = Regular bonus IncomeType -->
																					<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																				</xsl:element> <!-- EARNEDINCOME -->
																	
																			</xsl:if> <!-- Regular overtime income -->
																			
																		</xsl:otherwise> <!-- Full Income Details -->
																		
																	</xsl:choose>
																	
																</xsl:element> <!-- EMPLOYEDDETAILS -->
															
															</xsl:if> <!-- Applicant is employed -->
															
															<!-- Applicant Employment is self employed -->
															<xsl:if test="EmploymentTypeId='3'">
															
																<!-- SELFEMPLOYEDDETAILS -->
																<xsl:element name="SELFEMPLOYEDDETAILS">
																	<xsl:attribute name="DATEFINANCIALINTERESTHELD"></xsl:attribute>
																	<xsl:attribute name="OTHERBUSINESSCONNECTIONS"></xsl:attribute>
																	<xsl:attribute name="PERCENTSHARESHELD"></xsl:attribute>
																	<xsl:attribute name="REGISTRATIONNUMBER"><xsl:value-of select="CompanyRegNo"/></xsl:attribute>
																	<xsl:attribute name="VATNUMBER"><xsl:value-of select="VATNo"/></xsl:attribute>
																	<xsl:attribute name="NATUREOFBUSINESS"><xsl:value-of select="NatureOfBusiness"/></xsl:attribute>
																	
																	<!-- Have Full income details or only TOTAL income details been captured? 
																	Full income details are captured unless its a Self Certification, SubPrime - Self Cert, 
																	or Buy to Let Mortgage (and if its BTL the applicant is not a first time buyer) -->
																	<xsl:choose>
																		
																		<!-- Only Total Income -->
																		<xsl:when test="/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 3 or /REQUEST/ProposalRoot/Proposal/MortgageTypeId = 8 or (/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 2 and ../../FirstTimeBuyerYN = 0)">
																		
																			<!-- NETPROFIT -->
																			<xsl:element name="NETPROFIT">
																				<xsl:attribute name="EVIDENCEOFACCOUNTS"></xsl:attribute>
																				<xsl:attribute name="YEAR1"><xsl:call-template name="OutputCurrentYear"><xsl:with-param name="NOW" select="$Now" /></xsl:call-template></xsl:attribute>
																				<xsl:attribute name="YEAR1AMOUNT"><xsl:value-of select="TotalAnnualAmount"/></xsl:attribute>
																			</xsl:element> <!-- NETPROFIT -->
																		
																		</xsl:when> <!-- Only Total Income -->
																		
																		<!-- Full Income details -->
																		<xsl:otherwise>
																			
																			<!-- NETPROFIT -->
																			<xsl:element name="NETPROFIT">
																				<xsl:attribute name="EVIDENCEOFACCOUNTS"></xsl:attribute>
																				<xsl:attribute name="YEAR1"><xsl:value-of select="Year1Year"/></xsl:attribute>
																				<xsl:attribute name="YEAR1AMOUNT"><xsl:value-of select="NetProfitsYear1"/></xsl:attribute>
																				<xsl:attribute name="YEAR2"></xsl:attribute>
																				<xsl:attribute name="YEAR2AMOUNT"></xsl:attribute>
																				<xsl:attribute name="YEAR3"></xsl:attribute>
																				<xsl:attribute name="YEAR3AMOUNT"></xsl:attribute>
																			</xsl:element> <!-- NETPROFIT -->
																		
																		</xsl:otherwise> <!-- Full Income details -->
																		
																	</xsl:choose>
																	
																</xsl:element> <!-- SELFEMPLOYEDDETAILS -->
																				
															</xsl:if> <!-- Applicant Employment is self employed -->
															
														</xsl:element> <!-- EMPLOYMENT -->
														
													</xsl:element> <!-- EMPLOYERTHIRDPARTY -->
													
												</xsl:element> <!-- EMPLOYERADDRESS -->
												
											</xsl:element> <!-- EMPLOYMENTDETAILS -->
											
										</xsl:if> <!-- No Accountant details -->

										<!-- Has Accountant details, employed and related to the employer, or self employed with an accountant-->
										<xsl:if test="(RelatedToEmployerYN='1' and EmploymentTypeId='2') or (ConfirmationAddressYN='1' and EmploymentTypeId='3')">
											
											<!-- EMPLOYMENTDETAILSWITHACCOUNTANT -->
											<xsl:element name="EMPLOYMENTDETAILSWITHACCOUNTANT">
												<xsl:attribute name="CONTACTFORENAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'FORENAME'" /><xsl:with-param name="FULLNAME" select="ConfirmationContactName" /></xsl:call-template></xsl:attribute>
												<xsl:attribute name="CONTACTSURNAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'SURNAME'" /><xsl:with-param name="FULLNAME" select="ConfirmationContactName" /></xsl:call-template></xsl:attribute>
												<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
												<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
												<xsl:attribute name="EMAILADDRESS"></xsl:attribute>
												
												<!-- ACCOUNTANTTELEPHONEDETAILS -->
												<xsl:element name="ACCOUNTANTTELEPHONEDETAILS">
													<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ConfirmationAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ConfirmationAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
												</xsl:element> <!-- ACCOUNTANTTELEPHONEDETAILS -->
														
												<!-- ACCOUNTANTTELEPHONEDETAILS -->
												<xsl:element name="ACCOUNTANTTELEPHONEDETAILS">
													<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
													<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
													<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ConfirmationAddressFaxNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="ConfirmationAddressFaxNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
															
												</xsl:element> <!-- ACCOUNTANTTELEPHONEDETAILS -->
												
												<!-- ACCOUNTANTADDRESS -->
												<xsl:element name="ACCOUNTANTADDRESS">
													<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="ConfirmationAddressBuildingNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="ConfirmationAddressBuildingNo" /></xsl:call-template></xsl:attribute>
													<xsl:attribute name="FLATNUMBER"></xsl:attribute>
													<xsl:attribute name="STREET"><xsl:value-of select="ConfirmationAddressBuildingName"/></xsl:attribute>
													<xsl:attribute name="DISTRICT"><xsl:value-of select="ConfirmationAddressStreet"/></xsl:attribute>
													<xsl:attribute name="TOWN"><xsl:value-of select="ConfirmationAddressTown"/></xsl:attribute>
													<xsl:attribute name="COUNTY"><xsl:value-of select="ConfirmationAddressCounty"/></xsl:attribute>
													<xsl:attribute name="COUNTRY"><xsl:value-of select="ConfirmationAddressCountryTypeId"/></xsl:attribute>
													<xsl:attribute name="POSTCODE"><xsl:value-of select="ConfirmationAddressPostCode"/></xsl:attribute>
												
													<!-- ACCOUNTANTCOMPANYDETAILS -->
													<xsl:element name="ACCOUNTANTTHIRDPARTY">
														<xsl:attribute name="THIRDPARTYTYPE">5</xsl:attribute> <!-- Accountant Thirdparty type -->
														<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="ConfirmationName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
														
														<!-- ACCOUNTANT -->
														<xsl:element name="ACCOUNTANT">
															<xsl:attribute name="ACCOUNTANCYFIRMNAME"><xsl:value-of select="ConfirmationName"/></xsl:attribute>
															<xsl:attribute name="QUALIFICATIONS"></xsl:attribute>
															<xsl:attribute name="YEARSACTINGFORCUSTOMER"></xsl:attribute>
														
															<!-- EMPLOYMENTDETAILS -->
															<xsl:element name="EMPLOYMENTDETAILS">
																<xsl:attribute name="CONTACTFORENAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'FORENAME'" /><xsl:with-param name="FULLNAME" select="ContactName" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="CONTACTSURNAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'SURNAME'" /><xsl:with-param name="FULLNAME" select="ContactName" /></xsl:call-template></xsl:attribute>
																<xsl:attribute name="CONTACTTITLE"></xsl:attribute>
																<xsl:attribute name="CONTACTTYPE">10</xsl:attribute> <!-- Organisation Contact Type -->
																<xsl:attribute name="EMAILADDRESS"></xsl:attribute>
																
																<!-- EMPLOYERTELEPHONEDETAILS -->
																<xsl:element name="EMPLOYERTELEPHONEDETAILS">
																	<xsl:attribute name="USAGE">10</xsl:attribute> <!-- Work (ContactTelephoneUsage) -->
																	<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																	<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressTelephoneNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																			
																</xsl:element> <!-- EMPLOYERTELEPHONEDETAILS -->
																		
																<!-- EMPLOYERTELEPHONEDETAILS -->
																<xsl:element name="EMPLOYERTELEPHONEDETAILS">
																	<xsl:attribute name="USAGE">20</xsl:attribute> <!-- Fax (ContactTelephoneUsage) -->
																	<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
																	<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressFaxNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="TELENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="EmployedAddressFaxNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
																			
																</xsl:element> <!-- EMPLOYERTELEPHONEDETAILS -->
																
																<!-- EMPLOYERADDRESS -->
																<xsl:element name="EMPLOYERADDRESS">
																	<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="EmployedAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="EmployedAddressBuildingNo" /></xsl:call-template></xsl:attribute>
																	<xsl:attribute name="FLATNUMBER"></xsl:attribute>
																	<xsl:attribute name="STREET"><xsl:value-of select="EmployedAddressBuildingName"/></xsl:attribute>
																	<xsl:attribute name="DISTRICT"><xsl:value-of select="EmployedAddressStreet"/></xsl:attribute>
																	<xsl:attribute name="TOWN"><xsl:value-of select="EmployedAddressTown"/></xsl:attribute>
																	<xsl:attribute name="COUNTY"><xsl:value-of select="EmployedAddressCounty"/></xsl:attribute>
																	<xsl:attribute name="COUNTRY"><xsl:value-of select="EmployedAddressCountryTypeId"/></xsl:attribute>
																	<xsl:attribute name="POSTCODE"><xsl:value-of select="EmployedAddressPostCode"/></xsl:attribute>
																	
																	<!-- EMPLOYERCOMPANYDETAILS -->
																	<xsl:element name="EMPLOYERTHIRDPARTY">
																		<xsl:attribute name="THIRDPARTYTYPE">5</xsl:attribute> <!-- Employer Thirdparty type -->
																		<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="EmployerName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
																		
																		<!-- EMPLOYMENT -->
																		<xsl:element name="EMPLOYMENT">
																			<xsl:attribute name="DATELEFTORCEASEDTRADING"><xsl:value-of select="DateCompleted"/></xsl:attribute>
																			<xsl:attribute name="DATESTARTEDORESTABLISHED"><xsl:value-of select="DateStarted"/></xsl:attribute>
																			<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="EmploymentTypeId"/></xsl:attribute>
																			<xsl:attribute name="EMPLOYMENTTYPE"><xsl:value-of select="EmployedTypeId"/></xsl:attribute>
																			<xsl:attribute name="INDUSTRYTYPE"></xsl:attribute>
																			<xsl:attribute name="JOBTITLE"><xsl:value-of select="Occupation"/></xsl:attribute>
																			<xsl:attribute name="MAINSTATUS">
																				<xsl:choose>
																					<!-- If its the first employment node, then its the main status -->
																					<xsl:when test="position() = 1"><xsl:text>1</xsl:text></xsl:when>
																					<xsl:otherwise><xsl:text>0</xsl:text></xsl:otherwise>
																				</xsl:choose>
																			</xsl:attribute>
																			<xsl:attribute name="OCCUPATIONTYPE"></xsl:attribute>
																			<xsl:attribute name="NETMONTHLYINCOME"></xsl:attribute>
																			<xsl:attribute name="OTHEREMPLOYMENTSTATUS"></xsl:attribute>
																			<xsl:attribute name="SHARESOWNEDINDICATOR"></xsl:attribute>
																			
																			<!-- Is Applicant Employed, i.e. Not self employed -->
																			<xsl:if test="EmploymentTypeId='2'">
																			
																				<!-- EMPLOYEDDETAILS -->
																				<xsl:element name="EMPLOYEDDETAILS">
																					<xsl:attribute name="NOTICEPROBLEMINDICATOR"></xsl:attribute>
																					<xsl:attribute name="PAYROLLNUMBER"><xsl:value-of select="StaffNo"/></xsl:attribute>
																					<xsl:attribute name="PERCENTSHARESHELD"></xsl:attribute>
																					<xsl:attribute name="PROBATIONARYINDICATOR"></xsl:attribute>
																					<xsl:attribute name="P60SEENINDICATOR"></xsl:attribute>
																					<xsl:attribute name="WAGESLIPSSEENINDICATOR"></xsl:attribute>
																					<xsl:attribute name="EMPLOYMENTRELATIONSHIPIND">1</xsl:attribute> <!-- They are related to the employer -->
																					
																					<!-- DO INCOMES, BASIC, REGULAR, BONUS -->
																					
																					<!-- Have Full income details or only TOTAL income details been captured? 
																					Full income details are captured unless its a Self Certification, SubPrime - Self Cert, 
																					or Buy to Let Mortgage (and if its BTL the applicant is not a first time buyer) -->
																					<xsl:choose>
																						
																						<!-- Only Total Income -->
																						<xsl:when test="/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 3 or /REQUEST/ProposalRoot/Proposal/MortgageTypeId = 8 or (/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 2 and ../../FirstTimeBuyerYN = 0)">
																						
																							<!-- EARNEDINCOME -->
																							<xsl:element name="EARNEDINCOME">
																								<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="TotalAnnualIncome"/></xsl:attribute>
																								<xsl:attribute name="EARNEDINCOMETYPE">1</xsl:attribute> <!-- 1 = Basic IncomeType -->
																								<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																							</xsl:element> <!-- EARNEDINCOME -->
																							
																						</xsl:when> <!-- Only Total Income -->
																						
																						<xsl:otherwise>
																							
																							<!-- Is there a basic income -->
																							<xsl:if test="number(BasicAnnualIncome) > 0">
																								<!-- EARNEDINCOME -->
																								<xsl:element name="EARNEDINCOME">
																									<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:call-template name="CalculateBasicIncome"><xsl:with-param name="BasicAnnualIncome" select="BasicAnnualIncome" /><xsl:with-param name="OtherIncome" select="OtherIncome" /><xsl:with-param name="OtherIncomeTypeId" select="OtherIncomeTypeId" /></xsl:call-template></xsl:attribute>
																									<xsl:attribute name="EARNEDINCOMETYPE">1</xsl:attribute> <!-- 1 = Basic IncomeType -->
																									<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																								</xsl:element> <!-- EARNEDINCOME -->
																	
																							</xsl:if> <!-- Basic income -->
																	
																							<!-- Is there a regular overtime income -->
																							<xsl:if test="number(RegularOvertimeAmount) > 0">
																								<!-- EARNEDINCOME -->
																								<xsl:element name="EARNEDINCOME">
																									<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="RegularOvertimeAmount"/></xsl:attribute>
																									<xsl:attribute name="EARNEDINCOMETYPE">3</xsl:attribute> <!-- 3 = Regular overtime IncomeType -->
																									<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																								</xsl:element> <!-- EARNEDINCOME -->
																	
																							</xsl:if> <!-- Regular overtime income -->
																	
																							<!-- Is there a regular bonus or commission income -->
																							<xsl:if test="number(RegularBonusCommissionAmount) > 0">
																								<!-- EARNEDINCOME -->
																								<xsl:element name="EARNEDINCOME">
																									<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="RegularBonusCommissionAmount"/></xsl:attribute>
																									<xsl:attribute name="EARNEDINCOMETYPE">5</xsl:attribute> <!-- 5 = Regular bonus IncomeType -->
																									<xsl:attribute name="PAYMENTFREQUENCYTYPE">1</xsl:attribute> <!-- payment frequency is per annum -->
																								</xsl:element> <!-- EARNEDINCOME -->
																	
																							</xsl:if> <!-- Regular overtime income -->
																							
																						</xsl:otherwise> <!-- Full Income Details -->
																						
																					</xsl:choose>
																					
																				</xsl:element> <!-- EMPLOYEDDETAILS -->
																			
																			</xsl:if> <!-- Applicant Employment is Employed -->
																			
																			<!-- Applicant Employment is self employed -->
																			<xsl:if test="EmploymentTypeId='3'">
																				
																				<!-- SELFEMPLOYEDDETAILS -->
																				<xsl:element name="SELFEMPLOYEDDETAILS">
																					<xsl:attribute name="DATEFINANCIALINTERESTHELD"></xsl:attribute>
																					<xsl:attribute name="OTHERBUSINESSCONNECTIONS"></xsl:attribute>
																					<xsl:attribute name="PERCENTSHARESHELD"></xsl:attribute>
																					<xsl:attribute name="REGISTRATIONNUMBER"><xsl:value-of select="CompanyRegNo"/></xsl:attribute>
																					<xsl:attribute name="VATNUMBER"><xsl:value-of select="VATNo"/></xsl:attribute>
																					<xsl:attribute name="NATUREOFBUSINESS"><xsl:value-of select="NatureOfBusiness"/></xsl:attribute>
																					
																					<!-- Have Full income details or only TOTAL income details been captured? 
																					Full income details are captured unless its a Self Certification, SubPrime - Self Cert, 
																					or Buy to Let Mortgage (and if its BTL the applicant is not a first time buyer) -->
																					<xsl:choose>
																						
																						<!-- Only Total Income -->
																						<xsl:when test="/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 3 or /REQUEST/ProposalRoot/Proposal/MortgageTypeId = 8 or (/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 2 and ../../FirstTimeBuyerYN = 0)">
																						
																							<!-- NETPROFIT -->
																							<xsl:element name="NETPROFIT">
																								<xsl:attribute name="EVIDENCEOFACCOUNTS"></xsl:attribute>
																								<xsl:attribute name="YEAR1"><xsl:call-template name="OutputCurrentYear"><xsl:with-param name="NOW" select="$Now" /></xsl:call-template></xsl:attribute>
																								<xsl:attribute name="YEAR1AMOUNT"><xsl:value-of select="TotalAnnualAmount"/></xsl:attribute>
																							</xsl:element> <!-- NETPROFIT -->
																						
																						</xsl:when> <!-- Only Total Income -->
																						
																						<!-- Full Income details -->
																						<xsl:otherwise>
																							
																							<!-- NETPROFIT -->
																							<xsl:element name="NETPROFIT">
																								<xsl:attribute name="EVIDENCEOFACCOUNTS"></xsl:attribute>
																								<xsl:attribute name="YEAR1"><xsl:value-of select="Year1Year"/></xsl:attribute>
																								<xsl:attribute name="YEAR1AMOUNT"><xsl:value-of select="NetProfitsYear1"/></xsl:attribute>
																								<xsl:attribute name="YEAR2"></xsl:attribute>
																								<xsl:attribute name="YEAR2AMOUNT"></xsl:attribute>
																								<xsl:attribute name="YEAR3"></xsl:attribute>
																								<xsl:attribute name="YEAR3AMOUNT"></xsl:attribute>
																							</xsl:element> <!-- NETPROFIT -->
																						
																						</xsl:otherwise> <!-- Full Income details -->
																						
																					</xsl:choose>
																					
																				</xsl:element> <!-- SELFEMPLOYEDDETAILS -->
																				
																			</xsl:if> <!-- Applicant Employment is self employed -->
																			
																		</xsl:element> <!-- EMPLOYMENT -->
																		
																	</xsl:element> <!-- EMPLOYERTHIRDPARTY -->
																	
																</xsl:element> <!-- EMPLOYERADDRESS -->
																
															</xsl:element> <!-- EMPLOYMENTDETAILS -->
														
														</xsl:element> <!-- ACCOUNTANT -->
														
													</xsl:element> <!-- ACCOUNTANTTHIRDPARTY -->
													
												</xsl:element> <!-- ACCOUNTANTADDRESS -->
												
											</xsl:element> <!-- EMPLOYMENTDETAILSWITHACCOUNTANT -->
											
										</xsl:if> <!-- Has Accountant details -->
										
										<!-- Is the applicant unemployed -->
										<xsl:if test="EmploymentTypeId='4'">
										
											<!-- EMPLOYMENT -->
											<xsl:element name="EMPLOYMENT">
												<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="EmploymentTypeId"/></xsl:attribute>
												<xsl:attribute name="DATELEFTORCEASEDTRADING"><xsl:value-of select="DateStarted"/></xsl:attribute>
												<xsl:attribute name="MAINSTATUS">
													<xsl:choose>
														<!-- If its the first employment node, then its the main status -->
														<xsl:when test="position() = 1"><xsl:text>1</xsl:text></xsl:when>
														<xsl:otherwise><xsl:text>0</xsl:text></xsl:otherwise>
													</xsl:choose>
												</xsl:attribute>
												
											</xsl:element> <!-- EMPLOYMENT -->
											
										</xsl:if> <!-- Is the applicant unemployed -->
										
									</xsl:for-each> <!-- EmploymentRoot/Employment -->
									
									<!-- DO OTHER INCOMES -->
									
									<!-- ApplicantOtherIncomeRoot/ApplicantOtherIncome, make sure there is actually an amount -->
									<xsl:for-each select="ApplicantOtherIncomeRoot/ApplicantOtherIncome[number(Amount) > 0]">
										
										<!-- UNEARNEDINCOME -->
										<xsl:element name="UNEARNEDINCOME">
											<xsl:attribute name="UNEARNEDINCOMEAMOUNT"><xsl:value-of select="Amount"/></xsl:attribute>
											<xsl:attribute name="UNEARNEDINCOMETYPE"><xsl:value-of select="OtherIncomeTypeId"/></xsl:attribute>
											<xsl:attribute name="PAYMENTFREQUENCY">1</xsl:attribute> <!-- per annum -->
											<xsl:attribute name="OTHERINCOMEDETAILS"></xsl:attribute>
										</xsl:element> <!-- UNEARNEDINCOME -->
										
									</xsl:for-each> <!-- ApplicantOtherIncomeRoot/ApplicantOtherIncome -->
									
									<!-- Applicant employment other incomes! -->
									<xsl:for-each select="ApplicantEmploymentRoot/ApplicantEmployment[number(OtherIncome) > 0]"> <!-- If they have got an OtherIncome against the current employment-->
										<!-- UNEARNEDINCOME -->
										<xsl:element name="UNEARNEDINCOME">
											<xsl:attribute name="UNEARNEDINCOMEAMOUNT"><xsl:value-of select="OtherIncome"/></xsl:attribute>
											<xsl:attribute name="UNEARNEDINCOMETYPE"><xsl:value-of select="OtherIncomeTypeId"/></xsl:attribute>
											<xsl:attribute name="PAYMENTFREQUENCY">1</xsl:attribute>  <!-- per annum -->
											<xsl:attribute name="OTHERINCOMEDETAILS"></xsl:attribute>
										</xsl:element> <!-- UNEARNEDINCOME -->
									</xsl:for-each> <!-- Applicant employment other incomes! -->
									
									<!-- Risks -->
									<!-- Risk section is only captured for the first applicant -->
									<xsl:if test="position() = 1">
																			
										<!-- Applicants Commitments -->
										<xsl:for-each select="ApplicantCommitmentsRoot/ApplicantCommitments">
									
											<!-- LOANSLIABILITIESTHIRDPARTY -->
											<xsl:element name="LOANSLIABILITIESTHIRDPARTY">
												<xsl:attribute name="COMPANYNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="CompanyName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
												<xsl:attribute name="THIRDPARTYTYPE">3</xsl:attribute> <!-- Other Lender Thirdparty type -->
												
												<xsl:element name="ACCOUNT">
													<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="AccountNo"/></xsl:attribute>
												
													<xsl:element name="LOANSLIABILITIES">
														<xsl:attribute name="ENDDATE"><xsl:value-of select="DateToBeRepaid"/></xsl:attribute>
														<xsl:attribute name="LOANREPAYMENTINDICATOR"><xsl:value-of select="ClearedOnCompletionYN"/></xsl:attribute>
														<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MonthlyPayment"/></xsl:attribute>
														<xsl:attribute name="TOTALOUTSTANDINGBALANCE"><xsl:value-of select="OutstandingAmount"/></xsl:attribute>
														<xsl:attribute name="AGREEMENTTYPE"><xsl:value-of select="AgreementTypeId"/></xsl:attribute>
														<xsl:attribute name="ADDITIONALINDICATOR">0</xsl:attribute> <!-- There are no additional holders details -->
														
													</xsl:element> <!-- LOANSLIABILITIES -->
													
													<xsl:element name="ACCOUNTRELATIONSHIP">
														<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[0]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
														<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[0]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
														<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant role type -->
														<xsl:attribute name="ACCOUNTGUID">../@ACCOUNTGUID</xsl:attribute>
														
													</xsl:element> <!-- ACCOUNTRELATIONSHIP -->
													
												</xsl:element> <!-- ACCOUNT -->
												
											</xsl:element> <!-- LOANSLIABILITIESTHIRDPARTY -->
											
										</xsl:for-each> <!-- ApplicantCommitmentsRoot/ApplicantCommitments apart from credit cards-->
										
										<!-- Credit and bank cards are now put in loans and liabilities 
										- Applicatant Commitments, Credit and Bank cards -
										<xsl:for-each select="ApplicantCommitmentsRoot/ApplicantCommitments[AgreementTypeId = 2]">
										
											<xsl:element name="BANKCREDITCARD">
												<xsl:attribute name="AVERAGEMONTHLYREPAYMENT"><xsl:value-of select="MonthlyPayment"/></xsl:attribute>
												<xsl:attribute name="CARDPROVIDER"><xsl:value-of select="CompanyName"/></xsl:attribute>
												<xsl:attribute name="TOTALOUTSTANDINGBALANCE"><xsl:value-of select="OutstandingAmount"/></xsl:attribute>
												<xsl:attribute name="ADDITIONALINDICATOR">0</xsl:attribute> - There are no additional card holders -
											</xsl:element>
										
										</xsl:for-each> - Applicatant Commitments, Credit and Bank cards -
										-->
										
										<!-- Do CCJ's -->
										<!-- Is there any CCJ's -->
										<xsl:if test="AdverseHistoryYN = 1 and CCJSYN = 1">
										
											<!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=2] All the CCJ's from the applicant risks -->
											<xsl:for-each select="ApplicantRisksRoot/ApplicantRisks[RiskTypeId = 2]">
												
												<!-- CCJHISTORY -->	
												<xsl:element name="CCJHISTORY"> 
													<xsl:attribute name="DATEOFJUDGEMENT"><xsl:value-of select="StartDate"/></xsl:attribute>
													<xsl:attribute name="VALUEOFJUDGEMENT"><xsl:value-of select="Amount"/></xsl:attribute>
													<xsl:attribute name="DATECLEARED"><xsl:value-of select="EndDate"/></xsl:attribute>
													<xsl:attribute name="DEFAULTRECORD">
														<xsl:choose>
															<xsl:when test="CCJDefaultTypeId = 2">
																<xsl:text>1</xsl:text>
															</xsl:when>
															<xsl:otherwise>
																<xsl:text>0</xsl:text>
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
													<xsl:attribute name="CCJTYPE">
														<xsl:choose>
															<xsl:when test="string-length(EndDate) > 0">
																<xsl:text>1</xsl:text> <!-- Satisfied -->
															</xsl:when>
															<xsl:otherwise>
																<xsl:text>2</xsl:text> <!-- Unsatisfied -->
															</xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
													
													<!-- CUSTOMERVERSIONCCJHISTORY -->
													<xsl:element name="CUSTOMERVERSIONCCJHISTORY" />
													
												</xsl:element> <!-- CCJHISTORY -->
												
											</xsl:for-each> <!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=2] -->
										
										</xsl:if> <!-- Is there any CCJ's -->
																			
										<!-- Do Arrears -->
										<!-- Is there any Arrear's -->
										<xsl:if test="AdverseHistoryYN = 1 and ArrearsYN = 1">
																			
											<!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=3] All the Arrears from the applicant risks -->
											<xsl:for-each select="ApplicantRisksRoot/ApplicantRisks[RiskTypeId = 3]">
												
												<!-- ARREARSHISTORYACCOUNT -->
												<xsl:element name="ARREARSHISTORYACCOUNT">
													
													<!-- ARREARSHISTORY -->
													<xsl:element name="ARREARSHISTORY">
														<xsl:attribute name="DATECLEARED"><xsl:value-of select="EndDate"/></xsl:attribute>
														<xsl:attribute name="DESCRIPTIONOFLOAN"><xsl:value-of select="ArrearsTypeId"/></xsl:attribute>
														<xsl:attribute name="MAXIMUMBALANCE"><xsl:value-of select="HighestArrearsAmount"/></xsl:attribute>
														<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="Amount"/></xsl:attribute>
														<xsl:attribute name="OTHERDETAILS"></xsl:attribute>
														<xsl:attribute name="MAXIMUMNUMBEROFMONTHS"><xsl:value-of select="HighestArrearsMonths"/></xsl:attribute>
														<xsl:attribute name="ADDITIONALDETAILS"></xsl:attribute>
														<xsl:attribute name="ADDITIONALINDICATOR">0</xsl:attribute>
														<xsl:attribute name="CURRENTYEARSINARREARS"><xsl:value-of select="PaymentsMissed"/></xsl:attribute>
														
													</xsl:element> <!-- ARREARSHISTORY -->
													
													<xsl:element name="ACCOUNTRELATIONSHIP">
														<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[0]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
														<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[0]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
														<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant role type -->
														<xsl:attribute name="ACCOUNTGUID">../@ACCOUNTGUID</xsl:attribute>
														
													</xsl:element> <!-- ACCOUNTRELATIONSHIP -->
													
													<!-- If this is an 'Other' arrears history create a record in the OTHERARREARSHISTORY table to store the descrioption in -->
													<xsl:if test="ArrearsTypeId = 6">
													
														<!-- OTHERARREARSACCOUNT -->
														<xsl:element name="OTHERARREARSACCOUNT"/>
														
													</xsl:if> <!-- An 'other' arrears history -->
													
												</xsl:element> <!-- ARREARSHISTORYACCOUNT -->
												
											</xsl:for-each> <!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=3] -->
											
										</xsl:if> <!-- Is there any Arrear's -->
										
										<!-- Is there a bankruptsy detail -->
										<xsl:if test="AdverseHistoryYN = 1 and BankruptYN = 1">
											
											<!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=7] All the Bankrupsy from the applicant risks -->
											<xsl:for-each select="ApplicantRisksRoot/ApplicantRisks[RiskTypeId=7]">	
												
												<!-- BANKRUPTCYHISTORY -->
												<xsl:element name="BANKRUPTCYHISTORY">
													<xsl:attribute name="DATEOFDISCHARGE"><xsl:value-of select="EndDate"/></xsl:attribute>
													<xsl:attribute name="IVA">0</xsl:attribute> <!-- No - This is not an IVA -->
													
													<!-- CUSTOMERVERSIONBANKRUPTCYHISTORY -->
													<xsl:element name="CUSTOMERVERSIONBANKRUPTCYHISTORY" />
													
												</xsl:element> <!-- BANKRUPTCYHISTORY -->
												
											</xsl:for-each> <!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=7] -->
											
										</xsl:if> <!-- Is there a bankruptsy detail -->
										
										<!-- Is there an IVA detail -->
										<xsl:if test="AdverseHistoryYN = 1 and IVASYN = 1">
											
											<!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=10] All the IVA's from the applicant risks -->
											<xsl:for-each select="ApplicantRisksRoot/ApplicantRisks[RiskTypeId=10]">	
												
												<!-- BANKRUPTCYHISTORY -->
												<xsl:element name="BANKRUPTCYHISTORY">
													<xsl:attribute name="DATEDECLARED"><xsl:value-of select="StartDate"/></xsl:attribute>
													<xsl:attribute name="DATEOFDISCHARGE"><xsl:value-of select="EndDate"/></xsl:attribute>
													<xsl:attribute name="IVA">1</xsl:attribute> <!-- Yes - This is an IVA -->
													
													<!-- CUSTOMERVERSIONBANKRUPTCYHISTORY -->
													<xsl:element name="CUSTOMERVERSIONBANKRUPTCYHISTORY" />
													
												</xsl:element> <!-- BANKRUPTCYHISTORY -->
												
											</xsl:for-each> <!-- ApplicantRisksRoot/ApplicantRisks[RiskTypeId=10] -->
											
										</xsl:if> <!-- Is there a bankruptsy detail -->
										
									</xsl:if> <!-- Only for the first applicant -->
									
									<!-- IDENTIFICATION -->
									
									<!-- Primary identification type -->
									<xsl:for-each select="ApplicantPrimaryIdentificationRoot/ApplicantPrimaryIdentification">
										<xsl:element name="VERIFICATION">
											<xsl:attribute name="VERIFICATIONTYPE">1</xsl:attribute> <!-- Personal ID Verification Type -->
											<xsl:attribute name="IDENTIFICATIONTYPE"><xsl:value-of select="PrimaryIdentificationTypeId"/></xsl:attribute>
											<xsl:attribute name="REFERENCE"><xsl:value-of select="DocumentDetails"/></xsl:attribute>
											
										</xsl:element> <!-- VERIFICATION -->
										
									</xsl:for-each> <!-- Primary identification type -->
									
									<!-- Secondary identification type -->
									<xsl:for-each select="ApplicantSecondaryIdentificationRoot/ApplicantSecondaryIdentification">
										<xsl:element name="VERIFICATION">
											<xsl:attribute name="VERIFICATIONTYPE">2</xsl:attribute> <!-- Residency ID Verification Type -->
											<xsl:attribute name="IDENTIFICATIONTYPE"><xsl:value-of select="SecondaryIdentificationTypeId"/></xsl:attribute>
											<xsl:attribute name="REFERENCE"><xsl:value-of select="DocumentDetails"/></xsl:attribute>
											
										</xsl:element>  <!-- VERIFICATION -->
										
									</xsl:for-each> <!-- Secondary identification type -->
									
									<!-- Is there a previous name -->
									<xsl:if test="string-length(PreviousName) > 0">
										<!-- ALIASPERSON -->
										<xsl:element name="ALIASPERSON">
											<xsl:attribute name="TITLE"><xsl:value-of select="TitleTypeId"/></xsl:attribute>
											<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="Forename1"/></xsl:attribute>
											<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="Forename2"/></xsl:attribute>
											<xsl:attribute name="SURNAME"><xsl:value-of select="PreviousName"/></xsl:attribute>
											<!-- ALIAS -->
											<xsl:element name="ALIAS">
												<xsl:attribute name="ALIASTYPE">10</xsl:attribute> <!-- Alias -->
												
											</xsl:element> <!-- ALIAS -->
											
										</xsl:element> <!-- ALIASPERSON -->
										
									</xsl:if> <!-- There is a previous name -->
									
								</xsl:element> <!-- CUSTOMERVERSION -->
								
							</xsl:element> <!-- CUSTOMER -->
							
						</xsl:for-each> <!-- ApplicantRoot/Applicant -->
						
						<!-- COMPLICATED ACCOUNT RELATIONSHIPS FOR CURRENT ADDRESS, ITS DONE HERE AFTER ALL APPLICANTS HAVE BEEN CREATED -->
						<xsl:for-each select="ApplicantRoot/Applicant">
							
							<!-- Note the applicant number -->
							<xsl:variable name="ApplicantNumber" select="position()"/>
							
							<!-- Do they have a current address with a mortgage against it? -->
							<xsl:for-each select="ApplicantAddressRoot/ApplicantAddress[ResidentialStatusTypeId = 2 and HaveMortgageYN = 1 and AddressTypeId = 2 and MortgageLenderSameAsApplicant = 99]"> <!-- Its owner occupied, it has a mortgage, its the current address, and the sameasapplicant is different -->
								
								<!-- Get the Mortgage second charge indicator, becuase if this is set I will need to create account relationship for this account as well -->
								<xsl:variable name="SecondChargeInd" select="MortgageSecondCharge"/>
								
								<!-- Create the accout relationship for the main applicant -->
								<!-- ACCOUNT RELATIONSHIP -->
								<xsl:element name="ACCOUNTRELATIONSHIP">
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
									<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
									<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
									<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE=1]/LENDERLANDLORDDETAILS/LENDERLANDLORDADDRESS/LENDERLANDLORDTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																					
								</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
								
								<!-- Is there a second charge on this account -->
								<xsl:if test="$SecondChargeInd = 1">
								
									<!-- ACCOUNT RELATIONSHIP -->
									<xsl:element name="ACCOUNTRELATIONSHIP">
										<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
										<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
										<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
										<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE=1]/SECONDCHARGELENDERDETAILS/LENDERADDRESS/LENDERTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																						
									</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
									
								</xsl:if> <!-- Is there a second charge on this account -->
								
								<!-- Is there any previous mortgage in the last 12 months -->
								<xsl:if test="PrevMortgageYN = 1">
									<xsl:for-each select="PreviousLenderRoot/PreviousLender">
									
										<!-- ACCOUNT RELATIONSHIP -->
										<xsl:element name="ACCOUNTRELATIONSHIP">
											<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
											<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
											<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
											<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE=1]/PREVIOUSLENDERDETAILS[</xsl:text><xsl:value-of select="position() - 1"/><xsl:text>]/LENDERADDRESS/LENDERTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																							
										</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
									
									</xsl:for-each>
								</xsl:if>
								
								<!-- Loop through the applicants -->
								<xsl:for-each select="/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant">
																					
									<xsl:variable name="MortgageAccountApplicantCount" select="position()"/>

									<!-- Dont do the current applicant again -->
									<xsl:if test="$MortgageAccountApplicantCount != $ApplicantNumber">
																						
										<!-- Find any other applicants current address which are attached to this mortgage account -->
										<xsl:for-each select="ApplicantAddressRoot/ApplicantAddress[AddressTypeId = 2 and ResidentialStatusTypeId = 2 and HaveMortgageYN = 1 and MortgageLenderSameAsApplicant = $ApplicantNumber]">
																							
											<!-- ACCOUNTRELATIONSHIP -->
											<xsl:element name="ACCOUNTRELATIONSHIP">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
												<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
												<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE = 1]/LENDERLANDLORDDETAILS/LENDERLANDLORDADDRESS/LENDERLANDLORDTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																								
											</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
											
											<!-- Is there a second charge on this account -->
											<xsl:if test="$SecondChargeInd = 1">
											
												<!-- ACCOUNTRELATIONSHIP -->
												<xsl:element name="ACCOUNTRELATIONSHIP">
													<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
													<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
													<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
													<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE = 1]/SECONDCHARGELENDERDETAILS/LENDERADDRESS/LENDERTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																									
												</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
											
											</xsl:if> <!-- Is there a second charge on this account -->
																
											<!-- Is there any previous mortgage in the last 12 months -->
											<xsl:if test="PrevMortgageYN = 1">
												<xsl:for-each select="PreviousLenderRoot/PreviousLender">
												
													<!-- ACCOUNTRELATIONSHIP -->
													<xsl:element name="ACCOUNTRELATIONSHIP">
														<xsl:attribute name="CUSTOMERNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/@CUSTOMERNUMBER</xsl:text></xsl:attribute> <!-- Keys gained from an xpath -->
														<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$MortgageAccountApplicantCount - 1"/><xsl:text>]/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER</xsl:text></xsl:attribute>
														<xsl:attribute name="CUSTOMERROLE">1</xsl:attribute> <!-- Applicant Role type -->
														<xsl:attribute name="ACCOUNTGUID"><xsl:text>//CUSTOMER[</xsl:text><xsl:value-of select="$ApplicantNumber - 1"/><xsl:text>]/CUSTOMERVERSION/ADDRESS/CUSTOMERADDRESS[@ADDRESSTYPE = 1]/PREVIOUSLENDERDETAILS[</xsl:text><xsl:value-of select="position() - 1"/><xsl:text>]/LENDERADDRESS/LENDERTHIRDPARTY/ACCOUNT/@ACCOUNTGUID</xsl:text></xsl:attribute>
																										
													</xsl:element> <!-- ACCOUNT RELATIONSHIP -->
												
												</xsl:for-each>
											</xsl:if>
																													
										</xsl:for-each> <!-- Find any other applicants current address which are attached to this mortgage account -->
																						
									</xsl:if> <!-- Dont do the current applicant again -->
																					
								</xsl:for-each> <!-- Find any other applicants which are attached to this mortgage account -->
								
							</xsl:for-each> <!-- Do they have a current address with a mortgage against it? -->
							
						</xsl:for-each> <!-- ApplicantRoot/Applicant -->
						
						<!-- OTHERRESIDENT -->
						<xsl:call-template name="OutputOtherResidents"><xsl:with-param name="NOW" select="$Now" /></xsl:call-template>
						
						<!-- NEWLOAN -->
						<xsl:element name="NEWLOAN">
							<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="LoanAmountRequired"/></xsl:attribute>
							<xsl:attribute name="MIRASAMOUNT"></xsl:attribute>
							<xsl:attribute name="MIRASELIGIBILITY"></xsl:attribute>
							<xsl:attribute name="PURPOSEOFLOAN"></xsl:attribute>
							<xsl:attribute name="REPAYMENTTYPE"><xsl:value-of select="RepaymentTypeId"/></xsl:attribute>
							<xsl:attribute name="RATECHANGENOTIFICATION"></xsl:attribute>
							<xsl:attribute name="SAMEMORTGAGETYPE"></xsl:attribute>
							<xsl:attribute name="SAMEFINISHINDICATOR"></xsl:attribute>
							<xsl:attribute name="SOLICITORINDICATOR"></xsl:attribute>
							<xsl:attribute name="TERMMONTHS"><xsl:value-of select="TermMonths"/></xsl:attribute>
							<xsl:attribute name="TERMYEARS"><xsl:value-of select="Term"/></xsl:attribute>
							<xsl:attribute name="INTERESTONLYAMOUNT"></xsl:attribute>
							<xsl:attribute name="OTHERLOANPURPOSE"></xsl:attribute>
							<xsl:attribute name="ADDITIONALBORROWINGAMOUNT"></xsl:attribute>  <!-- TODO: FINDOUT WHAT THIS IS -->
						</xsl:element> <!-- NEWLOAN -->
						
						<!-- PropertyRoot/Property details -->
						<xsl:for-each select="PropertyRoot/Property">
						
							<!-- NEWPROPERTY -->
							<xsl:element name="NEWPROPERTY">
								<xsl:attribute name="ANYOTHERRESIDENTSINDICATOR"></xsl:attribute>
								<xsl:attribute name="APPLICANTTOOCCUPYINDICATOR"><xsl:value-of select="ResidentialYN"/></xsl:attribute>
								<xsl:attribute name="BUILDINGCONSTRUCTIONTYPE"><xsl:value-of select="WallConstructionTypeId"/></xsl:attribute>
								<xsl:attribute name="DESCRIPTIONOFPROPERTY"><xsl:value-of select="PropertyConstructionTypeId"/></xsl:attribute>
								<xsl:attribute name="FAMILYSALEINDICATOR">0</xsl:attribute> <!-- This is family sale/right to buy which BM dont do -->
								<xsl:attribute name="FLOORNUMBER"><xsl:value-of select="FlatFloorNo"/></xsl:attribute>
								<xsl:attribute name="GRANTAMOUNT"></xsl:attribute>
								<xsl:attribute name="LETORPARTLETINDICATOR">
									<xsl:if test="/REQUEST/ProposalRoot/Proposal/MortgageTypeId = 2"> <!-- Buy to Let mortgage -->
										<xsl:text>1</xsl:text>
									</xsl:if>
								</xsl:attribute>
								<xsl:attribute name="NUMBEROFSTOREYS"><xsl:value-of select="NoStoreys"/></xsl:attribute>
								<xsl:attribute name="NUMBEROFGARAGES"><xsl:value-of select="NoGarages"/></xsl:attribute>
								<xsl:attribute name="PROPERTYLOCATION"><xsl:value-of select="PropertyLocationTypeId"/></xsl:attribute>
								<xsl:attribute name="ROOFCONSTRUCTIONTYPE"><xsl:value-of select="RoofConstructionTypeId"/></xsl:attribute>
								<xsl:attribute name="TENURETYPE"><xsl:value-of select="TenureTypeId"/></xsl:attribute>
								<xsl:attribute name="TYPEOFPROPERTY"><xsl:value-of select="PropertyConstructionTypeId"/></xsl:attribute>
								<xsl:attribute name="VALUATIONTYPE"><xsl:value-of select="ValuationTypeId"/></xsl:attribute>
								<xsl:attribute name="YEARBUILT"><xsl:value-of select="ConstructionYear"/></xsl:attribute>
								<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="AnticipatedRentalIncome"/></xsl:attribute>
								<xsl:attribute name="RENTALINCOMESTATUS">10</xsl:attribute> <!-- Projected Monthly Rental Income -->
							</xsl:element> <!-- NEWPROPERTY -->
						
							<!-- NEWPROPERTYACTUALADDRESS -->
							<xsl:element name="NEWPROPERTYACTUALADDRESS">
								<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NAME'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="PropertyBuildingNo" /></xsl:call-template></xsl:attribute>
								<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:call-template name="OutputBuildingNameOrNumber"><xsl:with-param name="TYPE" select="'NUMBER'" /><xsl:with-param name="BUILDINGNAMENUMBER" select="PropertyBuildingNo" /></xsl:call-template></xsl:attribute>
								<xsl:attribute name="STREET"><xsl:value-of select="PropertyBuildingName"/></xsl:attribute>
								<xsl:attribute name="DISTRICT"><xsl:value-of select="PropertyStreet"/></xsl:attribute>
								<xsl:attribute name="TOWN"><xsl:value-of select="PropertyTown"/></xsl:attribute>
								<xsl:attribute name="COUNTY"><xsl:value-of select="PropertyCounty"/></xsl:attribute>
								<xsl:attribute name="COUNTRY">2</xsl:attribute> <!-- Unknown -->
								<xsl:attribute name="POSTCODE"><xsl:value-of select="PropertyPostCode"/></xsl:attribute>

								<!-- NEWPROPERTYADDRESS -->
								<!-- Need to get the valuercontactdetails out -->
								<xsl:element name="NEWPROPERTYADDRESS">
									<xsl:attribute name="ARRANGEMENTSFORACCESS"><xsl:value-of select="/REQUEST/ProposalRoot/Proposal/ContactDetailsForValuation/ValuationContactTypeId"/></xsl:attribute>
									<!-- Valuation contact address details -->
									<xsl:attribute name="OTHERARRANGEMENTSFORACCESS">
										<xsl:for-each select="/REQUEST/ProposalRoot/Proposal/ContactDetailsForValuation/Address">
											<xsl:text>Valuation Contact Address :&#13;&#10;</xsl:text>
											<xsl:value-of select="BuildingNo"/><xsl:text>, </xsl:text><xsl:value-of select="BuildingName"/><xsl:text>&#13;&#10;</xsl:text>
											<xsl:value-of select="Street"/><xsl:text>&#13;&#10;</xsl:text>
											<xsl:value-of select="District"/><xsl:text>&#13;&#10;</xsl:text>
											<xsl:value-of select="Town"/><xsl:text>&#13;&#10;</xsl:text>
											<xsl:value-of select="County"/><xsl:text>&#13;&#10;</xsl:text>
											<xsl:value-of select="PostCode"/><xsl:text>&#13;&#10;&#13;&#10;</xsl:text>
										</xsl:for-each>
									</xsl:attribute> <!-- OTHERARRANGEMENTSFORACCESS -->
									<xsl:attribute name="ACCESSTELEPHONENUMBER"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'TELEPHONENUMBER'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="/REQUEST/ProposalRoot/Proposal/ContactDetailsForValuation/ValuerTelephoneNo" /></xsl:call-template></xsl:attribute>
									<xsl:attribute name="ACCESSCONTACTNAME"><xsl:value-of select="/REQUEST/ProposalRoot/Proposal/ContactDetailsForValuation/ValuerName"/></xsl:attribute>
									<xsl:attribute name="COUNTRYCODE"></xsl:attribute>
									<xsl:attribute name="AREACODE"><xsl:call-template name="FindTelephoneNoPortion"><xsl:with-param name="PORTION" select="'AREACODE'" /><xsl:with-param name="FULLTELEPHONENUMBER" select="/REQUEST/ProposalRoot/Proposal/ContactDetailsForValuation/ValuerTelephoneNo" /></xsl:call-template></xsl:attribute>
									<xsl:attribute name="EXTENSIONNUMBER"></xsl:attribute>
								</xsl:element> <!-- NEWPROPERTYADDRESS -->

							</xsl:element> <!-- NEWPROPERTYACTUALADDRESS -->
						
							<!-- Do the number of the bedrooms -->
							<!-- NEWPROPERTYROOMTYPE -->
							<xsl:element name="NEWPROPERTYROOMTYPE"> 
								<xsl:attribute name="ROOMTYPE">2</xsl:attribute> <!-- Bedroom - New property room type, Bedroom -->
								<xsl:attribute name="NUMBEROFROOMS"><xsl:value-of select="NoBedrooms"/></xsl:attribute> 
							</xsl:element> <!-- NEWPROPERTYROOMTYPE -->
							
							<!-- Leasehold details -->
							<xsl:if test="TenureTypeId = 3"> <!-- Leasehold -->
							
								<xsl:element name="NEWPROPERTYLEASEHOLD"> 
									<xsl:attribute name="GROUNDRENT"><xsl:value-of select="GroundRentAmount"/></xsl:attribute>
									<xsl:attribute name="UNEXPIREDTERMOFLEASEYEARS"><xsl:value-of select="LeaseRemaining"/></xsl:attribute>  
									<xsl:attribute name="SERVICECHARGE"><xsl:value-of select="ServiceChargeAmount"/></xsl:attribute>  
									
								</xsl:element> <!-- NEWPROPERTYLEASEHOLD -->
								
							</xsl:if> <!-- Is this a leasehold property -->
							
						</xsl:for-each> <!-- PropertyRoot/Property -->
						
						<!-- Source of deposit -->
						<!-- NEWPROPERTYDEPOSIT -->
						<xsl:element name="NEWPROPERTYDEPOSIT"> 
							<xsl:attribute name="SOURCEOFFUNDING"><xsl:value-of select="DepositSourceTypeId"/></xsl:attribute> 
							<xsl:attribute name="AMOUNT"><xsl:value-of select="PropertyRoot/Property/PurchasePrice - LoanAmountRequired"/></xsl:attribute> 
						</xsl:element> <!-- NEWPROPERTYDEPOSIT -->
						
						<!-- Shared ownership, BM dont do this so set the indicator to 0 -->
						<!-- SHAREDOWNERSHIPDETAILS -->
						<xsl:element name="SHAREDOWNERSHIPDETAILS">
							<xsl:attribute name="SHAREDOWNERSHIPINDICATOR">0</xsl:attribute>
						</xsl:element> <!-- SHAREDOWNERSHIPDETAILS -->
						
						<!-- QUOTATION -->
						<xsl:element name="QUOTATION">
							<xsl:attribute name="DATEANDTIMEGENERATED"><xsl:value-of select="$Now"/></xsl:attribute>
							<xsl:attribute name="LIFESUBQUOTEREQUIRED">0</xsl:attribute>
							<xsl:attribute name="BCSUBQUOTEREQUIRED">0</xsl:attribute>
							<xsl:attribute name="PPSUBQUOTEREQUIRED">0</xsl:attribute>
							
							<!-- MORTGAGESUBQUOTE -->
							<xsl:element name="MORTGAGESUBQUOTE">
								<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="LoanAmountRequired"/></xsl:attribute>
								<xsl:attribute name="DATEANDTIMEGENERATED"><xsl:value-of select="$Now"/></xsl:attribute>
								<xsl:attribute name="DEPOSIT"><xsl:value-of select="PropertyRoot/Property/PurchasePrice - LoanAmountRequired"/></xsl:attribute>
								<xsl:attribute name="LTV"><xsl:value-of select="(number(LoanAmountRequired) div number(PropertyRoot/Property/PurchasePrice)) * 100"/></xsl:attribute>
								<xsl:attribute name="TOTALLOANAMOUNT"><xsl:value-of select="LoanAmountRequired"/></xsl:attribute>
								<xsl:attribute name="TYPEOFAPPLICATION">10</xsl:attribute> <!-- New Loan -->
								<xsl:attribute name="TYPEOFBUYER">10</xsl:attribute> <!-- Not applicable -->
								<xsl:attribute name="NONPANELLENDERSELECTD">0</xsl:attribute>
								
								<!-- LOANCOMPONENTS are setup within pre-processing as the calculations between the products nodes
								and capital raising nodes are very complicated and not possible with XSL -->
								
								<!-- <xsl:for-each select="ProductsRoot/Products"> --> <!-- ProductsRoot\Products -->
									<!-- LOANCOMPONENT -->
									<!-- <xsl:element name="LOANCOMPONENT">
										<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="SalesProductCode"/></xsl:attribute>
										<xsl:attribute name="STARTDATE"></xsl:attribute>
										<xsl:attribute name="REPAYMENTMETHOD"><xsl:call-template name="FindRepaymentType"><xsl:with-param name="INTERESTONLYAMOUNT" select="ProductAmountInterestOnly" /><xsl:with-param name="CAPITALANDINTERESTAMOUNT" select="ProductAmount" /></xsl:call-template></xsl:attribute>
										<xsl:attribute name="LOANAMOUNT"><xsl:value-of select="ProductAmountInterestOnly + ProductAmount"/></xsl:attribute>
										<xsl:attribute name="CAPITALANDINTERESTELEMENT"><xsl:call-template name="OutputLoanAmountsIfNeeded"><xsl:with-param name="INTERESTONLYAMOUNT" select="ProductAmountInterestOnly" /><xsl:with-param name="CAPITALANDINTERESTAMOUNT" select="ProductAmount" /><xsl:with-param name="LOANAMOUNTNEEDED" select="'CAPITALANDINTEREST'" /></xsl:call-template></xsl:attribute>
										<xsl:attribute name="INTERESTONLYELEMENT"><xsl:call-template name="OutputLoanAmountsIfNeeded"><xsl:with-param name="INTERESTONLYAMOUNT" select="ProductAmountInterestOnly" /><xsl:with-param name="CAPITALANDINTERESTAMOUNT" select="ProductAmount" /><xsl:with-param name="LOANAMOUNTNEEDED" select="'INTERESTONLY'" /></xsl:call-template></xsl:attribute>
										<xsl:attribute name="PURPOSEOFLOAN">1</xsl:attribute> House purchase 
										<xsl:attribute name="TERMINMONTHS"><xsl:value-of select="/REQUEST/ProposalRoot/Proposal/TermMonths"/></xsl:attribute>
										<xsl:attribute name="TERMINYEARS"><xsl:value-of select="/REQUEST/ProposalRoot/Proposal/Term"/></xsl:attribute>
										<xsl:attribute name="TOTALLOANCOMPONENTAMOUNT"><xsl:value-of select="ProductAmountInterestOnly + ProductAmount"/></xsl:attribute>
										
									</xsl:element> --> <!-- LOANCOMPONENT -->
									
								<!-- </xsl:for-each> --> <!-- ProductsRoot\Products -->
								
							</xsl:element> <!-- MORTGAGESUBQUOTE -->
							
						</xsl:element> <!-- QUOTATION -->
						
						<!-- This is created simply because Cost Modelling doesnt work without it -->
						<!-- LIFESUBQUOTE -->
						<xsl:element name="LIFESUBQUOTE">
							<xsl:attribute name="LIFESUBQUOTENUMBER">1</xsl:attribute>
							<xsl:element name="LIFEBENEFIT"/>
						</xsl:element> <!-- LIFESUBQUOTE -->
						
						<!-- FINANCIALSUMMARY -->
						<xsl:element name="FINANCIALSUMMARY">
						
							<xsl:attribute name="LOANLIABILITYINDICATOR">
								<xsl:choose>
									<xsl:when test="count(ApplicantRoot/Applicant[1]/ApplicantCommitmentsRoot/ApplicantCommitments) > 0">
										<xsl:text>1</xsl:text>
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> <!-- LOANLIABILITYINDICATOR -->
							
							<!-- CCJHISTORYINDICATOR -->
							<xsl:attribute name="CCJHISTORYINDICATOR">
								<xsl:choose>
									<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and ApplicantRoot/Applicant[1]/CCJSYN = 1">
										<xsl:text>1</xsl:text> 
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> <!-- CCJHISTORYINDICATOR -->
							
							<!-- ARREARSHISTORYINDICATOR -->
							<xsl:attribute name="ARREARSHISTORYINDICATOR">
								<xsl:choose>
									<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and ApplicantRoot/Applicant[1]/ArrearsYN = 1">
										<xsl:text>1</xsl:text> 
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> <!-- ARREARSHISTORYINDICATOR -->
								
							<!-- BANKCARDINDICATOR -->
							<xsl:attribute name="BANKCARDINDICATOR">0</xsl:attribute>
								<!-- BANKCARDS are not being put into the bank card area <xsl:choose>
									<xsl:when test="count(ApplicantRoot/Applicant[1]/ApplicantCommitmentsRoot/ApplicantCommitments[AgreementTypeId = 2]) > 0">
										<xsl:text>1</xsl:text>
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> --> <!-- BANKCARDINDICATOR -->
							
							<!-- BANKRUPTCYHISTORYINDICATOR -->
							<xsl:attribute name="BANKRUPTCYHISTORYINDICATOR">
								<xsl:choose>
									<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and (ApplicantRoot/Applicant[1]/BankruptYN = 1 or ApplicantRoot/Applicant[1]/IVASYN = 1)">
										<xsl:text>1</xsl:text>
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> <!-- BANKRUPTCYHISTORYINDICATOR -->
							
							<xsl:attribute name="EXISTINGMORTGAGEINDICATOR">
								<xsl:choose>
									<xsl:when test="count(ApplicantRoot/Applicant/ApplicantAddressRoot/ApplicantAddress[ResidentialStatusTypeId = 2 and HaveMortgageYN = 1]) > 0">
										<xsl:text>1</xsl:text>
									</xsl:when>
									<xsl:otherwise>
										<xsl:text>0</xsl:text>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute> <!-- EXISTINGMORTGAGEINDICATOR -->
							
							<!-- Declined mortgages and life products are asked by GIWS -->
							<xsl:attribute name="DECLINEDMORTGAGEINDICATOR">0</xsl:attribute>
							
							<xsl:attribute name="LIFEPRODUCTINDICATOR">0</xsl:attribute>
							
						</xsl:element> <!-- FINANCIALSUMMARY -->
						
						<!-- Direct Debit details -->
						<xsl:for-each select="DirectDebit">
						
							<!-- DIRECTDEBITADDRESS -->
							<xsl:element name="DIRECTDEBITADDRESS">
								
								<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="Address/BuildingName"/></xsl:attribute>
								<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="Address/BuildingNo"/></xsl:attribute>
								<xsl:attribute name="FLATNUMBER"></xsl:attribute>
								<xsl:attribute name="STREET"><xsl:value-of select="Address/Street"/></xsl:attribute>
								<xsl:attribute name="DISTRICT"><xsl:value-of select="Address/District"/></xsl:attribute>
								<xsl:attribute name="TOWN"><xsl:value-of select="Address/Town"/></xsl:attribute>
								<xsl:attribute name="COUNTY"><xsl:value-of select="Address/County"/></xsl:attribute>
								<xsl:attribute name="COUNTRY"></xsl:attribute>
								<xsl:attribute name="POSTCODE"><xsl:value-of select="Address/PostCode"/></xsl:attribute>
								
								<!-- DIRECTDEBITTHIRDPARTY -->
								<xsl:element name="DIRECTDEBITTHIRDPARTY">
									<xsl:attribute name="THIRDPARTYTYPE">3</xsl:attribute> <!-- Bank/ Building Soc/ Other Lender Thirdparty type -->
									<xsl:attribute name="COMPANYNAME"><xsl:value-of select="BankName"/></xsl:attribute>
									<xsl:attribute name="BRANCHNAME"><xsl:call-template name="SubstituteBlankString"><xsl:with-param name="VALUE" select="BranchName" /><xsl:with-param name="OUTPUTVALUE" select="'Unknown'" /></xsl:call-template></xsl:attribute>
									<xsl:attribute name="THIRDPARTYBANKSORTCODE"><xsl:value-of select="SortCode"/></xsl:attribute>
									
									<!-- APPLICATIONBANKBUILDINGSOC -->
									<xsl:element name="APPLICATIONBANKBUILDINGSOC">
										<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="AccountNo"/></xsl:attribute>
										<xsl:attribute name="ACCOUNTNAME"><xsl:value-of select="AccountName"/></xsl:attribute>
										<xsl:attribute name="REPAYMENTBANKACCOUNTINDICATOR">1</xsl:attribute> <!-- Indicates this is the account that the repayment will be taken from -->
										
									</xsl:element> <!-- APPLICATIONBANKBUILDINGSOC -->
						
								</xsl:element> <!-- DIRECTDEBITTHIRDPARTY -->
							
							</xsl:element> <!-- DIRECTDEBITADDRESS -->
						
						</xsl:for-each> <!-- Direct Debit details -->
						
						<!-- Credit card details are no longer ingested, as payments will be taken on the internet, and these will be processed in postprocessing
						- Credit card payment details -
						<xsl:for-each select="CardPayment"> 
							
							- APPLICATIONCREDITCARD -
							<xsl:element name="APPLICATIONCREDITCARD">
								<xsl:attribute name="FEEPAYMENTMETHOD">1</xsl:attribute> - Yes the credit card will be used to pay any fees -
								<xsl:attribute name="NAMEOFCARDHOLDERS"><xsl:value-of select="PaymentCardHolderName"/></xsl:attribute>
								<xsl:attribute name="CARDTYPE"><xsl:value-of select="PaymentTypeId"/></xsl:attribute>
								<xsl:attribute name="CARDNUMBER"><xsl:value-of select="PaymentCardNumber"/></xsl:attribute>
								<xsl:attribute name="ISSUENUMBER"><xsl:value-of select="PaymentIssueNumber"/></xsl:attribute>
								<xsl:attribute name="EXPIRYDATE"><xsl:value-of select="PaymentExpiryMonthTypeId"/><xsl:text>/</xsl:text><xsl:value-of select="substring(PaymentExpiryYearTypeId,3,2)"/></xsl:attribute> - Only the last 2 digits of the year are used -
								
							</xsl:element> - APPLICATIONCREDITCARD -
						
						</xsl:for-each> - Credit card payment details -
						-->
						
						<!-- Solicitors Details -->
						<xsl:for-each select="Solicitor"> 
							<!-- A change was put into Omiga so that solicitor details are copied from the nameandaddress 
							directory rather than having a link, this was done at the request of BM, so even though a search 
							is being done on POS im still going to create the records in the thirdparty tables -->
							<!-- Is this solicitor from the panel -->
							<!-- <xsl:choose>
								
								- Solicitor is on the panel -
								<xsl:when test="string-length(OmigaNameAndAddressGUID) > 1">
									
									<xsl:element name="APPLICATIONLEGALREP">
										<xsl:attribute name="SEPARATELEGALREPRESENTATIVE">0</xsl:attribute>
										<xsl:attribute name="DIRECTORYGUID"><xsl:value-of select="OmigaNameAndAddressGUID"/></xsl:attribute>
									</xsl:element> - APPLICATIONLEGALREP -
									
								</xsl:when> - Solicitor is on the panel -
							-->	
								<!-- <xsl:otherwise> - Solicitor isnt on the panel - -->
									
									<!-- SOLICITORDETAILS -->
									<xsl:element name="SOLICITORDETAILS">
										<xsl:attribute name="CONTACTFORENAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'FORENAME'" /><xsl:with-param name="FULLNAME" select="SolicitorName" /></xsl:call-template></xsl:attribute>
										<xsl:attribute name="CONTACTSURNAME"><xsl:call-template name="FindNamePortion"><xsl:with-param name="PORTION" select="'SURNAME'" /><xsl:with-param name="FULLNAME" select="SolicitorName" /></xsl:call-template></xsl:attribute>
										<xsl:attribute name="CONTACTTYPE">10</xsl:attribute>  <!-- Organisation Contact Type -->
										
										<!-- SOLICITORADDRESS -->
										<xsl:element name="SOLICITORADDRESS">
											
											<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="Address/BuildingName"/></xsl:attribute>
											<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="Address/BuildingNo"/></xsl:attribute>
											<xsl:attribute name="FLATNUMBER"></xsl:attribute>
											<xsl:attribute name="STREET"><xsl:value-of select="Address/Street"/></xsl:attribute>
											<xsl:attribute name="DISTRICT"><xsl:value-of select="Address/District"/></xsl:attribute>
											<xsl:attribute name="TOWN"><xsl:value-of select="Address/Town"/></xsl:attribute>
											<xsl:attribute name="COUNTY"><xsl:value-of select="Address/County"/></xsl:attribute>
											<xsl:attribute name="COUNTRY"></xsl:attribute>
											<xsl:attribute name="POSTCODE"><xsl:value-of select="Address/PostCode"/></xsl:attribute>
											
											<!-- SOLICITORTHIRDPARTY -->
											<xsl:element name="SOLICITORTHIRDPARTY">
												<xsl:attribute name="THIRDPARTYTYPE">10</xsl:attribute> <!-- Legal representative Thirdparty type -->
												<xsl:attribute name="COMPANYNAME"><xsl:value-of select="SolicitorFirm"/></xsl:attribute>
												
												<xsl:element name="APPLICATIONLEGALREP">
													<xsl:attribute name="SEPARATELEGALREPRESENTATIVE">0</xsl:attribute>
													
												</xsl:element> <!-- APPLICATIONLEGALREP -->
												
											</xsl:element> <!-- SOLICITORTHIRDPARTY -->
											
										</xsl:element> <!-- SOLICITORADDRESS -->
									
									</xsl:element> <!-- SOLICITORDETAILS -->
									
								<!-- </xsl:otherwise> - Solicitor isnt on the panel -
								
							</xsl:choose> - Is this solicitor from the panel - -->
							
						</xsl:for-each> <!-- Solicitors Details -->
						
						<!-- MemoPad entry for Additional details -->
						<!-- MEMOPAD -->
						<xsl:element name="MEMOPAD">
							<xsl:attribute name="ENTRYDATETIME"><xsl:value-of select="$Now"/></xsl:attribute>
							<xsl:attribute name="MEMOENTRY"><xsl:call-template name="OutputAdditionalDetails"/></xsl:attribute>
							<xsl:attribute name="USERID"><xsl:value-of select="$UserID"/></xsl:attribute>
							<xsl:attribute name="ENTRYTYPE">1</xsl:attribute> <!-- Advisor Comment -->
						</xsl:element> <!-- MEMOPAD -->
						
						<!-- Verification Questions -->
						<!-- Question 1 - Any Arrears -->
						<!-- APPLNADDITIONALQUESTIONS -->
						<xsl:element name="APPLNADDITIONALQUESTIONS">			
							<xsl:attribute name="QUESTIONREFERENCE">1</xsl:attribute> <!-- Any arrears -->
							<xsl:attribute name="TYPE">40</xsl:attribute> <!-- Verification -->
							<xsl:choose>
								<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and ApplicantRoot/Applicant[1]/ArrearsYN = 1">
									<xsl:attribute name="RESPONSE">1</xsl:attribute> <!-- Yes -->
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="RESPONSE">0</xsl:attribute> <!-- No -->
								</xsl:otherwise>
							</xsl:choose>
						</xsl:element> <!-- APPLNADDITIONALQUESTIONS -->
						
						<!-- Question 2 - Any CCJ's -->
						<!-- APPLNADDITIONALQUESTIONS -->
						<xsl:element name="APPLNADDITIONALQUESTIONS">			
							<xsl:attribute name="QUESTIONREFERENCE">2</xsl:attribute> <!-- Any ccjs -->
							<xsl:attribute name="TYPE">40</xsl:attribute> <!-- Verification -->
							<xsl:choose>
								<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and ApplicantRoot/Applicant[1]/CCJSYN = 1">
									<xsl:attribute name="RESPONSE">1</xsl:attribute> <!-- Yes -->
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="RESPONSE">0</xsl:attribute> <!-- No -->
								</xsl:otherwise>
							</xsl:choose>
						</xsl:element> <!-- APPLNADDITIONALQUESTIONS -->
						
						<!-- Question 3 - Any bankrupt or iva details -->
						<!-- APPLNADDITIONALQUESTIONS -->
						<xsl:element name="APPLNADDITIONALQUESTIONS">			
							<xsl:attribute name="QUESTIONREFERENCE">3</xsl:attribute> <!-- Any bankruptcy -->
							<xsl:attribute name="TYPE">40</xsl:attribute> <!-- Verification -->
							<xsl:choose>
								<xsl:when test="ApplicantRoot/Applicant[1]/AdverseHistoryYN = 1 and (ApplicantRoot/Applicant[1]/BankruptYN = 1 or ApplicantRoot/Applicant[1]/IVASYN = 1)">
									<xsl:attribute name="RESPONSE">1</xsl:attribute> <!-- Yes -->
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="RESPONSE">0</xsl:attribute> <!-- No -->
								</xsl:otherwise>
							</xsl:choose>
						</xsl:element> <!-- APPLNADDITIONALQUESTIONS -->
						
						<!-- Question 4 - Applicants convicted -->
						<!-- APPLNADDITIONALQUESTIONS -->
						<xsl:element name="APPLNADDITIONALQUESTIONS">			
							<xsl:attribute name="QUESTIONREFERENCE">4</xsl:attribute> <!-- Any arrears -->
							<xsl:attribute name="TYPE">40</xsl:attribute> <!-- Verification -->
							<xsl:attribute name="RESPONSE"><xsl:value-of select="ApplicantRoot/Applicant[1]/ConvictionsYN"/></xsl:attribute>
							<!-- If the applicant has been convicted put in the details --> 
							<xsl:if test="ApplicantRoot/Applicant[1]/ConvictionsYN = 1">
								<xsl:attribute name="FURTHERDETAILS"><xsl:call-template name="OutputConvictionsFurtherDetails"/></xsl:attribute>
							</xsl:if> <!-- Applicant convicted -->
							
						</xsl:element> <!-- APPLNADDITIONALQUESTIONS -->
						
						<!-- Question 5 - Owner of other properties -->
						<!-- APPLNADDITIONALQUESTIONS -->
						<xsl:element name="APPLNADDITIONALQUESTIONS">			
							<xsl:attribute name="QUESTIONREFERENCE">5</xsl:attribute> <!-- Any arrears -->
							<xsl:attribute name="TYPE">40</xsl:attribute> <!-- Verification -->
							<xsl:attribute name="RESPONSE"><xsl:value-of select="ApplicantRoot/Applicant[1]/OtherPropertiesYN"/></xsl:attribute>
							
						</xsl:element> <!-- APPLNADDITIONALQUESTIONS -->
						
						<!-- Attitude to borrowing - type of advice given -->
						<!-- ATTITUDETOBORROWING -->
						<xsl:element name="ATTITUDETOBORROWING">
							<xsl:attribute name="QUESTIONNUMBER">1</xsl:attribute> <!--Its always the first question -->
							<xsl:attribute name="RESPONSETOQUESTION"><xsl:value-of select="LevelOfAdviceTypeId"/></xsl:attribute>
							
						</xsl:element> <!-- ATTITUDETOBORROWING -->
								
					</xsl:element> <!-- APPLICATIONFACTFIND -->
					
					<!-- Details to set up task management -->
					<!-- CASEACTIVITY -->
					<xsl:element name="CASEACTIVITY">
						<xsl:attribute name="SOURCEAPPLICATION">Omiga</xsl:attribute>
						<xsl:attribute name="CASEID">../@APPLICATIONNUMBER</xsl:attribute>
						<xsl:attribute name="ACTIVITYID">10</xsl:attribute> <!-- Mortgage Application -->
						<xsl:attribute name="ACTIVITYINSTANCE">1</xsl:attribute> <!-- This is the first activity instance -->
						
						<!-- CASESTAGE -->
						<xsl:element name="CASESTAGE">
							<xsl:attribute name="CASESTAGESEQUENCENO">1</xsl:attribute> <!-- This is the first case stage -->
							<xsl:attribute name="STAGEID">10</xsl:attribute> <!-- Customer Registration -->
							<xsl:attribute name="STAGEORIGINATINGDATETIME"><xsl:value-of select="$Now"/></xsl:attribute>
							<xsl:attribute name="STAGEORIGINATINGUNITID"><xsl:value-of select="$UnitID"/></xsl:attribute>
							<xsl:attribute name="STAGEORIGINATINGUSERID"><xsl:value-of select="$UserID"/></xsl:attribute>
							
						</xsl:element> <!-- CASESTAGE -->
						
					</xsl:element> <!-- CASEACTIVITY -->
					
					<!-- APPLICATIONSTAGE -->
					<xsl:element name="APPLICATIONSTAGE">
						<xsl:attribute name="APPLICATIONNUMBER">../@APPLICATIONNUMBER</xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER">../APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER</xsl:attribute>
						<xsl:attribute name="CASESEQUENCENO">1</xsl:attribute> 
						<xsl:attribute name="STAGENUMBER">10</xsl:attribute> <!-- Customer Registration -->
						<xsl:attribute name="STAGENAME">Customer Registration</xsl:attribute> <!-- Name of the stage to display -->
						<xsl:attribute name="DATETIME"><xsl:value-of select="$Now"/></xsl:attribute>
						<xsl:attribute name="STAGESEQUENCENO">1</xsl:attribute> <!-- This is the first stage the application has entered -->
						
					</xsl:element> <!-- APPLICATIONSTAGE -->
					
				</xsl:element> <!-- APPLICATION -->
				
			</xsl:for-each> <!-- ProposalRoot/Proposal -->
			
		</xsl:element> <!-- REQUEST -->

	</xsl:template> <!-- Main "/" -->

	<!-- 
	Template - BuildApplicationCorrespondenceSaluation
	Purpose - Builds the salutation for all the customers within this application, i.e. Mr Smith, Mrs Jones
	Author - MO
	Date - 24/07/2002
	-->
	<xsl:template name="BuildApplicationCorrespondenceSalutation">
		<xsl:for-each select="/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant">
			<xsl:if test="position() != 1">
				<xsl:text>, </xsl:text>
			</xsl:if>
			<xsl:value-of select="Forename1" /><xsl:text> </xsl:text><xsl:value-of select="Surname" />
		</xsl:for-each>
	</xsl:template> <!-- BuildApplicationCorrespondenceSalutation -->

	<!-- 
	(No longer used as type of buyer is now always 'Not Applicable'
	Template - FindTypeOfBuyer
	Purpose - Determines what the application type of buyer should be.
	Author - MO
	Date - 24/07/2002
	-->
	<!-- <xsl:template name="FindTypeOfBuyer">
		- Run the rules to find out what sort of application this is -
		<xsl:choose>
			- are all the applicants first time buyers? -
			<xsl:when test="count(/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant[FirstTimeBuyerYN=1]) = count(/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant)">
				<xsl:text>10</xsl:text> - First time buyer -
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>30</xsl:text> - Subsequent -
			</xsl:otherwise>			
		</xsl:choose>
	</xsl:template> - FindTypeOfBuyer - -->
	
	<!-- 
	Template - FindTelephoneNoPortion
	Purpose - Splits individual items from the telephone number, 'AREACODE', 'TELEPHONENUMBER'
	Author - MO
	Date - 24/07/2002
		   25/11/2002 - Modified so that if the space occurs in the telephone number after 5 chars 
						it only returns the first 5 chars
	-->
	<xsl:template name="FindTelephoneNoPortion">
		<xsl:param name="PORTION"/> <!-- This should be either 'AREACODE' or 'TELEPHONENUMBER' -->
		<xsl:param name="FULLTELEPHONENUMBER"/>
		<xsl:if	test="$PORTION = 'AREACODE'">
			<!-- Is there a space in the telephone number and if there is its not greater than 5 chars -->
			<xsl:choose>
				<xsl:when test="substring-before($FULLTELEPHONENUMBER, ' ') != '' and string-length(substring-before($FULLTELEPHONENUMBER, ' ')) &lt;= 5">
					<xsl:value-of select="substring-before($FULLTELEPHONENUMBER, ' ')" />
				</xsl:when>
				<!-- No theres no space, take the first 5 chars -->
				<xsl:otherwise>
					<xsl:value-of select="substring($FULLTELEPHONENUMBER, 1, 5)" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
		<xsl:if	test="$PORTION = 'TELEPHONENUMBER'">
			<!-- Is there a space in the telephone number and if there is its not greater than 5 chars -->
			<xsl:choose>
				<xsl:when test="substring-after($FULLTELEPHONENUMBER, ' ') != '' and string-length(substring-before($FULLTELEPHONENUMBER, ' ')) &lt;= 5">
					<xsl:value-of select="substring-after($FULLTELEPHONENUMBER, ' ')" />
				</xsl:when>
				<!-- No theres no space, take the chars from the 6th char-->	
				<xsl:otherwise>
					<xsl:value-of select="substring($FULLTELEPHONENUMBER, 6)" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template> <!-- FindTelephoneNoPortion -->

	<!-- 
	Template - FindNamePortion
	Purpose - Splits individual forename and surname from the single name string, 'FORENAME', 'SURNAME'
	Author - MO
	Date - 24/07/2002
	-->
	<xsl:template name="FindNamePortion">
		<xsl:param name="PORTION"/> <!-- This should be either 'FORENAME' or 'SURNAME' -->
		<xsl:param name="FULLNAME"/>
		<xsl:if	test="$PORTION = 'FORENAME'">
			<!-- Is there a space in the name -->
			<xsl:choose>
				<xsl:when test="substring-before($FULLNAME, ' ') != ''">
					<xsl:value-of select="substring-before($FULLNAME, ' ')" />
				</xsl:when>
				<!-- No theres no space, return nothing cause the surname portion will return the whole thing-->
			</xsl:choose>
		</xsl:if>
		<xsl:if	test="$PORTION = 'SURNAME'">
			<!-- Is there a space in the name -->
			<xsl:choose>
				<xsl:when test="substring-after($FULLNAME, ' ') != ''">
					<xsl:value-of select="substring-after($FULLNAME, ' ')" />
				</xsl:when>
				<!-- No theres no space, return the whole thing-->	
				<xsl:otherwise>
					<xsl:value-of select="$FULLNAME" />
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template> <!-- FindNamePortion -->
	
	<!-- 
	Template - FindRepaymentType
	Purpose - determines from the interestonly and capital and interest amounts where this loan component is 
				a capital and interest, an interest only or a part and part component
	Author - MO
	Date - 24/07/2002
	-->
	<xsl:template name="FindRepaymentType">
		<xsl:param name="INTERESTONLYAMOUNT"/> 
		<xsl:param name="CAPITALANDINTERESTAMOUNT"/>
		<xsl:choose>
			<!-- Interest only -->
			<xsl:when test="$INTERESTONLYAMOUNT != '0' and $CAPITALANDINTERESTAMOUNT = '0'">
				<xsl:text>1</xsl:text>
			</xsl:when>
			<!-- Capital and Interest -->
			<xsl:when test="$INTERESTONLYAMOUNT = '0' and $CAPITALANDINTERESTAMOUNT != '0'">
				<xsl:text>2</xsl:text>
			</xsl:when>
			<!-- Part and Part -->
			<xsl:when test="$INTERESTONLYAMOUNT != '0' and $CAPITALANDINTERESTAMOUNT != '0'">
				<xsl:text>3</xsl:text>
			</xsl:when>
		</xsl:choose>
	</xsl:template> <!-- FindRepaymentType -->
	
	<!-- 
	Template - OutputLoanAmountsIfNeeded
	Purpose - omiga only expects interestonly and capitalandinterest elements if its a part and part
	Author - MO
	Date - 24/07/2002
	-->
	<xsl:template name="OutputLoanAmountsIfNeeded">
		<xsl:param name="INTERESTONLYAMOUNT"/> 
		<xsl:param name="CAPITALANDINTERESTAMOUNT"/>
		<xsl:param name="LOANAMOUNTNEEDED"/> <!-- 'CAPITALANDINTEREST' OR 'INTERESTONLY' -->
		<!-- Only if its a part and part -->
		<xsl:if test="$INTERESTONLYAMOUNT != '0' and $CAPITALANDINTERESTAMOUNT != '0'">
			<xsl:choose>
				<xsl:when test="$LOANAMOUNTNEEDED = 'CAPITALANDINTEREST'">
					<xsl:value-of select="$CAPITALANDINTERESTAMOUNT"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="$INTERESTONLYAMOUNT"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template> <!-- FindRepaymentType -->
	
	<!-- 
	Template - SubstituteBlankString
	Purpose - If its a blank string subsititute it for the value passed it
	Author - MO
	Date -19/08/2002
	-->
	<xsl:template name="SubstituteBlankString">
		<xsl:param name="VALUE"/> 
		<xsl:param name="OUTPUTVALUE"/>
		<!-- If the input value is nothing output the output value -->
		<xsl:choose>
		<xsl:when test="string-length($VALUE) = 0">
			<xsl:value-of select="$OUTPUTVALUE"/>
		</xsl:when>
		<xsl:otherwise>
			<xsl:value-of select="$VALUE"/>
		</xsl:otherwise>
		</xsl:choose>
	</xsl:template> <!-- SubstituteBlankString -->
	
	<!-- 
	Template - DetermineApplicantUKResident
	Purpose - Checks the applicants current address country, if its in the UK, returns 1, else 0
	Author - MO
	Date -30/08/2002
	-->
	<xsl:template name="DetermineApplicantUKResident">
		<xsl:param name="COUNTRY"/> 
		<xsl:choose>
			<xsl:when test="$COUNTRY = 1"> <!-- United Kingdom -->
				<xsl:text>1</xsl:text>
			</xsl:when>
			<xsl:when test="string-length($COUNTRY) = 0">
				<xsl:text></xsl:text>
			</xsl:when>
			<xsl:otherwise>
				<xsl:text>0</xsl:text>
			</xsl:otherwise>
		</xsl:choose>
		
	</xsl:template> <!-- DetermineApplicantUKResident -->

	<!-- 
	Template - OutputBuildingNameOrNumber
	Purpose - Checks to see if the value for building name/number is of a specific type and outputs it if needed
	Author - MO
	Date -30/08/2002
	-->
	<xsl:template name="OutputBuildingNameOrNumber">
		<xsl:param name="TYPE"/> <!-- The type you want to output 'NAME' or 'NUMBER'-->
		<xsl:param name="BUILDINGNAMENUMBER"/> <!-- This is the value to be outputted if applicable -->
		<xsl:choose>
			<xsl:when test="$TYPE = 'NAME'"> 
				<!-- Check to see if its not a number, basically convert  -->
				<xsl:if test="number($BUILDINGNAMENUMBER) != $BUILDINGNAMENUMBER">
					<xsl:value-of select="$BUILDINGNAMENUMBER"/>
				</xsl:if>
			</xsl:when>
			<xsl:when test="$TYPE = 'NUMBER'">
				<xsl:if test="number($BUILDINGNAMENUMBER) = $BUILDINGNAMENUMBER">
					<xsl:value-of select="$BUILDINGNAMENUMBER"/>
				</xsl:if>
			</xsl:when>
		</xsl:choose>
		
	</xsl:template> <!-- OutputBuildingNameOrNumber -->

	
	<!-- 
	Template - IsDateOlderThanYears
	Purpose - Checks to see if a date passed in is older than a number of years passed in from now.
				This function isnt called by code but is an excellent example of comparing dates
	Author - MO
	Date -20/08/2002
	-->
	<xsl:template name="IsDateOlderThanYears">
		<xsl:param name="NOW"/> 
		<xsl:param name="DATE"/>
		<xsl:param name="NUMBEROFYEARS"/>
		
		<!-- Split the 2 dates into day, month and years -->
		<!-- This code copes with dates formatted as : 
			'dd/mm/yyyy hh:nn:ss'
			'dd/mm/yyyy'
			'd/m/yyyy hh:nn:ss'
			'd/m/yyyy'
			and all formats list above with 'yy' 
			Please node that any time is ignored.
			-->
		<xsl:variable name="NOWDAY" select="substring-before($NOW, '/')"/>
		<xsl:variable name="NOWMONTH" select="substring-before(substring-after($NOW,'/'), '/')"/>
		<xsl:variable name="NOWYEAR">
		<xsl:choose>
			<!-- If a blank string is returned in the year a time wasnt passed in -->
			<xsl:when test="string-length(substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')) = 0">
				<xsl:value-of select="substring-after(substring-after($NOW,'/'), '/')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')"/>
			</xsl:otherwise>
		</xsl:choose>
		</xsl:variable>
		
		<xsl:variable name="DATEDAY" select="substring-before($DATE, '/')"/>
		<xsl:variable name="DATEMONTH" select="substring-before(substring-after($DATE,'/'), '/')"/>
		<xsl:variable name="DATEYEAR">
		<xsl:choose>
			<!-- If a blank string is returned in the year a time wasnt passed in -->
			<xsl:when test="string-length(substring-before(substring-after(substring-after($DATE,'/'), '/'), ' ')) = 0">
				<xsl:value-of select="substring-after(substring-after($DATE,'/'), '/')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="substring-before(substring-after(substring-after($DATE,'/'), '/'), ' ')"/>
			</xsl:otherwise>
		</xsl:choose>
		</xsl:variable>
		
		<xsl:element name="NOW"><xsl:value-of select="$NOWDAY"/>:<xsl:value-of select="$NOWMONTH"/>:<xsl:value-of select="$NOWYEAR"/></xsl:element>
		<xsl:element name="DATE"><xsl:value-of select="$DATEDAY"/>:<xsl:value-of select="$DATEMONTH"/>:<xsl:value-of select="$DATEYEAR"/></xsl:element>
		
		<xsl:if test="($NOWYEAR - $DATEYEAR > $NUMBEROFYEARS) or ($NOWYEAR - $DATEYEAR = $NUMBEROFYEARS and $NOWMONTH > $DATEMONTH) or ($NOWYEAR - $DATEYEAR = $NUMBEROFYEARS and $NOWMONTH = $DATEMONTH and $NOWDAY >= $DATEDAY)">
			<xsl:element name="OLDERTHANAGE">1</xsl:element>
		</xsl:if>
		
	</xsl:template> <!-- IsDateOlderThanYears -->
	
	<!-- 
	Template - OutputOtherResidents
	Purpose - Loops through the dependancies nodes and outputs other residents that are older than 17
	Author - MO
	Date -20/08/2002
	-->
	<xsl:template name="OutputOtherResidents">
		<xsl:param name="NOW"/> 
		
		<xsl:variable name="YEARSOLDERTHAN" select="'17'"/>
		
		<xsl:for-each select="/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant/ApplicantDependantRoot/ApplicantDependant">
			
			<xsl:variable name="DATE" select="DOB"/>
		
			<!-- Split the 2 dates into day, month and years -->

			<xsl:variable name="NOWDAY" select="substring-before($NOW, '/')"/>
			<xsl:variable name="NOWMONTH" select="substring-before(substring-after($NOW,'/'), '/')"/>
			<xsl:variable name="NOWYEAR">
			<xsl:choose>
				<!-- If a blank string is returned in the year a time wasnt passed in -->
				<xsl:when test="string-length(substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')) = 0">
					<xsl:value-of select="substring-after(substring-after($NOW,'/'), '/')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')"/>
				</xsl:otherwise>
			</xsl:choose>
			</xsl:variable>
		
			<xsl:variable name="DATEDAY" select="substring-before($DATE, '/')"/>
			<xsl:variable name="DATEMONTH" select="substring-before(substring-after($DATE,'/'), '/')"/>
			<xsl:variable name="DATEYEAR">
			<xsl:choose>
				<!-- If a blank string is returned in the year a time wasnt passed in -->
				<xsl:when test="string-length(substring-before(substring-after(substring-after($DATE,'/'), '/'), ' ')) = 0">
					<xsl:value-of select="substring-after(substring-after($DATE,'/'), '/')"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="substring-before(substring-after(substring-after($DATE,'/'), '/'), ' ')"/>
				</xsl:otherwise>
			</xsl:choose>
			</xsl:variable>
			
			<!-- Is the dependant over 17, if so they are an other resident -->
			<xsl:if test="($NOWYEAR - $DATEYEAR > $YEARSOLDERTHAN) or ($NOWYEAR - $DATEYEAR = $YEARSOLDERTHAN and $NOWMONTH > $DATEMONTH) or ($NOWYEAR - $DATEYEAR = $YEARSOLDERTHAN and $NOWMONTH = $DATEMONTH and $NOWDAY >= $DATEDAY)">
				<!-- PERSON -->
				<xsl:element name="PERSON">
					<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="DOB"/></xsl:attribute>
					<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="Forename1"/></xsl:attribute>
					<xsl:attribute name="SECONDFORENAME"><xsl:value-of select="Forename2"/></xsl:attribute>
					<xsl:attribute name="SURNAME"><xsl:value-of select="Surname"/></xsl:attribute>
											
					<!-- OTHERRESIDENT -->
					<xsl:element name="OTHERRESIDENT"></xsl:element>
											
				</xsl:element> <!-- PERSON -->
				
			</xsl:if>
			
		</xsl:for-each> <!-- ApplicantDependantRoot/ApplicantDependant -->
		
	</xsl:template>
	
	<!-- 
	Template - OutputAdditionalDetails
	Purpose - The method builds the additional details text
	Author - MO
	Date -30/08/2002
	-->
	<xsl:template name="OutputAdditionalDetails">
		<!-- Additional Detail, taken from app form -->
		<xsl:text>POS Additional Info :&#13;&#10;</xsl:text>
		<xsl:value-of select="/REQUEST/ProposalRoot/Proposal/AdditionalInfo"/><xsl:text>&#13;&#10;&#13;&#10;</xsl:text>
	</xsl:template> <!-- OutputAdditionalDetails -->
	
	<!-- 
	Template - OutputCurrentYear
	Purpose - Takes the current Now value and extracts the year value from it
	Author - MO
	Date -20/08/2002
	-->
	<xsl:template name="OutputCurrentYear">
		<xsl:param name="NOW"/> 
		<xsl:choose>
			<!-- If a blank string is returned in the year a time wasnt passed in -->
			<xsl:when test="string-length(substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')) = 0">
				<xsl:value-of select="substring-after(substring-after($NOW,'/'), '/')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="substring-before(substring-after(substring-after($NOW,'/'), '/'), ' ')"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<!-- 
	Template - CalculateBasicIncome
	Purpose -	Takes the values for basic and other incomes and the other income type and 
				works out the value for basic income
	Author - MO
	Date -15/10/2002
	-->
	<xsl:template name="CalculateBasicIncome">
		<xsl:param name="BasicAnnualIncome"/>
		<xsl:param name="OtherIncome"/> 
		<xsl:param name="OtherIncomeTypeId"/>  
		
		<!-- Other income is not added to basic income if its type 'Benefits' = 5 -->
		<xsl:choose>
			<xsl:when test="$OtherIncomeTypeId != 5 and number($OtherIncome) > 0">
				<xsl:value-of select="number($BasicAnnualIncome) + number($OtherIncome)"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$BasicAnnualIncome"/>
			</xsl:otherwise>
		</xsl:choose>
		
	</xsl:template>
	
	<!-- 
	Template - OutputConvictionsFurtherDetails
	Purpose - The method builds the further details for the application convictions additional question 
	Author - MO
	Date -19/11/2002
	-->
	<xsl:template name="OutputConvictionsFurtherDetails">
		<!-- Conviction details, taken from the applicant risks -->
		<xsl:text>Conviction details :&#13;&#10;&#13;&#10;</xsl:text>
		<!-- Select the correct ApplicantRisk -->
		<xsl:for-each select="/REQUEST/ProposalRoot/Proposal/ApplicantRoot/Applicant[1]/ApplicantRisksRoot/ApplicantRisks[RiskTypeId=4]">
			<xsl:text>Date convicted - </xsl:text><xsl:value-of select="StartDate"/><xsl:text>&#13;&#10;</xsl:text>
			<xsl:text>Sentence given - </xsl:text><xsl:value-of select="Reason"/>
		</xsl:for-each>
	</xsl:template> <!-- OutputConvictionsFurtherDetails -->
	
</xsl:stylesheet>
