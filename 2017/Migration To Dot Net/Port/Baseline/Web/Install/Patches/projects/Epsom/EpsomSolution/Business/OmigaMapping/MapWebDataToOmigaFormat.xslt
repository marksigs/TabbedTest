<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: Transform Amatica to Omiga

History:

Version Author    Date       	   Description
01.01	LDM			24/01/2007  EP2_596  Transform Amatica xml into Omiga4 format	

================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>

	<xsl:variable name="Application" select="/Application" />
	<xsl:variable name="Applicant" select="/Application/Applicants/Applicant"/>
	<xsl:variable name="Guarantor" select="/Application/Guarantor" />
	<xsl:variable name="AddlInfo" select="/Application/AdditionalInformation2"/>
	<xsl:variable name="RegInfo" select="/Application/RegulationInformation"/>
	<xsl:variable name="PaymentHist" select="/Application/PaymentHistory"/>
	<xsl:variable name="PropertyDetail" select="/Application/PropertyDetail"/>
	<xsl:variable name="Solicitor" select="/Application/Solicitor"/>
	<xsl:variable name="AppNo" select="/Application/ApplicationNumber"/>
	<xsl:variable name="AppFFNo" select="/Application/ApplicationFactFindNumber"/>	
	<xsl:variable name="SysDateNow" select="/Application/@SYSDATE" />
	<xsl:variable name="CCPayment" select="/Application/CreditCardPayment/Fees/CCFeePayment" />
	
	<xsl:variable name="CustomerNumber1" select="string(/Application/Applicants/Applicant[1]/CustomerNumber)" />
	<xsl:variable name="CustomerNumber2" select="string(/Application/Applicants/Applicant[2]/CustomerNumber)" />
	<xsl:variable name="CustomerVersionNumber1" select="string(/Application/Applicants/Applicant[1]/CustomerVersionNumber)" />
	<xsl:variable name="CustomerVersionNumber2" select="string(/Application/Applicants/Applicant[2]/CustomerVersionNumber)" />
	<xsl:variable name="GuarantorNumber" select="string(/Application/Guarantor/CustomerNumber)" />
	<xsl:variable name="GuarantorVersionNumber" select="string(/Application/Guarantor/CustomerVersionNumber)" />
	
	<xsl:template match="/">
		<xsl:element name="APPLICATION">
			<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
				<xsl:if test="$RegInfo/FeeToCustomerAmount!=''">
					<xsl:attribute name="PROCFEETOCUSTAMOUNT"><xsl:value-of select="$RegInfo/FeeToCustomerAmount"/></xsl:attribute>				
				</xsl:if>
				<xsl:if test="$RegInfo/FeeToCustomerKnown!=''">
					<xsl:attribute name="PROCFEETOCUST"><xsl:value-of select="$RegInfo/FeeToCustomerKnown"/></xsl:attribute>				
				</xsl:if>			
			<xsl:element name="APPLICATIONFACTFIND">
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
				<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
				<xsl:attribute name="PRODUCTSCHEME"><xsl:value-of select="/Application/ProductScheme"/></xsl:attribute>

				<xsl:if test="$RegInfo/RegulatedLoan!=''">
					<xsl:attribute name="REGULATIONINDICATOR"><xsl:value-of select="$RegInfo/RegulatedLoan" /></xsl:attribute>
				</xsl:if>
				<xsl:if test="$RegInfo/LevelOfService!=''">
					<xsl:attribute name="LEVELOFADVICE"><xsl:value-of select="$RegInfo/LevelOfService" /></xsl:attribute>									
				</xsl:if>

				<xsl:if test="$RegInfo/CustomerAcceptedKFI">
					<xsl:attribute name="KFIRECEIVEDBYALLAPPS">
						<xsl:choose>
							<xsl:when test="$RegInfo/CustomerAcceptedKFI/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
							<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
						</xsl:choose>						
					</xsl:attribute>
				</xsl:if>
				
				<!-- PURCHASEPRICEORESTIMATEDVALUEonly set for Remortgages -->
				<xsl:if test="$PropertyDetail/RemortgageEstimatedPropertyValue!='' and $PropertyDetail/RemortgageEstimatedPropertyValue!='0'">
					<xsl:attribute name="PURCHASEPRICEORESTIMATEDVALUE">
						<xsl:value-of select="$PropertyDetail/RemortgageEstimatedPropertyValue" />
					</xsl:attribute> 
				</xsl:if>								
				
				<xsl:if test="$PropertyDetail/MainResidence/IsNull='false'">
					<xsl:attribute name="MAINRESIDENCEIND">
						<xsl:choose>
							<xsl:when test="$PropertyDetail/MainResidence/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
							<xsl:otherwise><xsl:value-of select="'0'"/></xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>

				<xsl:if test="$PropertyDetail/MainResidence/IsNull='true'">
					<xsl:attribute name="PACKAGERVALUATIONFEE">
						<xsl:choose>
							<xsl:when test="$PropertyDetail/MainResidence/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
							<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>

				<!-- ik EP2_1169 NEWLOAN entries -->
				<xsl:element name="NEWLOAN">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
					<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="$Application/LoanAmountRequested"/></xsl:attribute>
					<xsl:if test="$Application/LoanAmountRequested != $Application/CreditLimit">
						<xsl:attribute name="DIPDRAWDOWN">
							<xsl:value-of select="$Application/CreditLimit - $Application/LoanAmountRequested"/>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$Application/IsFinancialBenefitOfAll/IsNull = 'false'">
						<xsl:attribute name="DIRECTFINANCIALBENEFITIND">
							<xsl:choose>
								<xsl:when test="$Application/IsFinancialBenefitOfAll/Value = 'true'">1</xsl:when>
								<xsl:otherwise>0</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
				</xsl:element>
				<!-- ik EP2_1169 ends -->

				 <!-- MEMOPAD Entries-->
				<xsl:for-each select="$Application/MemoPad/*">
					<xsl:call-template name="CreateMemoPad">
						<xsl:with-param name="MemoPadId" select="MemoPadId"/>
						<xsl:with-param name="EntryType" select="EntryType"/>
						<xsl:with-param name="MemoEntry" select="MemoEntry"/>
					</xsl:call-template>					
				</xsl:for-each>

				<!-- <xsl:for-each select="$Applicant | $Guarantor"> -->
				<xsl:for-each select="$Applicant">
					<xsl:variable name="CustIncome" select="Income"/>
					<xsl:variable name="CustomerNumber" select="CustomerNumber"/>
					<xsl:variable name="CustomerVersionNumber" select="CustomerVersionNumber"/>
					<xsl:variable name="Employment" select="Employments/Employment" />
						
					<xsl:element name="CUSTOMERROLE">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo" /></xsl:attribute>
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
						<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
						<xsl:element name="CUSTOMER">
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
							<xsl:element name="CUSTOMERVERSION">
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
								<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
								
								<xsl:if test="UnitedKingdomResident/IsNull='false'">
									<xsl:attribute name="UKRESIDENTINDICATOR">
										<xsl:choose>
											<xsl:when test="UnitedKingdomResident/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
											<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>								
								</xsl:if>
								<xsl:if test="DiplomaticImmunity/IsNull!='true'">
									<xsl:attribute name="DIPLOMATICIMMUNITY">	
										<xsl:choose>
											<xsl:when test="DiplomaticImmunity/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
											<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</xsl:if>
								<xsl:if test="UnitedKingdomWorkAndResidePermitted/IsNull!='true'">
									<xsl:attribute name="RIGHTTOWORK">
										<xsl:choose>
											<xsl:when test="UnitedKingdomWorkAndResidePermitted/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
											<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</xsl:if>
								<xsl:attribute name="COUNTRYOFRESIDENCY"><xsl:value-of select="CountryOfResidency"/></xsl:attribute>
								<xsl:attribute name="REASONFORCOUNTRYOFRESIDENCY"><xsl:value-of select="UnitedKingdomNonResidentExplanation"/></xsl:attribute>
								<xsl:if test="HousingBenefit/IsNull='false'">
									<xsl:attribute name="HOUSINGBENEFIT">
										<xsl:choose>
											<xsl:when test="HousingBenefit/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
											<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</xsl:if>

								<xsl:attribute name="NATIONALINSURANCENUMBER"><xsl:value-of select="NationalInsuranceNumber"/></xsl:attribute>
								
								<!-- ik EP2_1664 marketing opt-in -->
								<xsl:attribute name="MAILSHOTREQUIRED">
									<xsl:choose>
										<xsl:when test="MarketingDataProtectionOptIn='true'">1</xsl:when>
										<xsl:otherwise>0</xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<!-- ik EP2_1664 ends -->

								<!-- DECLINEDMORTGAGE -->
								<!-- Probably a better method of doing this. But this works for now! -->
								<xsl:choose>
									<xsl:when test="$CustomerNumber = $CustomerNumber1">
										<xsl:for-each select="$PaymentHist/DeclinedMortgages/DeclinedMortgage[Applicant1='true']">
											<xsl:element name="DECLINEDMORTGAGE">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
												<xsl:if test="SequenceNumber!='' and SequenceNumber!='0'">
													<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="SequenceNumber" /></xsl:attribute>
												</xsl:if>
												<xsl:attribute name="DECLINEDDETAILS"><xsl:value-of select="RefusedMortgageDetails" /></xsl:attribute>
											</xsl:element>
										</xsl:for-each>
									</xsl:when>
									<xsl:when test="$CustomerNumber = $CustomerNumber2">
										<xsl:for-each select="$PaymentHist/DeclinedMortgages/DeclinedMortgage[Applicant2='true']">
											<xsl:element name="DECLINEDMORTGAGE">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
												<xsl:if test="SequenceNumber!='' and SequenceNumber!='0'">
													<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="SequenceNumber" /></xsl:attribute>
												</xsl:if>
												<xsl:attribute name="DECLINEDDETAILS"><xsl:value-of select="RefusedMortgageDetails" /></xsl:attribute>
											</xsl:element>
										</xsl:for-each>
									</xsl:when>
									<xsl:when test="$CustomerNumber = $GuarantorNumber">
										<xsl:for-each select="$PaymentHist/DeclinedMortgages/DeclinedMortgage[Guarantor='true']">
											<xsl:element name="DECLINEDMORTGAGE">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
												<xsl:if test="SequenceNumber!='' and SequenceNumber!='0'">
													<xsl:attribute name="SEQUENCENUMBER"><xsl:value-of select="SequenceNumber" /></xsl:attribute>
												</xsl:if>												
												<xsl:attribute name="DECLINEDDETAILS"><xsl:value-of select="RefusedMortgageDetails" /></xsl:attribute>
											</xsl:element>
										</xsl:for-each>											
									</xsl:when>
								</xsl:choose>	
																
								<xsl:for-each select="TelephoneNumbers/PhoneNumber">
									<xsl:call-template name="CreateTelephone">
										<xsl:with-param name="Label">CUSTOMERTELEPHONENUMBER</xsl:with-param><!--creates types 1,2,3-->
									</xsl:call-template>
								</xsl:for-each>
								<xsl:for-each select="Addresses/Address">
									<xsl:element name="CUSTOMERADDRESS">
										<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
										<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
										<xsl:if test="CustomerAddressSequenceNumber!='' and CustomerAddressSequenceNumber!='0'">
											<xsl:attribute name="CUSTOMERADDRESSSEQUENCENUMBER"><xsl:value-of select="CustomerAddressSequenceNumber"/></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="ADDRESSTYPE"><xsl:value-of select="AddressType"/></xsl:attribute>

										<xsl:call-template name="createFormattedDate">
											<xsl:with-param name="label" select="'DATEMOVEDIN'" />
											<xsl:with-param name="date" select="ResidentFrom" />
										</xsl:call-template>
										<xsl:call-template name="createFormattedDate">
											<xsl:with-param name="label" select="'DATEMOVEDOUT'" />
											<xsl:with-param name="date" select="ResidentTo" />
										</xsl:call-template>

										<!-- ADDRESS -->							
										<xsl:call-template name="CreateAddress"/>
										
										<xsl:if test="NatureOfOccupancy!='' and NatureOfOccupancy!='0'">
											<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="NatureOfOccupancy"/></xsl:attribute>
											<!--<xsl:if test="(NatureOfOccupancy='10' or NatureOfOccupancy='11') and (Landlord/Name!='' and position()=1)">-->
											<xsl:if test="(NatureOfOccupancy='10' or NatureOfOccupancy='11') and (Landlord/Name!='')">
												<xsl:element	name="TENANCY">													
													<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
													<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>																		
													<xsl:if test="CustomerAddressSequenceNumber!='' and CustomerAddressSequenceNumber!='0'">
														<xsl:attribute name="CUSTOMERADDRESSSEQUENCENUMBER"><xsl:value-of select="CustomerAddressSequenceNumber" /></xsl:attribute>
													</xsl:if>
													<xsl:if test="Landlord/MonthlyPayment!=''">
														<xsl:attribute name="MONTHLYRENTAMOUNT"><xsl:value-of select="Landlord/MonthlyPayment" /></xsl:attribute>
													</xsl:if>
													<xsl:if test="Landlord/AccountNumber!=''">
														<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="Landlord/AccountNumber" /></xsl:attribute>
													</xsl:if>
													<xsl:element name="THIRDPARTY">
														<xsl:if test="Landlord/ThirdPartyGuid!=''">
															<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="Landlord/ThirdPartyGuid"/></xsl:attribute>
														</xsl:if>
														<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'8'"/></xsl:attribute>
														<xsl:attribute name="COMPANYNAME"><xsl:value-of select="Landlord/Name"/></xsl:attribute>

														<!-- Landlord Address  -->
														<xsl:for-each select="Landlord/Address">
															<xsl:call-template name="CreateAddress" />
														</xsl:for-each>

														<!-- Only output CONTACTDETAILS if relevant data exists -->															
														<xsl:if test="Landlord/EmailAddress!='' or Landlord/TelephoneNumbers/PhoneNumber">
															<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
																<xsl:if test="Landlord/ContactDetailsGuid!=''">
																	<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="Landlord/ContactDetailsGuid"/></xsl:attribute>
																</xsl:if>
																<xsl:attribute name="EMAILADDRESS">																
																	<xsl:choose>
																		<xsl:when test="Landlord/EmailAddress!=''"><xsl:value-of select="Landlord/EmailAddress"/></xsl:when>
																		<xsl:otherwise> </xsl:otherwise>
																	</xsl:choose>
																</xsl:attribute>
																<xsl:for-each select="Landlord/TelephoneNumbers/PhoneNumber">
																	<xsl:call-template name="CreateTelephone">
																		<xsl:with-param name="Label"><xsl:value-of select="'CONTACTTELEPHONEDETAILS'" /></xsl:with-param>
																		<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
																	</xsl:call-template>
																</xsl:for-each>
															</xsl:element>
														</xsl:if>
													</xsl:element>
												</xsl:element>
											</xsl:if>
										</xsl:if>
									</xsl:element>
									<!-- End CUSTOMERADDRESS-->
															
								</xsl:for-each>
								
								<xsl:for-each select="CorrespondenceAddress">
									<xsl:element name="CUSTOMERADDRESS">
										<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
										<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
										<xsl:if test="CustomerAddressSequenceNumber!='' and CustomerAddressSequenceNumber!='0'">
											<xsl:attribute name="CUSTOMERADDRESSSEQUENCENUMBER"><xsl:value-of select="CustomerAddressSequenceNumber"/></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="ADDRESSTYPE"><xsl:value-of select="'2'"/></xsl:attribute>
										<!-- ADDRESS -->							
										<xsl:call-template name="CreateAddress">
											<xsl:with-param name="namespace" select="'http://OmigaData.Omiga.vertex.co.uk'" />										
										</xsl:call-template>
									</xsl:element>
								</xsl:for-each>
								
								<xsl:if test="Income/PayUKTax/IsNull!='true'">								
									<xsl:element name="INCOMESUMMARY">
										<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="CustomerNumber"/></xsl:attribute>
										<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="CustomerVersionNumber"/></xsl:attribute>
										<xsl:attribute name="UKTAXPAYERINDICATOR">
											<xsl:choose>
												<xsl:when test="Income/PayUKTax/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
												<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
											</xsl:choose> 										
										</xsl:attribute>
										<xsl:if test="TaxOfficeName!=''">
											<xsl:attribute name="TAXOFFICE"><xsl:value-of select="TaxOfficeName"/></xsl:attribute>
										</xsl:if>
										<xsl:if test="TaxReferenceNumber!=''">
											<xsl:attribute name="TAXREFERENCENUMBER"><xsl:value-of select="TaxReferenceNumber"/></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="SELFASSESS">
											<xsl:choose>
												<xsl:when test="SelfAssess='true'"><xsl:value-of select="'1'"/></xsl:when>
												<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
											</xsl:choose> 
										</xsl:attribute>
									</xsl:element>
								</xsl:if>
								
								<xsl:for-each select="$Employment">
									
									<xsl:variable name="CurrEmployment" select="." />
									
									<!-- EMPLOYMENT -->
									<xsl:if test="EmployerName!=''">
										<xsl:element name="EMPLOYMENT">
											<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
											<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
											<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
												<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
											</xsl:if>
											<!-- ik_ep2_2146 -->
											<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="EmploymentStatus"/></xsl:attribute>
											<xsl:attribute name="JOBTITLE"><xsl:value-of select="JobTitle"/></xsl:attribute>
											<xsl:call-template name="createFormattedDate">
												<xsl:with-param name="label" select="'DATESTARTEDORESTABLISHED'" />
												<xsl:with-param name="date" select="DateJobStarted" />
											</xsl:call-template>
											<xsl:choose>
												<xsl:when test="DateJobEnded!='' and substring(DateJobEnded, 1, 4)!='0001'">
													<xsl:call-template name="createFormattedDate">
														<xsl:with-param name="label" select="'DATELEFTORCEASEDTRADING'" />
														<xsl:with-param name="date" select="DateJobEnded" />
													</xsl:call-template>											
													<xsl:attribute name="MAINSTATUS"><xsl:value-of select="'0'" /></xsl:attribute><!-- Has an end date so not main employment -->
												</xsl:when>
												<xsl:otherwise>
													<xsl:attribute name="MAINSTATUS"><xsl:value-of select="'1'" /></xsl:attribute>
												</xsl:otherwise>
											</xsl:choose>
											<xsl:if test="EmploymentType!='' and EmploymentType!='0'">
												<xsl:attribute name="EMPLOYMENTTYPE"><xsl:value-of select="EmploymentType"/></xsl:attribute>
											</xsl:if>
											<!-- EP2_1568 -->
											<xsl:if test="NatureOfBusiness!='' and NatureOfBusiness!='0'">
												<xsl:attribute name="INDUSTRYTYPE"><xsl:value-of select="NatureOfBusiness"/></xsl:attribute>
											</xsl:if>
											<xsl:attribute name="SHARESOWNEDINDICATOR">
												<xsl:choose>
													<xsl:when test="EmploymentType!='10' and PercentShare!='' and PercentShare >= '5'"><xsl:value-of select="1"/></xsl:when>
													<xsl:otherwise><xsl:value-of select="0"/></xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
											<!-- EP2_1568 End -->
											
											<!-- NETPROFIT -->
											<xsl:if test="../../Income/SelfEmploymentIncome!='' and ../../Income/SelfEmploymentIncome!=''">
												<xsl:element name="NETPROFIT">
													<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
													<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
													<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
														<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
													</xsl:if>
													<xsl:attribute name="YEAR1AMOUNT"><xsl:value-of select="../../Income/SelfEmploymentIncome" /></xsl:attribute>
													<xsl:attribute name="YEAR1"><xsl:value-of select="substring($SysDateNow, 1, 4)" /></xsl:attribute>
												</xsl:element>
											</xsl:if>
											
											<!-- THIRDPARTY -->
											<xsl:if test="EmployerName!=''">
												<xsl:element name="THIRDPARTY">
													<xsl:if test="EmployerThirdPartyGuid!=''">
														<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="EmployerThirdPartyGuid"/></xsl:attribute>
													</xsl:if>
													<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'5'"/></xsl:attribute>
													<xsl:attribute name="COMPANYNAME"><xsl:value-of select="EmployerName"/></xsl:attribute>
													<xsl:if test="NatureOfBusiness!='' and NatureOfBusiness!='0'">
														<xsl:attribute name="ORGANISATIONTYPE"><xsl:value-of select="NatureOfBusiness"/></xsl:attribute>
													</xsl:if>
													<!-- ADDRESS -->
													<xsl:for-each select="EmployerAddress"><!-- Using For Each so context node changes for the CreateAddress template -->
														<xsl:call-template name="CreateAddress"/>
													</xsl:for-each>
													
													<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
														<xsl:if test="EmployerEmail">
															<xsl:if test="ContactDetailsGuid!=''">
																<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="ContactDetailsGuid"/></xsl:attribute>
															</xsl:if>
															<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="EmployerEmail"/></xsl:attribute>
														</xsl:if>
														<xsl:for-each select="EmployerTelephones/PhoneNumber">
															<xsl:call-template name="CreateTelephone">
																<xsl:with-param name="Label">CONTACTTELEPHONEDETAILS</xsl:with-param> <!-- type 5-->
																<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
															</xsl:call-template>
														</xsl:for-each>
													</xsl:element>
												</xsl:element>
											</xsl:if>
											
											<!-- SELFEMPLOYEDDETAILS -->	
											<xsl:element name="SELFEMPLOYEDDETAILS">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
												<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
													<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
												</xsl:if>
												<xsl:call-template name="createFormattedDate">
													<xsl:with-param name="label" select="'DATEFINANCIALINTERESTHELD'" />
													<xsl:with-param name="date" select="DateOfFinacialInterest" />
												</xsl:call-template>
												<xsl:if test="CompanyStatus!=''">
													<xsl:attribute name="COMPANYSTATUSTYPE"><xsl:value-of select="CompanyStatus"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="VatNumber!=''">
													<xsl:attribute name="VATNUMBER"><xsl:value-of select="VatNumber"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="CompanyRegNumber!=''">
													<xsl:attribute name="REGISTRATIONNUMBER"><xsl:value-of select="CompanyRegNumber"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="PercentShare!=''">
													<xsl:attribute name="PERCENTSHARESHELD"><xsl:value-of select="PercentShare"/></xsl:attribute>
												</xsl:if>
											</xsl:element>
	
											<!-- EMPLOYEDDETAILS-->
											<!-- These boolean values are changing into Null boolean values -->
											<xsl:element name="EMPLOYEDDETAILS">
												<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
												<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
												<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
													<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
												</xsl:if>
												<!-- EP2_1568 -->
												<xsl:if test="PercentShare!=''">
													<xsl:attribute name="PERCENTSHARESHELD"><xsl:value-of select="PercentShare"/></xsl:attribute>
												</xsl:if>
												<!-- EP2_1568 End -->
												<xsl:if test="ProbationPeriod/IsNull='false'">
													<xsl:attribute name="PROBATIONARYINDICATOR">
														<xsl:choose>
															<xsl:when test="ProbationPeriod/Value='true'"><xsl:value-of select="1"/></xsl:when>
															<xsl:otherwise><xsl:value-of select="'0'"/></xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
												</xsl:if>											
												<xsl:if test="UnderNotice/IsNull='false'">
													<xsl:attribute name="NOTICEPROBLEMINDICATOR">
														<xsl:choose>
															<xsl:when test="UnderNotice/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
															<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
												</xsl:if>
												<xsl:if test="EmployeeNumber!=''">
													<xsl:attribute name="PAYROLLNUMBER"><xsl:value-of select="EmployeeNumber"/></xsl:attribute>
												</xsl:if>
												<xsl:if test="FamilyOwnedCompany/IsNull='false'">
													<xsl:attribute name="EMPLOYMENTRELATIONSHIPIND">
														<xsl:choose>
															<xsl:when test="FamilyOwnedCompany/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
															<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
														</xsl:choose>
													</xsl:attribute>
												</xsl:if>
											</xsl:element>
											<!-- End EMPLOYEDDETAILS -->
											
											<!--<xsl:if test="Guid!='' or AccountantFirm!=''">-->
											<xsl:if test="HasAccountant='true'">
												<xsl:element name="ACCOUNTANT">
													<xsl:if test="AccountantGuid!=''">
														<xsl:attribute name="ACCOUNTANTGUID"><xsl:value-of select="AccountantGuid"/></xsl:attribute>
													</xsl:if>
													<xsl:attribute name="YEARSACTINGFORCUSTOMER"><xsl:value-of select="AccountantActingYears"/></xsl:attribute>
													<xsl:attribute name="ACCOUNTANCYFIRMNAME"><xsl:value-of select="AccountantFirm"/></xsl:attribute>
													<xsl:if test="AccountantQuals!='' and AccountantQuals!='0'">
														<xsl:attribute name="QUALIFICATIONS"><xsl:value-of select="AccountantQuals"/></xsl:attribute>
													</xsl:if>
													<xsl:if test="AccountantFirm!=''">
														<xsl:element name="THIRDPARTY">
															<xsl:if test="AccountantThirdPartyGuid!=''">
																<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="AccountantThirdPartyGuid"/></xsl:attribute>
															</xsl:if>
															<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'1'"/></xsl:attribute>
															<xsl:attribute name="COMPANYNAME"><xsl:value-of select="AccountantFirm" /></xsl:attribute> 
															
															<xsl:for-each select="AccountantAddress"><!-- Only ever one, but used to change context in CreateAddress template -->
																<xsl:call-template name="CreateAddress" />
															</xsl:for-each> 
			
															<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
																<xsl:if test="AccountantContactDetailsGuid!=''">
																	<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="AccountantContactDetailsGuid"/></xsl:attribute>
																</xsl:if>
																<xsl:attribute name="CONTACTFORENAME"><xsl:value-of select="AccountantFirstName"/></xsl:attribute>
																<xsl:attribute name="CONTACTSURNAME"><xsl:value-of select="AccountantLastName"/></xsl:attribute>
																<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="AccountantEmail"/></xsl:attribute>
																<xsl:for-each select="AccountantTelephones/PhoneNumber">
																	<xsl:call-template name="CreateTelephone">
																		<xsl:with-param name="Label">CONTACTTELEPHONEDETAILS</xsl:with-param> <!-- type 6-->
																		<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
																	</xsl:call-template>
																</xsl:for-each>
															</xsl:element>
														</xsl:element>											
													</xsl:if>
												</xsl:element>
											</xsl:if>
											
											<!-- CONTRACTDETAILS -->
											<xsl:if test="LengthOfContractMonths!='' and LengthOfContractYears!=''">
												<xsl:element name="CONTRACTDETAILS">
													<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
													<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
													<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
														<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
													</xsl:if>
													<xsl:attribute name="LENGTHOFCONTRACTYEARS"><xsl:value-of select="LengthOfContractYears"/></xsl:attribute>
													<xsl:attribute name="LENGTHOFCONTRACTMONTHS"><xsl:value-of select="LengthOfContractMonths"/></xsl:attribute>
													<xsl:if test="RenewableContract/IsNull='false'">
														<xsl:attribute name="LIKELYTOBERENEWED">
															<xsl:choose>
																<xsl:when test="RenewableContract/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
																<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
															</xsl:choose>
														</xsl:attribute>
													</xsl:if>
													<xsl:attribute name="NUMBEROFRENEWALS"><xsl:value-of select="NumTimesContractRenewed"/></xsl:attribute>
												</xsl:element>
											</xsl:if>
											
											<!-- 	EARNEDINCOME 
													Only exists for the Current Employment, so it doesnt matter that the EarnedIncome collection doesn't relate to an 
													Employment record in the structure.  
											-->
											<xsl:for-each select="$CustIncome/EarnedIncome/EarnedIncome">
												<xsl:if test="EarnedIncomeAmount!='' and EarnedIncomeAmount!='0' and ($CurrEmployment/DateJobEnded='' or substring($CurrEmployment/DateJobEnded, 1, 4)='0001')">
													<xsl:element name="EARNEDINCOME">
														<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
														<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
														<xsl:if test="EmploymentSequenceNumber!='' and EmploymentSequenceNumber!='0'">
															<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="EmploymentSequenceNumber"/></xsl:attribute>
														</xsl:if>
														<xsl:if test="EarnedIncomeSequenceNumber!='' and EarnedIncomeSequenceNumber!='0'">
															<xsl:attribute name="EARNEDINCOMESEQUENCENUMBER"><xsl:value-of select="EarnedIncomeSequenceNumber"/></xsl:attribute>
														</xsl:if>
														<xsl:attribute name="EARNEDINCOMETYPE"><xsl:value-of select="EarnedIncomeType"/></xsl:attribute>
														<xsl:attribute name="EARNEDINCOMEAMOUNT"><xsl:value-of select="EarnedIncomeAmount"/></xsl:attribute>
														<!-- All EarnedIncome amounts are annual so hard code RegularOutgoingsPaymentFreq combo value of 1 = Annually -->
														<xsl:attribute name="PAYMENTFREQUENCYTYPE"><xsl:value-of select="'1'" /></xsl:attribute>
													</xsl:element>
												</xsl:if>
											</xsl:for-each>
										</xsl:element>
									</xsl:if>
								</xsl:for-each>
								<xsl:if test="/Application/MarketingOptIn">
									<xsl:attribute name="MAILSHOTREQUIRED"><xsl:value-of select="/Application/MarketingOptIn"/></xsl:attribute>
								</xsl:if>
								
							</xsl:element>
							<!-- CustomerVersion End-->
						</xsl:element>
					</xsl:element>
				</xsl:for-each>
				<!-- New Property -->

				<xsl:element name="NEWPROPERTY">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
					<!-- EP2_1768 -->
					<xsl:if test="$PropertyDetail/OriginalValuerAddress/Postcode!=''">
						<xsl:attribute name="ORIGINALVALUERCOMPANYPOSTCODE"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/Postcode"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYFLATNONAME"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/FlatNumber"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/BuildingNumber"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/BuildingName"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYSTREET"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/Street"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYDISTRICT"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/District"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYTOWN"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/Town"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYCOUNTY"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/County"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYCOUNTRY"><xsl:value-of select="$PropertyDetail/OriginalValuerAddress/CountryType"/></xsl:attribute>
						<xsl:attribute name="ORIGINALVALUERCOMPANYNAME"><xsl:value-of select="$PropertyDetail/OriginalValuerCompanyName"/></xsl:attribute>
						<xsl:call-template name="createFormattedDate">
							<xsl:with-param name="label" select="'DATEOFORIGINALVALUATION'" />
							<xsl:with-param name="date" select="$PropertyDetail/OriginalValuationDate" />
						</xsl:call-template>
					</xsl:if>
					<!-- EP2_1768 End -->
					<xsl:if test="$PropertyDetail/privateSale/IsNull='false'">
						<xsl:attribute name="FAMILYSALEINDICATOR">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/privateSale/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>			
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/VendorRelated/IsNull='false'">
						<xsl:attribute name="BUSINESSCONNECTION">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/VendorRelated/Value='true'"><xsl:value-of select="'1'"/></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>	
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/FullMarketValue/IsNull='false'">
						<xsl:attribute name="FULLMARKETVALUE">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/FullMarketValue/Value"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'"/></xsl:otherwise>
							</xsl:choose>	
						</xsl:attribute>
					</xsl:if>
					<xsl:call-template name="createFormattedDate">
						<xsl:with-param name="label" select="'DATEOFENTRY'" />
						<xsl:with-param name="date" select="$PropertyDetail/EntryDate" />
					</xsl:call-template>
					<xsl:if test="$PropertyDetail/OtherFinancialIncentives/IsNull='false'">
						<xsl:attribute name="FINANCIALINCENTIVES">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/OtherFinancialIncentives/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:call-template name="createFormattedDate">
						<xsl:with-param name="label" select="'PREEMPTIONENDDATE'" />
						<xsl:with-param name="date" select="$PropertyDetail/PreemptionDate" />
					</xsl:call-template>
					<xsl:if test="$PropertyDetail/Tenure!='' and $PropertyDetail/Tenure!='0'">
						<xsl:attribute name="TENURETYPE"><xsl:value-of select="$PropertyDetail/Tenure"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/YearOfConstruction!='' and $PropertyDetail/YearOfConstruction!='0'">
						<xsl:attribute name="YEARBUILT"><xsl:value-of select="$PropertyDetail/YearOfConstruction"/></xsl:attribute>					
					</xsl:if>
					<xsl:if test="$PropertyDetail/NewBuild/IsNull='false'">
						<xsl:attribute name="NEWPROPERTYINDICATOR">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/NewBuild/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/HouseBuildersGuarantee!='' and $PropertyDetail/HouseBuildersGuarantee!='0'">
						<xsl:attribute name="HOUSEBUILDERSGUARANTEE"><xsl:value-of select="$PropertyDetail/HouseBuildersGuarantee"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/TypeOfProperty!='' and $PropertyDetail/TypeOfProperty!='0'">
						<xsl:attribute name="TYPEOFPROPERTY"><xsl:value-of select="$PropertyDetail/TypeOfProperty"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/DetachmentType!='' and $PropertyDetail/DetachmentType!='0'">
						<xsl:attribute name="DESCRIPTIONOFPROPERTY"><xsl:value-of select="$PropertyDetail/DetachmentType"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/FlatNumberOfFloors!=''">
						<xsl:attribute name="NUMBEROFSTOREYS"><xsl:value-of select="$PropertyDetail/FlatNumberOfFloors"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="NOFLATSINBLOCK"><xsl:value-of select="$PropertyDetail/FlatNumberOfFlats"/></xsl:attribute>
					<xsl:attribute name="FLOORNUMBER"><xsl:value-of select="$PropertyDetail/FlatWhichFloor"/></xsl:attribute>
					<xsl:attribute name="NUMBEROFGARAGES"><xsl:value-of select="$PropertyDetail/NumGarages" /></xsl:attribute>
					<xsl:attribute name="NUMBEROFOUTBUILDINGS"><xsl:value-of select="$PropertyDetail/NumOutBuildings" /></xsl:attribute>

					<xsl:if test="$PropertyDetail/FlatAboveCommercialPremises/IsNull='false'">
						<xsl:attribute name="FLATABOVECOMMERCIAL">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/FlatAboveCommercialPremises/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/BuildingConstructionType!=''">
						<xsl:attribute name="BUILDINGCONSTRUCTIONTYPE"><xsl:value-of select="$PropertyDetail/BuildingConstructionType" /></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/TotalLand!='' and $PropertyDetail/TotalLand!='0'">
						<xsl:attribute name="TOTALLANDTYPE"><xsl:value-of select="$PropertyDetail/TotalLand"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/RoofConstructionType!=''">
						<xsl:attribute name="ROOFCONSTRUCTIONTYPE"><xsl:value-of select="$PropertyDetail/RoofConstructionType" /></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/AgriculturalRestriction/IsNull='false'">
						<xsl:attribute name="AGRICULTURALRESTRICTIONS">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/AgriculturalRestriction/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/BusinessPurpose/IsNull='false'">
						<xsl:attribute name="RESIDENTIALUSEONLYINDICATOR">
							<xsl:choose>
								<!-- If Business is true then not Residential use only -->
								<xsl:when test="$PropertyDetail/BusinessPurpose/Value='true'"><xsl:value-of select="'0'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'1'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/LocalAuthorityWithinThreeYears/IsNull='false'">
						<xsl:attribute name="EXLOCALAUTHORITYINDICATOR">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/LocalAuthorityWithinThreeYears/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/AnyOtherPurpose/IsNull='false'">
						<xsl:attribute name="LETORPARTLETINDICATOR">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/AnyOtherPurpose/Value='true'"><xsl:value-of select="'1'"/></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDetail/AnyOtherOccupiers">
						<xsl:attribute name="ANYOTHERRESIDENTSINDICATOR"><xsl:value-of select="1"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="$PropertyDetail/BuyToLetIncome"/></xsl:attribute>
					<xsl:if test="$PropertyDetail/WillRelativeOccupy/IsNull='false'">
						<xsl:attribute name="REGULATIONINDICATOR">
							<xsl:choose>
								<xsl:when test="$PropertyDetail/WillRelativeOccupy='true'"><xsl:value-of select="'1'" /></xsl:when>
								<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>

					<xsl:attribute name="VALUATIONTYPE"><xsl:value-of select="$PropertyDetail/ValuationType"/></xsl:attribute>
					
					<!-- New Property Vendor -->
					<xsl:if test="$PropertyDetail/VendorFirstname!='' and $PropertyDetail/VendorSurname!='' ">
						<xsl:element name="NEWPROPERTYVENDOR">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo" /></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo" /></xsl:attribute>
							<xsl:element name="THIRDPARTY">
								<xsl:if test="$PropertyDetail/VendorThirdPartyGuid!=''">
									<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="$PropertyDetail/VendorThirdPartyGuid"/></xsl:attribute>
								</xsl:if>
								<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'12'"/></xsl:attribute>
								<xsl:attribute name="COMPANYNAME"><xsl:value-of select="concat($PropertyDetail/VendorFirstname, ' ', $PropertyDetail/VendorSurname)"/></xsl:attribute>

								<xsl:for-each select="$PropertyDetail/VendorAddress">
									<xsl:call-template name="CreateAddress" />
								</xsl:for-each>
								
								<!-- CONTACTDETAILS -->
								<xsl:if test="$PropertyDetail/VendorTelArea!='' or $PropertyDetail/VendorTelNumber!='' or $PropertyDetail/VendorTelExt!='' or $PropertyDetail/VendorFirstname!='' or $PropertyDetail/VendorSurname!=''">
									<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
										<xsl:if test="$PropertyDetail/VendorContactDetailsGuid!=''">
											<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="$PropertyDetail/VendorContactDetailsGuid" /></xsl:attribute>
										</xsl:if>
										<xsl:attribute name="CONTACTFORENAME"><xsl:value-of select="$PropertyDetail/VendorFirstname" /></xsl:attribute>	
										<xsl:attribute name="CONTACTSURNAME"><xsl:value-of select="$PropertyDetail/VendorSurname" /></xsl:attribute>
										
										<xsl:element name="CONTACTTELEPHONEDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
										   <xsl:if test="$PropertyDetail/VendorContactDetailsGuid!=''">
											   <xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="$PropertyDetail/VendorContactDetailsGuid"/></xsl:attribute>
											</xsl:if>
											<xsl:attribute name="TELEPHONESEQNUM"><xsl:value-of select="'1'"/></xsl:attribute>
											<xsl:attribute name="TELENUMBER"><xsl:value-of select="$PropertyDetail/VendorTelNumber"/></xsl:attribute>
											<xsl:attribute name="AREACODE"><xsl:value-of select="$PropertyDetail/VendorTelArea"/></xsl:attribute>				
											<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="$PropertyDetail/VendorTelExt"/></xsl:attribute>
											<xsl:attribute name="USAGE">30</xsl:attribute>
										</xsl:element>																											
									</xsl:element>							
								</xsl:if>

							</xsl:element>
						</xsl:element>
					</xsl:if>
					
					<!-- NEWPROPERTYADDRESS -->
					<xsl:if test="$PropertyDetail/PropertyAddress/Postcode!=''">
						<xsl:element name="NEWPROPERTYADDRESS">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
							<xsl:if test="$PropertyDetail/PropertyAddress/AddressGuid!=''">
								<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="$PropertyDetail/PropertyAddress/AddressGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="$PropertyDetail/ValuationContact!='' and $PropertyDetail/ValuationContact!='0'">
								<xsl:attribute name="ARRANGEMENTSFORACCESS"><xsl:value-of select="$PropertyDetail/ValuationContact"/></xsl:attribute>
							</xsl:if>
							<!-- ADDRESS -->
							<xsl:for-each select="$PropertyDetail/PropertyAddress[1]">
								<xsl:call-template name="CreateAddress">
									<xsl:with-param name="namespace" select="'http://OmigaData.Omiga.vertex.co.uk'" />
								</xsl:call-template>
							</xsl:for-each>
						</xsl:element>
					</xsl:if>
					
					<!-- NEWPROPERTYROOMTYPE -->
					<xsl:for-each select="$PropertyDetail/NewPropertyRoomTypes/NewPropertyRoomType">
						<xsl:element name="NEWPROPERTYROOMTYPE">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
							<xsl:attribute name="ROOMTYPE"><xsl:value-of select="RoomType"/></xsl:attribute>
							<xsl:attribute name="NUMBEROFROOMS"><xsl:value-of select="NumberOfRooms"/></xsl:attribute>
						</xsl:element>
					</xsl:for-each>					
					
					<!-- NEWPROPERTYLEASEHOLD -->
					<xsl:if test="$PropertyDetail/RemainingTermOfLease!='' and $PropertyDetail/RemainingTermOfLease!='0'">
						<xsl:element name="NEWPROPERTYLEASEHOLD">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
							<xsl:attribute name="UNEXPIREDTERMOFLEASEYEARS"><xsl:value-of select="$PropertyDetail/RemainingTermOfLease"/></xsl:attribute>
						</xsl:element>
					</xsl:if>					
				</xsl:element><!-- End NEWPROPERTY -->				

				<!-- APPLICATIONESTATEAGENT -->
				<xsl:if test="$PropertyDetail/SellingAgentName!=''">
					<xsl:element name="APPLICATIONESTATEAGENT">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo" /></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo" /></xsl:attribute>
						<xsl:element name="THIRDPARTY">
							<xsl:if test="$PropertyDetail/SellingAgentThirdPartyGuid">
								<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="$PropertyDetail/SellingAgentThirdPartyGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'6'"/></xsl:attribute>
							<xsl:attribute name="COMPANYNAME"><xsl:value-of select="$PropertyDetail/SellingAgentName"/></xsl:attribute>

							<xsl:if test="$PropertyDetail/SellingAgentTelephone/PhoneNumber/Number!=''">
								<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
									<xsl:if test="$PropertyDetail/SellingAgentContactDetailsGuid!=''">
										<xsl:attribute name="CONTACTDETAILSGUID">
											<xsl:value-of select="$PropertyDetail/SellingAgentContactDetailsGuid" />
										</xsl:attribute>
									</xsl:if>
									
									<xsl:for-each select="$PropertyDetail/SellingAgentTelephone/PhoneNumber">
										<xsl:call-template name="CreateTelephone">
											<xsl:with-param name="Label"><xsl:value-of select="'CONTACTTELEPHONEDETAILS'" /></xsl:with-param>
											<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
										</xsl:call-template>
									</xsl:for-each>
								</xsl:element>
							</xsl:if>

						</xsl:element>
					</xsl:element>
				</xsl:if>
				
				<xsl:variable name="Deposits" select="/Application/DepositSources/DepositSource"/>
				<xsl:for-each select="$Deposits">
					<xsl:element name="NEWPROPERTYDEPOSIT">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
						<xsl:if test="DepositSequenceNumber!='' and DepositSequenceNumber!='0'">
							<xsl:attribute name="NPDEPOSITSEQUENCENUMBER"><xsl:value-of select="DepositSequenceNumber"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="SOURCEOFFUNDING"><xsl:value-of select="DepositSourceType"/></xsl:attribute>
						<xsl:attribute name="AMOUNT"><xsl:value-of select="Amount"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>

				<xsl:for-each select="$PropertyDetail/OtherOccupiers/OtherOccupier">
					<xsl:element name="OTHERRESIDENT">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
						<xsl:if test="OtherResidentSequenceNumber!='' and OtherResidentSequenceNumber!='0'">
							<xsl:attribute name="OTHERRESIDENTSEQUENCENUMBER"><xsl:value-of select="OtherResidentSequenceNumber" /></xsl:attribute>
						</xsl:if>
						
						<xsl:element name="PERSON">
							<xsl:if test="PersonGuid!=''">
								<xsl:attribute name="PERSONGUID"><xsl:value-of select="PersonGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="FIRSTFORENAME"><xsl:value-of select="FirstName"/></xsl:attribute>
							<xsl:attribute name="SURNAME"><xsl:value-of select="LastName"/></xsl:attribute>
							<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="DateOfBirth"/></xsl:attribute>
							<xsl:call-template name="createFormattedDate">
								<xsl:with-param name="label" select="'DATEOFBIRTH'" />
								<xsl:with-param name="date" select="DateOfBirth" />
							</xsl:call-template>							
							<xsl:attribute name="RELATIONSHIPTOAPPLICANT"><xsl:value-of select="Relationship"/></xsl:attribute>
						</xsl:element>
					</xsl:element>
				</xsl:for-each>

				<!-- Financial Summary-->
				<xsl:element name="FINANCIALSUMMARY">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
					<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
					
					<xsl:if test="$PaymentHist/BankruptInLastSixYears/IsNull='false'">
						<xsl:attribute name="BANKRUPTCYHISTORYINDICATOR">
							<xsl:choose>
								<xsl:when test="$PaymentHist/BankruptInLastSixYears/Value='true'">1</xsl:when>
								<xsl:otherwise>0</xsl:otherwise> 
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					
					<xsl:if test="$PaymentHist/Ccj/IsNull='false'">
						<xsl:attribute name="CCJHISTORYINDICATOR">
							<xsl:choose>
								<xsl:when test="$PaymentHist/Ccj/Value='true' ">1</xsl:when>
								<xsl:otherwise>0</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					
					<xsl:if test="$PaymentHist/RefusedMortgage/IsNull='false'">
						<xsl:attribute name="DECLINEDMORTGAGEINDICATOR">
							<xsl:choose>
								<xsl:when test="$PaymentHist/RefusedMortgage/Value='true'">1</xsl:when>
								<xsl:otherwise>0</xsl:otherwise>
							</xsl:choose>
						</xsl:attribute>
					</xsl:if>
					
					<!-- test if ArrearsIndicator should be set -->
					<xsl:variable name="HasLandlordArrears">
						<xsl:choose>
							<xsl:when test="count(//Application/Applicants/Applicant/Addresses/Address/Landlord[InArrears='true'])>0">true</xsl:when>
							<xsl:otherwise>false</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<xsl:variable name="HasCommitmentArrears">
						<xsl:choose>
							<xsl:when test="count(//Application/Commitments/Commitment[MonthsInArrears!='' and MonthsInArrears!='0'])>0">true</xsl:when>
							<xsl:otherwise>false</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>

					<xsl:if test="$HasLandlordArrears='true' or $HasCommitmentArrears='true'">
						<xsl:attribute name="ARREARSHISTORYINDICATOR">1</xsl:attribute>
					</xsl:if> 
					
				</xsl:element>
					
				<!-- BANKRUPTCYHISTORY -->
				<xsl:if test="$PaymentHist/BankruptInLastSixYears/Value='true'">
					<xsl:for-each select="$PaymentHist/Bankruptcies/Bankruptcy">
						<xsl:element name="BANKRUPTCYHISTORY">
							<xsl:if test="BankruptcyGuid!=''">
								<xsl:attribute name="BANKRUPTCYHISTORYGUID"><xsl:value-of select="BankruptcyGuid" /></xsl:attribute>
							</xsl:if>
							<xsl:if test="DateOfDischarge!='' and substring(DateOfDischarge, 1, 4)!='0001'">
								<xsl:call-template name="createFormattedDate">
									<xsl:with-param name="label" select="'DATEOFDISCHARGE'" />
									<xsl:with-param name="date" select="DateOfDischarge" />
								</xsl:call-template>								
							</xsl:if>
							<xsl:if test="AmountOfDebt!=''">
								<xsl:attribute name="AMOUNTOFDEBT"><xsl:value-of select="AmountOfDebt" /></xsl:attribute>
							</xsl:if> 
							<xsl:if test="MonthlyRepayment!=''">
								<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MonthlyRepayment" /></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="IVA">
								<xsl:choose>
									<xsl:when test="IVAInLastSixYears='true'">1</xsl:when>
									<xsl:otherwise>0</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:if test="IvaDate!='' and substring(IvaDate, 1, 4)!='0001'">
								<xsl:call-template name="createFormattedDate">
									<xsl:with-param name="label" select="'OTHERDETAILS'" />
									<xsl:with-param name="date" select="IvaDate" />
								</xsl:call-template>								
							</xsl:if>
						
							<!-- CUSTOMERVERSIONBANKRUPTCYHISTORY -->
							<xsl:if test="Applicant1='true'">
								<xsl:element name="CUSTOMERVERSIONBANKRUPTCYHISTORY">
									<xsl:if test="BankruptcyGuid!=''">
										<xsl:attribute name="BANKRUPTCYHISTORYGUID"><xsl:value-of select="BankruptcyGuid" /></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber1" /></xsl:attribute>
									<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber1" /></xsl:attribute>
								</xsl:element>
							</xsl:if>
							<xsl:if test="Applicant2='true'">
								<xsl:element name="CUSTOMERVERSIONBANKRUPTCYHISTORY">
									<xsl:if test="BankruptcyGuid!=''">
										<xsl:attribute name="BANKRUPTCYHISTORYGUID"><xsl:value-of select="BankruptcyGuid" /></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber2" /></xsl:attribute>
									<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber2" /></xsl:attribute>								
								</xsl:element>							
							</xsl:if>
							<xsl:if test="Guarantor='true'">
								<xsl:element name="CUSTOMERVERSIONBANKRUPTCYHISTORY">
									<xsl:if test="BankruptcyGuid!=''">
										<xsl:attribute name="BANKRUPTCYHISTORYGUID"><xsl:value-of select="BankruptcyGuid" /></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$GuarantorNumber" /></xsl:attribute>
									<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$GuarantorVersionNumber" /></xsl:attribute>																
								</xsl:element>
							</xsl:if>
						</xsl:element>
					</xsl:for-each>
				</xsl:if>

				<!--ACCOUNT / MORTGAGEACCOUNT -->
				<xsl:for-each select="/Application/Commitments/Commitment">
				
					<xsl:if test="IsMortgage='true'">
						
						<xsl:variable name="AccGuid" select="AccountGuid" />					
					
						<xsl:element name="ACCOUNT">
							<xsl:if test="AccountGuid!=''">
								<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="AccountNumber!=''">
								<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="AccountNumber"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="Applicant1='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber1"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber1"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="Applicant2='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber2"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber2"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="Guarantor='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$GuarantorNumber"/>
									<xsl:with-param name="CustomerVersionNumber" select="$GuarantorVersionNumber"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="2"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:element name="MORTGAGEACCOUNT">
								<xsl:if test="$AccGuid!=''">
									<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
								</xsl:if>
								<xsl:if test="MortgageAddress/AddressGuid!=''">
									<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="MortgageAddress/AddressGuid" /></xsl:attribute> 
								</xsl:if>
								<xsl:attribute name="TOTALCOLLATERALBALANCE"><xsl:value-of select="AmountOutstanding"/></xsl:attribute>
								<xsl:if test="RedemptionStatus!=''">
									<xsl:attribute name="REDEMPTIONSTATUS"><xsl:value-of select="RedemptionStatus"/></xsl:attribute>
								</xsl:if>
								<xsl:attribute name="MONTHLYRENTALINCOME"><xsl:value-of select="RentalIncome"/></xsl:attribute>
								<xsl:attribute name="OTHERPROPERTYMORTGAGE">
									<xsl:choose>
										<xsl:when test="OtherPropertyMortgage='true'"><xsl:value-of select="'1'" /></xsl:when>
										<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<xsl:attribute name="ADDITIONALDETAILS"><xsl:value-of select="AdditionalDetails"/></xsl:attribute>
						
								<xsl:if test="SecondCharge/IsNull='false'">
									<xsl:attribute name="SECONDCHARGEINDICATOR">
										<xsl:choose>
											<xsl:when test="SecondCharge/Value='true'"><xsl:value-of select="'1'" /></xsl:when>
											<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
										</xsl:choose>
									</xsl:attribute>
								</xsl:if>
								
								<xsl:choose>
									<xsl:when test="MortgageAddress/AddressGuid!=''"><!-- Do Nothing --></xsl:when>
									<xsl:otherwise>
										<!-- Only create Address if it doesn't exist in the database. i.e. doesn't have an AddressGuid-->
										<xsl:for-each select="MortgageAddress">
											<xsl:call-template name="CreateAddress">
												<xsl:with-param name="namespace" select="'http://OmigaData.Omiga.vertex.co.uk'" />
											</xsl:call-template>
										</xsl:for-each>
									</xsl:otherwise>
								</xsl:choose>
								
								<!-- IK EP2_2329 -->
								<xsl:attribute name="TOTALMONTHLYCOST"><xsl:value-of select="MonthlyPayment"/></xsl:attribute>
								
								<xsl:element name="MORTGAGELOAN">
									<xsl:if test="MortgageLoanGuid!=''">
										<xsl:attribute name="MORTGAGELOANGUID"><xsl:value-of select="MortgageLoanGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MonthlyPayment"/></xsl:attribute>
									<xsl:if test="RepaymentType!=''">
										<xsl:attribute name="REPAYMENTTYPE"><xsl:value-of select="RepaymentType"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="OUTSTANDINGBALANCE"><xsl:value-of select="AmountOutstanding"/></xsl:attribute>
									<xsl:call-template name="createFormattedDate">
										<xsl:with-param name="label" select="'STARTDATE'" />
										<xsl:with-param name="date" select="StartDate" />
									</xsl:call-template>
									<xsl:attribute name="ORIGINALLOANAMOUNT"><xsl:value-of select="OriginalAmount"/></xsl:attribute>
									<xsl:if test="RedemptionStatus!=''">
										<xsl:attribute name="REDEMPTIONSTATUS"><xsl:value-of select="RedemptionStatus"/></xsl:attribute>
									</xsl:if>

								</xsl:element>							
							</xsl:element>
							
							<xsl:if test="LenderName!=''">
								<xsl:element name="THIRDPARTY">
									<xsl:if test="LenderThirdPartyGuid!=''">
										<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="LenderThirdPartyGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="3"/></xsl:attribute>
									<xsl:attribute name="COMPANYNAME"><xsl:value-of select="LenderName"/></xsl:attribute>
	
									<xsl:if test="LenderTelephoneNumbers/PhoneNumber/Number!=''">
										<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
											
											<xsl:if test="LenderContactDetailsGuid!=''">
												<xsl:if test="LenderContactDetailsGuid!=''">
													<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="LenderContactDetailsGuid"/></xsl:attribute>
												</xsl:if>																			
											</xsl:if>
											<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="''"/></xsl:attribute><!-- Writes empty string to force CREATE CRUD_OP -->
											
											<xsl:for-each select="LenderTelephoneNumbers/PhoneNumber">
												<xsl:call-template name="CreateTelephone">
													<xsl:with-param name="Label"><xsl:value-of select="'CONTACTTELEPHONEDETAILS'" /></xsl:with-param>
													<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
												</xsl:call-template>
											</xsl:for-each>
											
										</xsl:element>
									</xsl:if>

									<xsl:for-each select="LenderAddress"><!-- Changes context -->
										<xsl:call-template name="CreateAddress" />										
									</xsl:for-each>									
									
								</xsl:element>
							</xsl:if>
							
							<xsl:if test="MonthsInArrears!='' and MonthsInArrears!='0'">
								<xsl:element name="ARREARSHISTORY">
									<xsl:if test="ArrearsHistoryGuid!=''">
										<xsl:attribute name="ARREARSHISTORYGUID"><xsl:value-of select="ArrearsHistoryGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:if test="$AccGuid!=''">
										<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="MAXIMUMNUMBEROFMONTHS"><xsl:value-of select="MonthsInArrears"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
						</xsl:element>
					</xsl:if>						
				</xsl:for-each>				
				
				<!--ACCOUNT / LOANSLIABILITIES  -->
				<xsl:for-each select="/Application/Commitments/Commitment">
					<xsl:if test="IsMortgage='false'">
						<xsl:variable name="AccGuid" select="AccountGuid"/>							
						<xsl:element name="ACCOUNT">
							<xsl:if test="$AccGuid!=''">
								<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="ADDITIONALINDICATOR"><xsl:value-of select="'0'" /></xsl:attribute>
							<xsl:if test="LenderThirdPartyGuid!=''">
								<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="LenderThirdPartyGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="AccountNumber"/></xsl:attribute>
							<xsl:if test="MonthsInArrears!='' and MonthsInArrears!='0'">
								<xsl:element name="ARREARSHISTORY">
									<xsl:if test="ArrearsHistoryGuid!=''">
										<xsl:attribute name="ARREARSHISTORYGUID"><xsl:value-of select="ArrearsHistoryGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:if test="$AccGuid!=''">
										<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="MAXIMUMNUMBEROFMONTHS"><xsl:value-of select="MonthsInArrears"/></xsl:attribute>
								</xsl:element>
							</xsl:if>						
							<xsl:element name="LOANSLIABILITIES">
								<xsl:if test="$AccGuid!=''">
									<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
								</xsl:if>
								<xsl:attribute name="AGREEMENTTYPE"><xsl:value-of select="CommitmentType"/></xsl:attribute>
								<xsl:attribute name="MONTHLYREPAYMENT"><xsl:value-of select="MonthlyPayment"/></xsl:attribute>

								<!-- EP2_2216 TOTALOUTSTANDINGBALANCE: only set on CREDIT CARD entries -->		
								<xsl:if test="IsCreditCard = 'true'">
									<xsl:attribute name="TOTALOUTSTANDINGBALANCE"><xsl:value-of select="AmountOutstanding"/></xsl:attribute>
								</xsl:if>
								<!-- EP2_2216 ends -->
								
								<xsl:call-template name="createFormattedDate">
									<xsl:with-param name="label" select="'ENDDATE'" />
									<xsl:with-param name="date" select="EndDate" />
								</xsl:call-template>					
																	
								<xsl:attribute name="LOANREPYMENTIDINDICATOR">
									<xsl:choose>
										<xsl:when test="IsRepaidOnOrBeforeCompletion='true'"><xsl:value-of select="'1'" /></xsl:when>
										<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>
								<xsl:attribute name="PAIDFORBYBUSINESS">
									<xsl:choose>
										<xsl:when test="IsBusinessCommitment='true'"><xsl:value-of select="'1'" /></xsl:when>
										<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
									</xsl:choose>
								</xsl:attribute>							
							</xsl:element>
							<xsl:if test="Applicant1='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber1"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber1"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="Applicant2='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber2"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber2"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="Guarantor='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$GuarantorNumber"/>
									<xsl:with-param name="CustomerVersionNumber" select="$GuarantorVersionNumber"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="2"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="LenderName!=''">
								<xsl:element name="THIRDPARTY">
									<xsl:if test="LenderThirdPartyGuid!=''">
										<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="LenderThirdPartyGuid"/></xsl:attribute>
									</xsl:if>
									<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'3'"/></xsl:attribute>
									<xsl:attribute name="COMPANYNAME"><xsl:value-of select="LenderName"/></xsl:attribute>
								</xsl:element>
							</xsl:if>						
						</xsl:element>
					</xsl:if>
				</xsl:for-each>
				
				<!-- LANDLORD ARREARSHISTORY -->
				<xsl:for-each select="//Application/Applicants/Applicant/Addresses/Address/Landlord[InArrears='true']">
				
					<xsl:variable name="AccGuid" select="ArrearsHistoryAccountGuid" />
					
					<xsl:if test="NumberOfMonthsInArrears!='' and NumberOfMonthsInArrears!='0'">
						<xsl:element name="ACCOUNT">
							<xsl:if test="$AccGuid!=''">
								<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
							</xsl:if>
							<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="AccountNumber" /></xsl:attribute>
							<!-- ik_wip -->
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="../../../CustomerNumber" /></xsl:attribute>
							<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="../../../CustomerVersionNumber" /></xsl:attribute>
							<xsl:attribute name="CUSTOMERADDRESSSEQUENCENUMBER"><xsl:value-of select="../CustomerAddressSequenceNumber" /></xsl:attribute>
							<!--<xsl:if test="ThirdPartyGuid!=''">
								<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="ThirdPartyGuid"/></xsl:attribute>
							</xsl:if>-->
							<xsl:element name="ARREARSHISTORY">
								<xsl:if test="ArrearsHistoryGuid!=''">
									<xsl:attribute name="ARREARSHISTORYGUID"><xsl:value-of select="ArrearsHistoryGuid"/></xsl:attribute>
								</xsl:if>
								<xsl:if test="$AccGuid!=''">
									<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
								</xsl:if>

								<!-- Hard coded for Landlord: ArrearsLoanType - Other = 4-->					
								<xsl:attribute name="DESCRIPTIONOFLOAN">4</xsl:attribute>
								<xsl:attribute name="MAXIMUMNUMBEROFMONTHS"><xsl:value-of select="NumberOfMonthsInArrears"/></xsl:attribute>
							</xsl:element>
						
							<xsl:if test="Applicant1Arrears='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber1"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber1"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="Applicant2Arrears='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$CustomerNumber2"/>
									<xsl:with-param name="CustomerVersionNumber" select="$CustomerVersionNumber2"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="1"/>
								</xsl:call-template>
							</xsl:if>
							<xsl:if test="GuarantorArrears='true'">
								<xsl:call-template name="CreateAccountRelationship">
									<xsl:with-param name="CustomerNumber" select="$GuarantorNumber"/>
									<xsl:with-param name="CustomerVersionNumber" select="$GuarantorVersionNumber"/>
									<xsl:with-param name="AccGuid" select="$AccGuid"/>
									<xsl:with-param name="CustomerRoleType" select="2"/>
								</xsl:call-template>
							</xsl:if>
							
						</xsl:element>
					</xsl:if>
				</xsl:for-each>
			
				<!-- APPLICATIONLEGALREP -->
				<xsl:if test="$Solicitor/CompanyName!=''">
					<xsl:element name="APPLICATIONLEGALREP">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo" /></xsl:attribute> 
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo" /></xsl:attribute> 
						
						<!-- THIRDPARTY -->
						<xsl:element name="THIRDPARTY">
							<xsl:if test="$Solicitor/ThirdPartyGuid!=''">
								<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="$Solicitor/ThirdPartyGuid"/></xsl:attribute>
							</xsl:if>								
							<xsl:if test="$Solicitor/DirectoryGuid!=''">
								<xsl:attribute name="DIRECTORYGUID"><xsl:value-of select="$Solicitor/DirectoryGuid" /></xsl:attribute>							
							</xsl:if>
							<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="10"/></xsl:attribute>
							<xsl:attribute name="COMPANYNAME"><xsl:value-of select="$Solicitor/CompanyName"/></xsl:attribute>

							<xsl:if test="$Solicitor/DxNumber!=''">
								<xsl:attribute name="DXID"><xsl:value-of select="$Solicitor/DxNumber"/></xsl:attribute>
							</xsl:if>
							<xsl:if test="$Solicitor/District!=''">
								<xsl:attribute name="DXLOCATION"><xsl:value-of select="$Solicitor/District" /></xsl:attribute>
							</xsl:if>							
							
							<!-- CONTACTDETAILS -->
							<xsl:element name="CONTACTDETAILS" namespace="http://msgtypes.Omiga.vertex.co.uk">
								<xsl:if test="$Solicitor/ContactDetailsGuid!=''">
									<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="$Solicitor/ContactDetailsGuid"/></xsl:attribute>
								</xsl:if>
								<xsl:attribute name="CONTACTFORENAME"><xsl:value-of select="$Solicitor/FirstName"/></xsl:attribute>
								<xsl:attribute name="CONTACTSURNAME"><xsl:value-of select="$Solicitor/LastName"/></xsl:attribute>
								<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="$Solicitor/EmailAddress"/></xsl:attribute>							

								<!-- CONTACTTELEPHONEDETAILS -->
								<xsl:for-each select="$Solicitor/TelephoneNumbers/PhoneNumber">
									<xsl:call-template name="CreateTelephone">
										<xsl:with-param name="Label">CONTACTTELEPHONEDETAILS</xsl:with-param> <!-- type 99,12-->
										<xsl:with-param name="namespace"><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:with-param>
									</xsl:call-template>
								</xsl:for-each>									
							</xsl:element>

							<!-- ADDRESS -->
							<xsl:for-each select="$Solicitor/Address"><!-- Changes context to Address node although there is just one Address -->
								<xsl:call-template name="CreateAddress"/>
							</xsl:for-each>
						</xsl:element>
					</xsl:element>
				</xsl:if>
				
				<xsl:if test="/Application/DirectDebit/BankAccountNumber!=''">

					<xsl:variable name="DirectDebit" select="/Application/DirectDebit" />
					
					<xsl:element name="APPLICATIONBANKBUILDINGSOC">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
						<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo"/></xsl:attribute>
						<xsl:attribute name="BANKACCOUNTSEQUENCENUMBER"><xsl:value-of select="'1'"/></xsl:attribute>
						<xsl:attribute name="ACCOUNTNAME"><xsl:value-of select="$DirectDebit/AccountHolder"/></xsl:attribute>
						<xsl:attribute name="ACCOUNTNUMBER"><xsl:value-of select="$DirectDebit/BankAccountNumber"/></xsl:attribute>
						<xsl:attribute name="ROLLNUMBER"><xsl:value-of select="$DirectDebit/RollNumber"/></xsl:attribute>
						<xsl:if test="$DirectDebit/BankName != ''">
							<xsl:element name="THIRDPARTY">
								<xsl:if test="$DirectDebit/BankThirdPartyGuid">
									<xsl:attribute name="THIRDPARTYGUID"><xsl:value-of select="/Application/DirectDebit/BankThirdPartyGuid"/></xsl:attribute>
								</xsl:if>							
								<xsl:attribute name="THIRDPARTYTYPE"><xsl:value-of select="'3'"/></xsl:attribute>
								<xsl:attribute name="COMPANYNAME"><xsl:value-of select="$DirectDebit/BankName"/></xsl:attribute>
								<xsl:attribute name="THIRDPARTYBANKSORTCODE"><xsl:value-of select="concat($DirectDebit/BankSortCode1, '-', $DirectDebit/BankSortCode2, '-', $DirectDebit/BankSortCode3)" /></xsl:attribute>

								<!-- Bank Address  -->
								<xsl:for-each select="$DirectDebit/BankAddress">
									<xsl:call-template name="CreateAddress" />
								</xsl:for-each>								
							</xsl:element>
						</xsl:if>
					</xsl:element>
				</xsl:if>				
			</xsl:element>
			<!-- ApplicationFactFind - End-->

			<!-- Deserialization issue: Could be namespace related -->			
			<!-- PAYMENTRECORD -->
			<xsl:for-each select="$CCPayment[AmountPaid!='' and AmountPaid!='0']">
				<xsl:element name="PAYMENTRECORD">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
					<xsl:attribute name="PAYMENTSEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
					<xsl:attribute name="AMOUNT"><xsl:value-of select="AmountPaid"/></xsl:attribute>
					<xsl:call-template name="createFormattedDate">
						<xsl:with-param name="label" select="'CREATIONDATETIME'" />
						<xsl:with-param name="date" select="$SysDateNow" />
					</xsl:call-template>
				</xsl:element>
			</xsl:for-each>

			<!-- APPLICATIONFEETYPE -->
			<xsl:for-each select="$CCPayment[AmountPaid!='' and AmountPaid!='0']">
				<xsl:element name="APPLICATIONFEETYPE">
					<xsl:if test="$AppNo!=''">
						<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="FeeType!=''">
						<xsl:attribute name="FEETYPE"><xsl:value-of select="FeeType"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="AmountPaid!=''">
						<xsl:attribute name="AMOUNT"><xsl:value-of select="AmountPaid"/></xsl:attribute>
					</xsl:if>
					<xsl:element name="FEEPAYMENT">
						<xsl:if test="$AppNo!=''">
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="PAYMENTSEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
						<xsl:if test="FeeType!=''">
							<xsl:attribute name="FEETYPE"><xsl:value-of select="FeeType"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="FEETYPESEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:attribute>
						<xsl:if test="AmountPaid!=''">
							<xsl:attribute name="AMOUNTPAID"><xsl:value-of select="AmountPaid"/></xsl:attribute>
						</xsl:if>
						<xsl:attribute name="PAYMENTEVENT"><xsl:value-of select="10"/></xsl:attribute>
						<xsl:if test="RefundAmount!='' and RefundAmount!='0'">
							<xsl:attribute name="REFUNDAMOUT"><xsl:value-of select="RefundAmount"/></xsl:attribute>
						</xsl:if>
					</xsl:element>
				</xsl:element>
			</xsl:for-each>
			
		</xsl:element>
		<!-- Application - End-->
	</xsl:template>

	<xsl:template name="CreateAddress">

		<xsl:param name="namespace" />
		
		<xsl:variable name="resolvedNamespace">
			<xsl:choose>
				<xsl:when test="$namespace!=''"><xsl:value-of select="$namespace" /></xsl:when>
				<xsl:otherwise><xsl:value-of select="'http://msgtypes.Omiga.vertex.co.uk'" /></xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		
		<xsl:if test="AddressGuid!='' or BuildingNumber!='' or BuildingName!='' or FlatNumber!='' or Street!='' or District!='' or Town!='' or County!='' or CountryType!='' or Postcode!=''">
			<!--<xsl:element name="ADDRESS" namespace="http://msgtypes.Omiga.vertex.co.uk">-->
			<xsl:element name="ADDRESS" namespace="{string($resolvedNamespace)}">
				<xsl:if test="AddressGuid!=''">
					<xsl:attribute name="ADDRESSGUID"><xsl:value-of select="AddressGuid"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="BuildingName!=''">
					<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="BuildingName"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="BuildingNumber!=''">
					<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="BuildingNumber"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="FlatNumber!=''">
					<xsl:attribute name="FLATNUMBER"><xsl:value-of select="FlatNumber"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="Street!=''">
					<xsl:attribute name="STREET"><xsl:value-of select="Street"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="District!=''">
					<xsl:attribute name="DISTRICT"><xsl:value-of select="District"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="Town!=''">
					<xsl:attribute name="TOWN"><xsl:value-of select="Town"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="County!=''">
					<xsl:attribute name="COUNTY"><xsl:value-of select="County"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="CountryType!=''">
					<xsl:attribute name="COUNTRY"><xsl:value-of select="CountryType"/></xsl:attribute>
				</xsl:if>
				<xsl:if test="Postcode!=''">
					<xsl:attribute name="POSTCODE"><xsl:value-of select="Postcode"/></xsl:attribute>
				</xsl:if>			
			</xsl:element>
		</xsl:if>
	</xsl:template>
	
	<xsl:template name="CreateTelephone">
		<xsl:param name="Label"/>
		<xsl:param name="namespace" />
		
		<xsl:if test="Number!=''">
			<xsl:element name="{$Label}" namespace="{string($namespace)}">
	
				<!-- CUSTOMERTELEPHONENUMBER -->
			   <xsl:if test="$Label='CUSTOMERTELEPHONENUMBER'">
				   <xsl:if test="../../CustomerNumber!=''">
						<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="../../CustomerNumber"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="../../CustomerVersionNumber!='' and ../../CustomerVersionNumber!='0'">
						<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="../../CustomerVersionNumber"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="TelephoneSeqNum!='' and TelephoneSeqNum!='0'">
						<xsl:attribute name="TELEPHONESEQUENCENUMBER"><xsl:value-of select="TelephoneSeqNum"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="Number"/></xsl:attribute>
				</xsl:if>
				
				<!-- CONTACTTELEPHONEDETAILS -->
			   <xsl:if test="$Label='CONTACTTELEPHONEDETAILS'">
				   <xsl:if test="ContactDetailsGuid!=''">
						<xsl:attribute name="CONTACTDETAILSGUID"><xsl:value-of select="ContactDetailsGuid"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="TelephoneSeqNum!='' and TelephoneSeqNum!='0'">
						<xsl:attribute name="TELEPHONESEQNUM"><xsl:value-of select="TelephoneSeqNum"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="TELENUMBER"><xsl:value-of select="Number"/></xsl:attribute>
				</xsl:if>
				
				<!-- GENERIC PHONE DATA -->
				<xsl:attribute name="AREACODE"><xsl:value-of select="AreaCode"/></xsl:attribute>				
				<xsl:attribute name="EXTENSIONNUMBER"><xsl:value-of select="Extension"/></xsl:attribute>
				<xsl:if test="Preferred!=''">
					<xsl:attribute name="PREFERREDMETHODOFCONTACT">
						<xsl:choose>
							<xsl:when test="Preferred='true'"><xsl:value-of select="'1'" /></xsl:when>
							<xsl:otherwise><xsl:value-of select="'0'" /></xsl:otherwise>
						</xsl:choose>
					</xsl:attribute>
				</xsl:if>
				<xsl:if test="Usage!=''">
					<xsl:attribute name="USAGE"><xsl:value-of select="Usage"/></xsl:attribute>
				</xsl:if>
			</xsl:element>
		</xsl:if>

	</xsl:template>
	
		<xsl:template name="createFormattedDate">
			<xsl:param name="label" />
			<xsl:param name="date" />
			<xsl:if test="$date!='' and substring($date, 1, 4)!='0001'">
				<xsl:attribute name="{$label}">
					<xsl:value-of select="concat(substring($date, 9, 2), '/', substring($date, 6, 2), '/', substring($date, 1, 4))" />
				</xsl:attribute>
			</xsl:if>
		</xsl:template>
	
	<xsl:template name="CreateMemoPad">
		<xsl:param name="MemoPadId"/>		
		<xsl:param name="EntryType"/>
		<xsl:param name="MemoEntry"/>				
		
		<xsl:if test="$MemoEntry!='' and $EntryType!=''">
			<xsl:element name="MEMOPAD">
				<xsl:if test="$MemoPadId!='' and $MemoPadId!='0'">
					<xsl:attribute name="MEMOPADID"><xsl:value-of select="$MemoPadId"/></xsl:attribute>
				</xsl:if>
				<xsl:attribute name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="$AppFFNo" /></xsl:attribute> 
				<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="$AppNo"/></xsl:attribute>
				<xsl:attribute name="MEMOENTRY">
					<xsl:choose>
						<xsl:when test="substring($MemoEntry, 1, 4)!='0001'">
							<xsl:value-of select="$MemoEntry" />
						</xsl:when>
						<xsl:otherwise>
							<!-- Do nothing - invalid date data -->
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:attribute name="ENTRYTYPE"><xsl:value-of select="$EntryType"/></xsl:attribute>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	
	<xsl:template name="CreateAccountRelationship">
		<xsl:param name="CustomerNumber"/>
		<xsl:param name="CustomerVersionNumber"/>
		<xsl:param name="AccGuid"/>
		<xsl:param name="CustomerRoleType"/>
		
		<xsl:element name="ACCOUNTRELATIONSHIP">
			<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$CustomerNumber"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$CustomerVersionNumber"/></xsl:attribute>
			<xsl:if test="$AccGuid!=''">
				<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="$AccGuid"/></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="CUSTOMERROLETYPE"><xsl:value-of select="$CustomerRoleType"/></xsl:attribute>			
		</xsl:element>
	</xsl:template>
	
</xsl:stylesheet>
