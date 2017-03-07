<?xml version="1.0" encoding="UTF-8"?>
<!--
History:

Author	Date			Description
GHun	12/02/2006	EP2_954 Additional borrowing changes
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="/">
		<xsl:element name="RESPONSE">

			<xsl:variable name="Mortgage" select="/XML/Response/Outputs/Mortgage"/>
			<xsl:variable name="InputsControl" select="/XML/Response/Inputs/Control"/>
			<xsl:variable name="LoanComponents" select="//LOANCOMPONENTLIST/LOANCOMPONENT"/>
			
			<xsl:element name="LOANCOMPONENTLIST">
				<!-- LoanComponent -->
				<xsl:for-each select="$LoanComponents">
					<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
					<xsl:variable name="ElementGroup" select="$Mortgage/ElementGroup[@Id=concat('ElementGroup',$LCSeqNo)]"/>
					
					<xsl:element name="LOANCOMPONENT">
						<xsl:element name="APPLICATIONNUMBER">
							<xsl:value-of select="APPLICATIONNUMBER"/>
						</xsl:element>
						<xsl:element name="APPLICATIONFACTFINDNUMBER">
							<xsl:value-of select="APPLICATIONFACTFINDNUMBER"/>
						</xsl:element>
						<xsl:element name="MORTGAGESUBQUOTENUMBER">
							<xsl:value-of select="MORTGAGESUBQUOTENUMBER"/>
						</xsl:element>
						<xsl:element name="LOANCOMPONENTSEQUENCENUMBER">
							<xsl:value-of select="$LCSeqNo"/>
						</xsl:element>
						<xsl:element name="TOTALLOANCOMPONENTAMOUNT">
							<xsl:value-of select="TOTALLOANCOMPONENTAMOUNT"/>
						</xsl:element>
						<!-- JD BMIDS805 added interestOnlyElement, capitalAndInterestElement and PortedLoan for dataingestion (SaveOneOffCostDetails) -->
						<!-- JD BMIDS816 If there is more than one element then this is a Part and Part, so extra fields should be populated -->
						<xsl:if test="count($ElementGroup/Element) &gt; 1">
							<xsl:element name="INTERESTONLYELEMENT">
								<xsl:value-of select="INTERESTONLYELEMENT"/>
							</xsl:element>
							<xsl:element name="CAPITALANDINTERESTELEMENT">
								<xsl:value-of select="CAPITALANDINTERESTELEMENT"/>
							</xsl:element>
						</xsl:if>
						<xsl:element name="PORTEDLOAN">
							<xsl:value-of select="PORTEDLOAN"/>
						</xsl:element>
						<xsl:element name="APR">
							<xsl:value-of select="$ElementGroup/@APR"/>
						</xsl:element>
						<!-- JD BMIDS804 add finalpayment -->
						<xsl:element name="FINALPAYMENT">
							<xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='FinalPayment']/@Amount"/>
						</xsl:element>
						<xsl:variable name="MonthlyCost" select="$ElementGroup/@InitialMonthlyPayment"/>
						<xsl:element name="NETMONTHLYCOST">
							<xsl:value-of select="$MonthlyCost"/>
						</xsl:element>
						<!-- MonthlyCostLessDrawDown is calculated by making an extra call to Alpha+ and populated later-->
						<xsl:element name="GROSSMONTHLYCOST">
							<xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment']/@Amount"/>
						</xsl:element>
						
						<!-- SVR comparison -->
						<xsl:if test="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-SVR')]">
							<xsl:element name="FINALRATE">
								<xsl:value-of select="$InputsControl/Mortgage/ParameterOverride[@Id=concat('EG', $LCSeqNo, '-POIIR-SVR')]/@OverrideValue"/>
							</xsl:element>
							<xsl:element name="FINALRATEMONTHLYCOST">
								<xsl:value-of select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-SVR')]/Values[@Type='MonthlyPayment']/@Amount"/>
							</xsl:element>
						</xsl:if>
						
						<!-- Alpha+ currently doesn't calculate this
						<xsl:element name="FINALRATEAPR"></xsl:element>
						-->

						<xsl:element name="RESOLVEDRATE">
							<xsl:value-of select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[1]/ALPHAPLUSRATE"/>
						</xsl:element>
						<xsl:element name="AMOUNTPERUNITBORROWED">
							<xsl:value-of select="$ElementGroup/@AmountPerUnitBorrowed"/>
						</xsl:element>
						<!-- Cost at floored rate -->
						<xsl:variable name="MinMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-Min1')]/Values[@Type='MonthlyPayment']/@Amount"/>
						<xsl:if test="string($MinMonthlyCost)">
							<xsl:element name="MINMONTHLYCOST">
								<xsl:value-of select="$MinMonthlyCost"/>
							</xsl:element>
						</xsl:if>
						<!-- Cost at capped rate -->
						<xsl:variable name="MaxMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-Max1')]/Values[@Type='MonthlyPayment']/@Amount"/>
						<xsl:if test="string($MaxMonthlyCost)">
							<xsl:element name="MAXMONTHLYCOST">
								<xsl:value-of select="$MaxMonthlyCost"/>
							</xsl:element>
						</xsl:if>
						<xsl:element name="UNROUNDEDAPR">
							<xsl:value-of select="$ElementGroup/@UnroundedAPR"/>
						</xsl:element>
						<xsl:element name="ACCRUEDINTEREST">
							<xsl:value-of select="$ElementGroup/@AccruedInterest"/>
						</xsl:element>
						<!-- GHun Only get the IncreasedMonthlyCost if the 1% increase applies to the 1st InterestRateType -->
						<xsl:if test="string(MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[1]/ALPHAPLUSINCREASEDRATE)">
							<xsl:variable name="IncreasedMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-OnePercent')]/Values[@Type='MonthlyPayment']/@Amount"/>
							<xsl:element name="INCREASEDMONTHLYCOST">
								<xsl:value-of select="$IncreasedMonthlyCost"/>
							</xsl:element>
							
							<!-- Calculate the difference between the MonthlyCost and the IncreasedMontlyCost -->
							<xsl:element name="INCREASEDMONTHLYCOSTDIFFERENCE">
								<xsl:value-of select="format-number(number($IncreasedMonthlyCost) - number($MonthlyCost), '0.00')"/>
							</xsl:element>
						</xsl:if>
						<xsl:element name="TOTALAMOUNTPAYABLE">
							<xsl:value-of select="$ElementGroup/@TotalPayments"/>
						</xsl:element>
						
						<!-- If there is more than one element then this is a Part and Part, so extra fields should be populated -->
						<xsl:if test="count($ElementGroup/Element) &gt; 1">
							<xsl:element name="CAPINTMONTHLYCOST">
								<xsl:value-of select="$ElementGroup/Element[@Type='Repayment']/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment']/@Amount"/>
							</xsl:element>
							<xsl:element name="INTONLYMONTHLYCOST">
								<xsl:value-of select="$ElementGroup/Element[@Type='InterestOnly']/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment']/@Amount"/>
							</xsl:element>
							<xsl:element name="CAPINTACCRUEDINTEREST">
								<xsl:value-of select="$ElementGroup/Element[@Type='Repayment']/@AccruedInterest"/>
							</xsl:element>
							<xsl:element name="INTONLYACCRUEDINTEREST">
								<xsl:value-of select="$ElementGroup/Element[@Type='InterestOnly']/@AccruedInterest"/>
							</xsl:element>
						</xsl:if>
					</xsl:element>
				</xsl:for-each>
	
				<!-- LoanComponentPaymentSchedule -->
				<xsl:for-each select="$LoanComponents">
					<xsl:variable name="AppNo" select="APPLICATIONNUMBER"/>
					<xsl:variable name="AppFFNo" select="APPLICATIONFACTFINDNUMBER"/>
					<xsl:variable name="MSQNo" select="MORTGAGESUBQUOTENUMBER"/>
					<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
					<xsl:variable name="ElementGroup" select="$Mortgage/ElementGroup[@Id=concat('ElementGroup',$LCSeqNo)]"/>
					
					<!-- GHun Find the InterestRateTypeSequence number that has a 1% increase applied to the rate -->
					<xsl:variable name="IncreaseIRTSeqNo" select="number(MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[string(ALPHAPLUSINCREASEDRATE)]/INTERESTRATETYPESEQUENCENUMBER)"/>
					
					<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE">
						<xsl:sort data-type="number" select="number(INTERESTRATETYPESEQUENCENUMBER)"/>
						<xsl:variable name="IRTSeqNo" select="number(INTERESTRATETYPESEQUENCENUMBER)"/>
						<xsl:variable name="RateIndex" select="position()"/>
						<xsl:variable name="MonthlyCost" select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/>
						
						<!-- BMIDS859 Only create a LoanComponentPaymentSchedule element if there is a monthly cost -->
						<xsl:if test="string($MonthlyCost)">
							<xsl:element name="LOANCOMPONENTPAYMENTSCHEDULE">
								<xsl:element name="APPLICATIONNUMBER">
									<xsl:value-of select="$AppNo"/>
								</xsl:element>
								<xsl:element name="APPLICATIONFACTFINDNUMBER">
									<xsl:value-of select="$AppFFNo"/>
								</xsl:element>
								<xsl:element name="MORTGAGESUBQUOTENUMBER">
									<xsl:value-of select="$MSQNo"/>
								</xsl:element>
								<xsl:element name="LOANCOMPONENTSEQUENCENUMBER">
									<xsl:value-of select="$LCSeqNo"/>
								</xsl:element>
								<xsl:element name="INTERESTRATETYPESEQUENCENUMBER">
									<xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/>
								</xsl:element>
								<!-- MAR122 GHun -->
								<xsl:element name="INTERESTRATE">
									<xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@InterestRate"/>
								</xsl:element>
								<!-- MAR122 End -->
								<xsl:variable name="StartDate1" select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Date"/>	
								<xsl:element name="STARTDATE">
									<xsl:value-of select="substring($StartDate1,9,2)"/>/<xsl:value-of select="substring($StartDate1,6,2)"/>/<xsl:value-of select="substring($StartDate1,1,4)"/>
								</xsl:element>
								<xsl:element name="MONTHLYCOST">
									<xsl:value-of select="$MonthlyCost"/>
								</xsl:element>
								<!-- Cost at floored rate -->
								<xsl:variable name="MinMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-Min', $IRTSeqNo)]/Values[@Type='MonthlyPayment']/@Amount"/>
								<xsl:if test="string($MinMonthlyCost)">
									<xsl:element name="MINMONTHLYCOST">
										<xsl:value-of select="$MinMonthlyCost"/>
									</xsl:element>
								</xsl:if>
								<!-- Cost at capped rate -->
								<xsl:variable name="MaxMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-Max', $IRTSeqNo)]/Values[@Type='MonthlyPayment']/@Amount"/>
								<xsl:if test="string($MaxMonthlyCost)">
									<xsl:element name="MAXMONTHLYCOST">
										<xsl:value-of select="$MaxMonthlyCost"/>
									</xsl:element>
								</xsl:if>
								<!-- GHun Only get the IncreaseMonthlyCost for the one InterestRateType where the 1% increase has been applied -->
								<xsl:if test="$IncreaseIRTSeqNo=$IRTSeqNo">
									<xsl:variable name="IncreasedMonthlyCost" select="$ElementGroup/Schedule[@Type='IncreasedPayments' and @ParameterOverrideId=concat('EG',$LCSeqNo,'-POIIR-OnePercent')]/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/>
									<xsl:element name="INCREASEDMONTHLYCOST">
										<xsl:value-of select="$IncreasedMonthlyCost"/>
									</xsl:element>
									<!-- Calculate the difference between the MonthlyCost and the IncreasedMontlyCost -->
									<xsl:element name="INCREASEDMONTHLYCOSTDIFFERENCE">
										<xsl:value-of select="format-number(number($IncreasedMonthlyCost) - number($MonthlyCost), '0.00')"/>
									</xsl:element>
								</xsl:if>
								<!-- If there is more than one element then this is a Part and Part, so 2 extra fields should also be saved-->
								<xsl:if test="count($ElementGroup/Element) &gt; 1">
									<xsl:element name="CAPINTMONTHLYCOST">
										<xsl:value-of select="$ElementGroup/Element[@Type='Repayment']/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/>
									</xsl:element>
									<xsl:element name="INTONLYMONTHLYCOST">
										<xsl:value-of select="$ElementGroup/Element[@Type='InterestOnly']/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/>							
									</xsl:element>
								</xsl:if>
							</xsl:element>
						</xsl:if>
					</xsl:for-each>
	
				</xsl:for-each>			
	
				<!-- LoanComponentBalanceSchedule -->			
				<xsl:for-each select="$LoanComponents">
					<xsl:variable name="AppNo" select="APPLICATIONNUMBER"/>
					<xsl:variable name="AppFFNo" select="APPLICATIONFACTFINDNUMBER"/>
					<xsl:variable name="MSQNo" select="MORTGAGESUBQUOTENUMBER"/>
					<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
					<xsl:variable name="ScheduleType" select="MORTGAGEPRODUCTDETAILS/MORTGAGELENDERPARAMETERS/OUTPUTBALANCESCHEDULE"/>
					
					<!-- Loop through each OutstandingLoan Schedule Value for the current LoanComponent/ElementGroup -->
					<xsl:for-each select="$Mortgage/ElementGroup[@Id=concat('ElementGroup',$LCSeqNo)]/Schedule[@Type='OutstandingLoan']/Values">
						<xsl:variable name="Date" select="@Date"/>
						<xsl:element name="LOANCOMPONENTBALANCESCHEDULE">
							<xsl:element name="APPLICATIONNUMBER">
								<xsl:value-of select="$AppNo"/>
							</xsl:element>
							<xsl:element name="APPLICATIONFACTFINDNUMBER">
								<xsl:value-of select="$AppFFNo"/>
							</xsl:element>
							<xsl:element name="MORTGAGESUBQUOTENUMBER">
								<xsl:value-of select="$MSQNo"/>
							</xsl:element>
							<xsl:element name="LOANCOMPONENTSEQUENCENUMBER">
								<xsl:value-of select="$LCSeqNo"/>
							</xsl:element>
							<xsl:element name="SCHEDULETYPE">
								<xsl:value-of select="$ScheduleType"/>
							</xsl:element>
							<xsl:element name="STARTDATE">
								<xsl:value-of select="@Date"/>
							</xsl:element>
							<xsl:element name="BALANCE">
								<xsl:value-of select="@Amount"/>
							</xsl:element>
							<!-- If there is more than one element then this is a Part and Part, so 2 extra fields should also be saved-->
							<xsl:if test="count(../../Element) &gt; 1">
								<xsl:element name="CAPINTBALANCE">
									<xsl:value-of select="../../Element[@Type='Repayment']/Schedule[@Type='OutstandingLoan']/Values[@Date=$Date]/@Amount"/>
								</xsl:element>
								<xsl:element name="INTONLYBALANCE">
									<xsl:value-of select="../../Element[@Type='InterestOnly']/Schedule[@Type='OutstandingLoan']/Values[@Date=$Date]/@Amount"/>
								</xsl:element>
							</xsl:if>
						</xsl:element>
					</xsl:for-each>
				</xsl:for-each>
	
				<!-- LoanComponentRedemptionFee -->			
				<xsl:for-each select="$LoanComponents">
					<xsl:variable name="AppNo" select="APPLICATIONNUMBER"/>
					<xsl:variable name="AppFFNo" select="APPLICATIONFACTFINDNUMBER"/>
					<xsl:variable name="MSQNo" select="MORTGAGESUBQUOTENUMBER"/>
					<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
			
					<xsl:for-each select="MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND">
						<xsl:sort data-type="number" select="number(REDEMPTIONFEESTEPNUMBER)"/>
						
						<!-- Find the Redemption Fee Amount -->
						<xsl:variable name="PriorStepNo" select="number(REDEMPTIONFEESTEPNUMBER)-1"/>
						<xsl:variable name="Duration" select="sum(../REDEMPTIONFEEBAND[REDEMPTIONFEESTEPNUMBER&lt;=$PriorStepNo]/PERIOD)"/>
						<!-- Extract the position of the Duration from the RepaymentChargeTerms list in the input -->
						<!-- RepaymentChargeTerms is space separted, so the position can be found by counting the number of spaces that occur before the Duration -->
						<xsl:variable name="Before" select="substring-before($InputsControl/@RepaymentChargeTerms ,$Duration)"/>
						<xsl:variable name="Stripped" select="translate($Before, ' ', '')"/>
						<xsl:variable name="Position" select="string-length($Before) - string-length($Stripped) + 1"/>
						<xsl:variable name="RedemptionFee" select="$Mortgage/ElementGroup[@Id=concat('ElementGroup',$LCSeqNo)]/Schedule[@Type='RepaymentCharge']/Values[$Position]/@Amount"/>

						<!-- BMIDS859 Only create LoanComponentRedemptionFee elements if a fee exists -->
						<xsl:if test="string($RedemptionFee)">
							<xsl:element name="LOANCOMPONENTREDEMPTIONFEE">
								<xsl:element name="APPLICATIONNUMBER">
									<xsl:value-of select="$AppNo"/>
								</xsl:element>
								<xsl:element name="APPLICATIONFACTFINDNUMBER">
									<xsl:value-of select="$AppFFNo"/>
								</xsl:element>
								<xsl:element name="MORTGAGESUBQUOTENUMBER">
									<xsl:value-of select="$MSQNo"/>
								</xsl:element>
								<xsl:element name="LOANCOMPONENTSEQUENCENUMBER">
									<xsl:value-of select="$LCSeqNo"/>
								</xsl:element>
								<xsl:element name="REDEMPTIONFEESTEPNUMBER">
									<xsl:value-of select="REDEMPTIONFEESTEPNUMBER"/>
								</xsl:element>
								<xsl:choose>
									<xsl:when test="string(PERIODENDDATE)">
										<xsl:element name="REDEMPTIONFEEPERIODENDDATE">
											<xsl:value-of select="PERIODENDDATE"/>
										</xsl:element>
									</xsl:when>
									<xsl:otherwise>
										<xsl:element name="REDEMPTIONFEEPERIOD">
											<xsl:value-of select="PERIOD"/>
										</xsl:element>
									</xsl:otherwise>
								</xsl:choose>
								<xsl:element name="REDEMPTIONFEEAMOUNT">
									<xsl:value-of select="$RedemptionFee"/>				
								</xsl:element>
							</xsl:element>
						</xsl:if>
					</xsl:for-each>
				</xsl:for-each>
				
				<!-- EP2_954 GHun MortgageLoanPaymentSchedule -->
				<xsl:if test="/XML/REQUEST/MORTGAGELOANLIST/MORTGAGELOAN">
					<xsl:for-each select="/XML/REQUEST/MORTGAGELOANLIST/MORTGAGELOAN">
						<xsl:variable name="LoanGUID" select="MORTGAGELOANGUID"/>
						<xsl:variable name="MLSeqNo" select="position()"/>
						<xsl:variable name="ElementGroup" select="/XML/ADDITIONALBORROWING/Response/Outputs/Mortgage/ElementGroup[@Id=concat('ExistingMortgageAccount',$MLSeqNo)]"/>
						<xsl:for-each select="INTERESTRATETYPE">
							<xsl:variable name="RateIndex" select="position()"/>
							<xsl:variable name="MonthlyCost" select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/>
							<xsl:if test="string($MonthlyCost)">
								<xsl:element name="MORTGAGELOANPAYMENTSCHEDULE">
									<xsl:element name="MORTGAGELOANGUID"><xsl:value-of select="$LoanGUID"/></xsl:element>
									<xsl:element name="MONTHLYCOST"><xsl:value-of select="$MonthlyCost"/></xsl:element>
									<xsl:element name="RATEPERIODSTARTDATE"><xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Date"/></xsl:element>
									<xsl:element name="INTERESTRATE"><xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@InterestRate"/></xsl:element>
									<xsl:element name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:element>
									<xsl:element name="INTERESTRATEPERIOD"><xsl:value-of select="INTERESTRATEPERIOD"/></xsl:element>
								</xsl:element>
							</xsl:if>
						</xsl:for-each>
					</xsl:for-each>
					
					<!-- TotalMortgageBorrowing -->
					<xsl:element name="TOTALMORTGAGEBORROWING">
						<xsl:element name="APPLICATIONNUMBER"><xsl:value-of select="$LoanComponents/APPLICATIONNUMBER"/></xsl:element>
						<xsl:element name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="$LoanComponents/MORTGAGESUBQUOTENUMBER"/></xsl:element>
						<xsl:element name="TOTALMORTGAGEBALANCE">
							<xsl:variable name="ExistingMortgageBalance" select="sum(/XML/REQUEST/MORTGAGELOANLIST/MORTGAGELOAN/OUTSTANDINGBALANCE)"/>
							<xsl:variable name="NewLendingBalance" select="sum($LoanComponents/TOTALLOANCOMPONENTAMOUNT)"/>
							<xsl:value-of select="$ExistingMortgageBalance + $NewLendingBalance"/>
						</xsl:element>
						<xsl:element name="TOTALINITIALMONTHLYCOST">
							<xsl:variable name="ExistingMortgageCost" select="/XML/ADDITIONALBORROWING/Response/Outputs/Mortgage/@InitialMonthlyPayment"/>
							<xsl:variable name="NewLendingCost" select="$Mortgage/@InitialMonthlyPayment"/>
							<xsl:value-of select="number($ExistingMortgageCost) + number($NewLendingCost)"/>
						</xsl:element>
					</xsl:element>					
					
					<!-- Create a temporary nodeset holding the the merged list of payments for loan components and mortgage loans -->
					<xsl:variable name="TotalMortgageBorrowingPayments">
						<xsl:for-each select="$LoanComponents">
							<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
							<xsl:variable name="ElementGroup" select="$Mortgage/ElementGroup[@Id=concat('ElementGroup',$LCSeqNo)]"/>
							<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE">
								<xsl:variable name="RateIndex" select="position()"/>
								<xsl:element name="PAYMENT">
									<xsl:attribute name="ISMORTGAGELOAN">0</xsl:attribute>
									<xsl:attribute name="MONTHLYCOST"><xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/></xsl:attribute>
									<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
									<xsl:attribute name="RATETYPE"><xsl:value-of select="../INTERESTRATETYPE[number($RateIndex)-1]/RATETYPE"/></xsl:attribute>
									<xsl:attribute name="DURATION"><xsl:value-of select="$InputsControl/Mortgage/ParameterOverride[@Id=concat('EG', $LCSeqNo, '-POIR')]/ParameterOverrideValue[$RateIndex]/@Duration"/></xsl:attribute>
									<xsl:attribute name="ID"><xsl:value-of select="concat('LC', $LCSeqNo)"/></xsl:attribute>
									<xsl:element name="ID"><xsl:value-of select="concat('LC', $LCSeqNo)"/></xsl:element>
								</xsl:element>
							</xsl:for-each>
						</xsl:for-each>
						<xsl:for-each select="/XML/REQUEST/MORTGAGELOANLIST/MORTGAGELOAN">
							<xsl:variable name="MLSeqNo" select="position()"/>
							<xsl:variable name="ElementGroup" select="/XML/ADDITIONALBORROWING/Response/Outputs/Mortgage/ElementGroup[@Id=concat('ExistingMortgageAccount',$MLSeqNo)]"/>
							<xsl:for-each select="INTERESTRATETYPE">
								<xsl:variable name="RateIndex" select="position()"/>
								<xsl:element name="PAYMENT">
									<xsl:attribute name="ISMORTGAGELOAN">1</xsl:attribute>
									<xsl:attribute name="MONTHLYCOST"><xsl:value-of select="$ElementGroup/Schedule[@Type='Payments']/Values[@Type='MonthlyPayment'][$RateIndex]/@Amount"/></xsl:attribute>
									<xsl:attribute name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="INTERESTRATETYPESEQUENCENUMBER"/></xsl:attribute>
									<xsl:attribute name="RATETYPE"><xsl:value-of select="../INTERESTRATETYPE[number($RateIndex)-1]/RATETYPE"/></xsl:attribute>
									<xsl:attribute name="DURATION"><xsl:value-of select="/XML/ADDITIONALBORROWING/Response/Inputs/Control/Mortgage/ParameterOverride[@Id=concat('EMA', $MLSeqNo, '-Payments')]/ParameterOverrideValue[$RateIndex]/@Duration"/></xsl:attribute>
									<xsl:attribute name="ID"><xsl:value-of select="concat('ML', $MLSeqNo)"/></xsl:attribute>									
									<xsl:element name="ID"><xsl:value-of select="concat('ML', $MLSeqNo)"/></xsl:element>
								</xsl:element>
							</xsl:for-each>
						</xsl:for-each>
					</xsl:variable>
					
					<!-- TotalMortgageBorrowingSchedule -->
					<xsl:variable name="Payments" select="msxsl:node-set($TotalMortgageBorrowingPayments)/PAYMENT"/>
					<xsl:variable name="ID" select="$Payments[not(ID=preceding-sibling::PAYMENT/ID)]/ID"/>
					
					<xsl:for-each select="$Payments[@DURATION>0]">
						<xsl:sort select="@DURATION"/>
						<xsl:sort select="@ISMORTGAGELOAN" order="descending"/>
						<xsl:sort select="@ID"/>
						<xsl:element name="TOTALMORTGAGEBORROWINGSCHEDULE">
							<xsl:element name="APPLICATIONNUMBER"><xsl:value-of select="$LoanComponents/APPLICATIONNUMBER"/></xsl:element>
							<xsl:element name="MORTGAGESUBQUOTENUMBER"><xsl:value-of select="$LoanComponents/MORTGAGESUBQUOTENUMBER"/></xsl:element>
							<xsl:variable name="CurrentDuration" select="number(@DURATION)"/>
							<xsl:variable name="PreviousDuration" select="../PAYMENT[@DURATION &lt; $CurrentDuration][last()]/@DURATION"/>
							<xsl:element name="STEPDURATION"><xsl:value-of select="$CurrentDuration - $PreviousDuration"/></xsl:element>
							<xsl:element name="SEQUENCENUMBER"><xsl:value-of select="position()"/></xsl:element>
							<xsl:element name="INTERESTRATETYPESEQUENCENUMBER"><xsl:value-of select="@INTERESTRATETYPESEQUENCENUMBER"/></xsl:element>
							<xsl:element name="TOTALDURATIONMONTHLYCOST">
								<xsl:call-template name="SumTotalDurationCost">
									<xsl:with-param name="Payments" select="$Payments[number(@DURATION) &lt;= $CurrentDuration]"/>
									<xsl:with-param name="LCID" select="$ID"/>
									<xsl:with-param name="RunningTotal" select="0"/>
								</xsl:call-template>
							</xsl:element>
							<xsl:element name="RATETYPE"><xsl:value-of select="@RATETYPE"/></xsl:element>
							<xsl:element name="ISMORTGAGELOAN"><xsl:value-of select="@ISMORTGAGELOAN"/></xsl:element>
							<xsl:element name="DURATION"><xsl:value-of select="@DURATION"/></xsl:element>
						</xsl:element>
					</xsl:for-each>
				</xsl:if>
				<!-- EP2_954 End -->
				
			</xsl:element>
			
			<!-- MortgageSubQuote totals -->
			<xsl:element name="TOTALNETMONTHLYCOST">
				<xsl:value-of select="$Mortgage/@InitialMonthlyPayment"/>
			</xsl:element>
			<xsl:element name="TOTALGROSSMONTHLYCOST">
				<xsl:value-of select="$Mortgage/@InitialMonthlyPayment"/>
			</xsl:element>			
			<xsl:element name="TOTALACCRUEDINTEREST">
				<xsl:value-of select="$Mortgage/@AccruedInterest"/>
			</xsl:element>
			<xsl:element name="AMOUNTPERUNITBORROWED">
				<xsl:value-of select="$Mortgage/@AmountPerUnitBorrowed"/>
			</xsl:element>
			<xsl:element name="APR">
				<xsl:value-of select="$Mortgage/@SecuredAPR"/>
			</xsl:element>
			<xsl:element name="TOTALAMOUNTPAYABLE">
				<xsl:value-of select="$Mortgage/@TotalPayments"/>
			</xsl:element>
			<xsl:element name="TOTALMORTGAGEPAYMENTS">
				<xsl:value-of select="$Mortgage/@TotalMortgagePayments"/>
			</xsl:element>
			
		</xsl:element>			
	</xsl:template>

	<!-- EP2_954 GHun Recursive template that sums up the monthly costs across the loan components (including mortgage loans).
		The monthly cost from the most recent payment (highest duration) per loan component is used and relies on the Payments begin sorted by duration (ascending) -->
	<xsl:template name="SumTotalDurationCost">
		<xsl:param name="Payments"/>
		<xsl:param name="LCID"/>
		<xsl:param name="RunningTotal"/>
		<xsl:choose>
			<xsl:when test="not($LCID)">
				<xsl:copy-of select="format-number($RunningTotal, '##########.##')"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="CurrentTotal" select="$RunningTotal + number($Payments[@ID=$LCID[1]][last()]/@MONTHLYCOST)"/>
				<xsl:call-template name="SumTotalDurationCost">
					<xsl:with-param name="Payments" select="$Payments"/>
					<xsl:with-param name="LCID" select="$LCID[position()>1]"/>
					<xsl:with-param name="RunningTotal" select="$CurrentTotal"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!-- EP2_954 End -->
</xsl:stylesheet>