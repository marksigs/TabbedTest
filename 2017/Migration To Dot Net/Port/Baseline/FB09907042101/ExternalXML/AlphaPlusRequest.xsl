<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="/">
		<xsl:element name="Request">
			<xsl:element name="Inputs">
				<xsl:attribute name="Function">Mortgage</xsl:attribute>
				<!-- Mortgage Calculation Method -->
				<xsl:attribute name="Method">FullFixedTermQuote</xsl:attribute>
				<xsl:element name="Control">
					<xsl:variable name="LoanComponents" select="//LOANCOMPONENTLIST/LOANCOMPONENT"/>
					<xsl:variable name="FirstLoanComp" select="$LoanComponents[1]"/>
					<xsl:attribute name="Id"><xsl:value-of select="concat('App-',$FirstLoanComp/APPLICATIONNUMBER)"/></xsl:attribute>
					<xsl:attribute name="IsTraceRequired">No</xsl:attribute>
					<xsl:attribute name="QuotationDate"><xsl:value-of select="/@DATE"/></xsl:attribute>
					<xsl:attribute name="OutputPaymentSchedule">
						<xsl:value-of select="$FirstLoanComp/MORTGAGEPRODUCTDETAILS/MORTGAGELENDERPARAMETERS/OUTPUTPAYMENTSCHEDULE/@TEXT"/>
					</xsl:attribute>
					<xsl:attribute name="OutputBalanceSchedule">
						<xsl:value-of select="$FirstLoanComp/MORTGAGEPRODUCTDETAILS/MORTGAGELENDERPARAMETERS/OUTPUTBALANCESCHEDULE/@TEXT"/>
					</xsl:attribute>
					<!-- RepaymentChargeTerms are populated later once the early repayment charge durations for each component have been calculated -->
					<xsl:attribute name="APRRequired">Composite</xsl:attribute>
					<xsl:element name="Mortgage">
						<xsl:attribute name="CompletionDate"><xsl:value-of select="/@COMPLETIONDATE"/></xsl:attribute>
						
						<xsl:attribute name="PropertyValue"><xsl:value-of select="sum($LoanComponents/LOANAMOUNT)"/></xsl:attribute>

						<!-- Set up the mortgage lender parameters as parameter overrides-->
						<xsl:apply-templates select="$FirstLoanComp/MORTGAGEPRODUCTDETAILS/MORTGAGELENDERPARAMETERS"/>
												
						<!-- Create an ElementGroup for each loan component -->
						<xsl:for-each select="$LoanComponents">
							<xsl:variable name="LCSeqNo" select="LOANCOMPONENTSEQUENCENUMBER"/>
							<xsl:variable name="NumRepaymentCharges" select="count(MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND)"/>

							<xsl:element name="ElementGroup">
								<!-- Element Group Id -->
								<xsl:attribute name="Id"><xsl:value-of select="concat('ElementGroup', $LCSeqNo)"/></xsl:attribute>
								<xsl:attribute name="IsSecured">Yes</xsl:attribute>
								<!-- Element Term (in months) -->
								<xsl:attribute name="Term"><xsl:value-of select="number(TERMINYEARS)*12 + number(TERMINMONTHS)"/></xsl:attribute>
								<!-- Rate Adjustment: The manual adjustment is already included in the interest rates so rate adjustment should be 0 -->
								<xsl:attribute name="RateAdjustment">0.0</xsl:attribute>

								<!-- ParameterOverrideIds are now set by doing a second pass through the produced XML -->
							
								<xsl:choose>
									<xsl:when test="REPAYMENTMETHOD/VALIDATIONTYPELIST/VALIDATIONTYPE='P'">
									<!-- Handle Part and Part separately as it needs to be split into 2 Elements -->
										
										<!-- Capital and Interest -->
										<xsl:element name="Element">
											<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, 'Element1')"/></xsl:attribute>
											<xsl:attribute name="Type">Repayment</xsl:attribute>
											<xsl:element name="Loan">
												<xsl:attribute name="Amount"><xsl:value-of select="CAPITALANDINTERESTELEMENT"/></xsl:attribute>
												<xsl:attribute name="Frequency">0</xsl:attribute>
											</xsl:element>
										</xsl:element>
										
										<!-- Interest Only -->
										<xsl:element name="Element">
											<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, 'Element2')"/></xsl:attribute>
											<xsl:attribute name="Type">InterestOnly</xsl:attribute>
											<xsl:element name="Loan">
												<xsl:attribute name="Amount"><xsl:value-of select="INTERESTONLYELEMENT"/></xsl:attribute>
												<xsl:attribute name="Frequency">0</xsl:attribute>
											</xsl:element>
										</xsl:element>	
									</xsl:when>								

									<xsl:otherwise>
									<!-- All other repayment types (non part and part)-->
										<xsl:element name="Element">
											<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, 'Element1')"/></xsl:attribute>
											<xsl:attribute name="Type">
												<xsl:if test="REPAYMENTMETHOD/VALIDATIONTYPELIST/VALIDATIONTYPE='I'">InterestOnly</xsl:if>
												<xsl:if test="REPAYMENTMETHOD/VALIDATIONTYPELIST/VALIDATIONTYPE='C'">Repayment</xsl:if>
											</xsl:attribute>
											<xsl:element name="Loan">
												<xsl:attribute name="Amount">
													<xsl:choose>
														<xsl:when test="string(//LOANCOMPONENTLIST/@FROMPAYPROC)='TRUE'">
															<xsl:value-of select="TOTALLOANCOMPONENTAMOUNT"/>
														</xsl:when>
														<xsl:otherwise>
															<xsl:value-of select="LOANAMOUNT"/>
														</xsl:otherwise>
													</xsl:choose>
												</xsl:attribute>
												<xsl:attribute name="Frequency">0</xsl:attribute>
											</xsl:element>
										</xsl:element>		
									</xsl:otherwise>
								</xsl:choose>
																	
							</xsl:element>

							<!-- Interest Rate "table"-->
							<xsl:element name="ParameterOverride">
								<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo,'-POIR')"/></xsl:attribute>
								<xsl:attribute name="Parameter">InterestRate</xsl:attribute>
								<xsl:element name="ParameterOverrideKey">
									<xsl:attribute name="Name">Duration</xsl:attribute>
									<xsl:attribute name="NumberOfValues"><xsl:value-of select="count(MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE)"/></xsl:attribute>
								</xsl:element>
								<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE">
									<xsl:sort data-type="number" select="number(INTERESTRATETYPESEQUENCENUMBER)"/>
									<xsl:variable name="PriorSeqNo" select="number(INTERESTRATETYPESEQUENCENUMBER)-1"/>
									<xsl:element name="ParameterOverrideValue">
										<xsl:attribute name="Duration">
											<!-- The duration is the sum of all the periods before the current one -->
											<xsl:value-of select="sum(../INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER&lt;=$PriorSeqNo]/INTERESTRATEPERIOD)"/>
										</xsl:attribute>
										<xsl:attribute name="OverrideValue"><xsl:value-of select="ALPHAPLUSRATE"/></xsl:attribute>
									</xsl:element>
								</xsl:for-each>							
							</xsl:element>
							
							<!-- Increased Interest Rate "table" for 1% increase 
							The 1% increase only applies to the first variable rate, so there will not alway be one -->
							<xsl:if test="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE/ALPHAPLUSINCREASEDRATE">
								<xsl:element name="ParameterOverride">
									<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, '-POIIR-OnePercent')"/></xsl:attribute>
									<xsl:attribute name="Parameter">IncreasedInterestRate</xsl:attribute>
									<xsl:element name="ParameterOverrideKey">
										<xsl:attribute name="Name">Duration</xsl:attribute>
										<xsl:attribute name="NumberOfValues"><xsl:value-of select="count(MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE)"/></xsl:attribute>
									</xsl:element>
									<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE">
										<xsl:sort data-type="number" select="number(INTERESTRATETYPESEQUENCENUMBER)"/>
										<xsl:variable name="PriorSeqNo" select="number(INTERESTRATETYPESEQUENCENUMBER)-1"/>
										<xsl:element name="ParameterOverrideValue">
											<xsl:attribute name="Duration">
												<!-- The duration is the sum of all the periods before the current one -->
												<xsl:value-of select="sum(../INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER&lt;=$PriorSeqNo]/INTERESTRATEPERIOD)"/>
											</xsl:attribute>
											<xsl:attribute name="OverrideValue">
												<xsl:choose>
													<xsl:when test="ALPHAPLUSINCREASEDRATE"><xsl:value-of select="ALPHAPLUSINCREASEDRATE"/></xsl:when>
													<xsl:otherwise><xsl:value-of select="ALPHAPLUSRATE"/></xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
										</xsl:element>
									</xsl:for-each>
								</xsl:element>
							</xsl:if>
							
							<!-- Add IncreaseInterestRate "tables" for min and max values of capped/floored rates if applicable -->
							<xsl:for-each select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[RATETYPE='C']">
								
								<xsl:variable name="CappedSeqNo" select="INTERESTRATETYPESEQUENCENUMBER"/>

								<!-- Increased Interest Rate "table" for floored rates (minimum cost) -->
								<xsl:if test="string(FLOOREDRATE)">
									<xsl:element name="ParameterOverride">
										<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, '-POIIR-Min', $CappedSeqNo)"/></xsl:attribute>
										<xsl:attribute name="Parameter">IncreasedInterestRate</xsl:attribute>
										<xsl:element name="ParameterOverrideKey">
											<xsl:attribute name="Name">Duration</xsl:attribute>
											<xsl:attribute name="NumberOfValues"><xsl:value-of select="count(../INTERESTRATETYPE)"/></xsl:attribute>
										</xsl:element>
										<xsl:for-each select="../INTERESTRATETYPE">
											<xsl:sort data-type="number" select="number(INTERESTRATETYPESEQUENCENUMBER)"/>
											<xsl:variable name="PriorSeqNo" select="number(INTERESTRATETYPESEQUENCENUMBER)-1"/>
											<xsl:element name="ParameterOverrideValue">
												<xsl:attribute name="Duration">
													<!-- The duration is the sum of all the periods before the current one -->
													<xsl:value-of select="sum(../INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER&lt;=$PriorSeqNo]/INTERESTRATEPERIOD)"/>
												</xsl:attribute>
												<xsl:attribute name="OverrideValue">
													<!-- Use the Floored Rate if the current sequence number matches sequence number of the current capped/floored rate, otherwise use the normal rate -->
													<xsl:choose>
														<xsl:when test="INTERESTRATETYPESEQUENCENUMBER=$CappedSeqNo"><xsl:value-of select="FLOOREDRATE"/></xsl:when>
														<xsl:otherwise><xsl:value-of select="ALPHAPLUSRATE"/></xsl:otherwise>
													</xsl:choose>
												</xsl:attribute>
											</xsl:element>
										</xsl:for-each>	
									</xsl:element>
								</xsl:if>

								<!-- Increased Interest Rate "table" for capped rates (maximum cost) -->
								<xsl:if test="string(CEILINGRATE)">
									<xsl:element name="ParameterOverride">
										<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, '-POIIR-Max', $CappedSeqNo)"/></xsl:attribute>
										<xsl:attribute name="Parameter">IncreasedInterestRate</xsl:attribute>
										<xsl:element name="ParameterOverrideKey">
											<xsl:attribute name="Name">Duration</xsl:attribute>
											<xsl:attribute name="NumberOfValues"><xsl:value-of select="count(../INTERESTRATETYPE)"/></xsl:attribute>
										</xsl:element>
										<xsl:for-each select="../INTERESTRATETYPE">
											<xsl:sort data-type="number" select="(INTERESTRATETYPESEQUENCENUMBER)"/>
											<xsl:variable name="PriorSeqNo" select="number(INTERESTRATETYPESEQUENCENUMBER)-1"/>
											<xsl:element name="ParameterOverrideValue">
												<xsl:attribute name="Duration">
													<!-- The duration is the sum of all the periods before the current one -->
													<xsl:value-of select="sum(../INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER&lt;=$PriorSeqNo]/INTERESTRATEPERIOD)"/>
												</xsl:attribute>
												<xsl:attribute name="OverrideValue">
													<!-- Use the Ceiling Rate if the current sequence number matches sequence number of the current capped/floored rate, otherwise use the normal rate -->
													<xsl:choose>
														<xsl:when test="INTERESTRATETYPESEQUENCENUMBER=$CappedSeqNo"><xsl:value-of select="CEILINGRATE"/></xsl:when>
														<xsl:otherwise><xsl:value-of select="ALPHAPLUSRATE"/></xsl:otherwise>
													</xsl:choose>
												</xsl:attribute>
											</xsl:element>
										</xsl:for-each>	
									</xsl:element>
								</xsl:if>
								
							</xsl:for-each>

							<xsl:if test="number($NumRepaymentCharges) &gt; 0">
								<!-- Early Repayment Basis "table" -->
								<xsl:element name="ParameterOverride">
									<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, '-POERB')"/></xsl:attribute>
									<xsl:attribute name="Parameter">EarlyRepaymentBasis</xsl:attribute>
									<xsl:element name="ParameterOverrideKey">
									 	<xsl:attribute name="Name">Duration</xsl:attribute>
									 	<xsl:attribute name="NumberOfValues"><xsl:value-of select="$NumRepaymentCharges"/></xsl:attribute>
									 </xsl:element>
									<xsl:for-each select="MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND">
										<xsl:sort data-type="number" select="number(REDEMPTIONFEESTEPNUMBER)"/>
										<xsl:variable name="PriorStepNo" select="number(REDEMPTIONFEESTEPNUMBER)-1"/>
										<xsl:element name="ParameterOverrideValue">
											<xsl:attribute name="Duration">
												<!-- The duration is the sum of all the periods before the current one -->
												<xsl:value-of select="sum(../REDEMPTIONFEEBAND[REDEMPTIONFEESTEPNUMBER&lt;=$PriorStepNo]/PERIOD)"/>
											</xsl:attribute>
											<xsl:attribute name="OverrideValue">
												<!-- If FEEMONTHSINTEREST has a value output 1 otherwise if FEEPERCENTAGE has a value output 2 -->
												<xsl:if test="string(FEEMONTHSINTEREST)">1</xsl:if>
												<xsl:if test="string(FEEPERCENTAGE)">2</xsl:if>
											</xsl:attribute>
										</xsl:element>
									</xsl:for-each>
								</xsl:element>

								<!-- Early Repayment Charge "table" -->
								<xsl:element name="ParameterOverride">
									<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo, '-POERC')"/></xsl:attribute>
									<xsl:attribute name="Parameter">EarlyRepaymentCharge</xsl:attribute>
									<xsl:element name="ParameterOverrideKey">
									 	<xsl:attribute name="Name">Duration</xsl:attribute>
									 	<xsl:attribute name="NumberOfValues"><xsl:value-of select="$NumRepaymentCharges"/></xsl:attribute>
									 </xsl:element>
									<xsl:for-each select="MORTGAGEPRODUCTDETAILS/REDEMPTIONFEEBANDLIST/REDEMPTIONFEEBAND">
										<xsl:sort data-type="number" select="number(REDEMPTIONFEESTEPNUMBER)"/>
										<xsl:variable name="PriorStepNo" select="number(REDEMPTIONFEESTEPNUMBER)-1"/>
										<xsl:element name="ParameterOverrideValue">
											<xsl:attribute name="Duration">
												<!-- The duration is the sum of all the periods before the current one -->
												<xsl:value-of select="sum(../REDEMPTIONFEEBAND[REDEMPTIONFEESTEPNUMBER&lt;=$PriorStepNo]/PERIOD)"/>
											</xsl:attribute>
											<xsl:attribute name="OverrideValue">
												<!-- Output the value of whichever node has text -->
												<xsl:if test="string(FEEMONTHSINTEREST)"><xsl:value-of select="FEEMONTHSINTEREST"/></xsl:if>
												<xsl:if test="string(FEEPERCENTAGE)"><xsl:value-of select="FEEPERCENTAGE"/></xsl:if>											
											</xsl:attribute>
										</xsl:element>
									</xsl:for-each>
								</xsl:element>
							</xsl:if>

							<!-- SVR comparison -->
							<xsl:if test="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[INTERESTRATEPERIOD='-1']">
								<xsl:element name="ParameterOverride">
									<xsl:attribute name="Id"><xsl:value-of select="concat('EG', $LCSeqNo,'-POIIR-SVR')"/></xsl:attribute>
									<xsl:attribute name="Parameter">IncreasedInterestRate</xsl:attribute>
									<xsl:attribute name="OverrideValue"><xsl:value-of select="MORTGAGEPRODUCTDETAILS/INTERESTRATETYPELIST/INTERESTRATETYPE[INTERESTRATEPERIOD='-1']/ALPHAPLUSRATE"/></xsl:attribute>
								</xsl:element>
							</xsl:if>
							
						</xsl:for-each>
						
						<!-- Fees (One off costs) -->
						<!-- EP1067 GHun OneOffCosts are under MORTGAGEONEOFFCOST when called by RateChange -->
						<xsl:for-each select="($LoanComponents/ONEOFFCOSTLIST/ONEOFFCOST | $LoanComponents/ONEOFFCOSTLIST/MORTGAGEONEOFFCOST)">
							<xsl:variable name="LCSeqNo" select="../../LOANCOMPONENTSEQUENCENUMBER"/>
							<xsl:if test="number(AMOUNT) &gt; 0">							

								<!-- CM100 calls have IDENTIFIER and AMOUNT nodes -->
								<xsl:if test="IDENTIFIER">
									<xsl:if test="IDENTIFIER != 'STA' and IDENTIFIER != 'TID'">
										<xsl:element name="Fee">
											<xsl:attribute name="Amount"><xsl:value-of select="AMOUNT"/></xsl:attribute>
											<xsl:attribute name="Type"><xsl:value-of select="IDENTIFIER"/></xsl:attribute>
											<xsl:attribute name="ChargedToAccount">
												<xsl:choose>
													<xsl:when test="ADDTOLOAN ='1'">ImmediateInterest</xsl:when>
													<xsl:otherwise>Paid</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
											<!-- BMIDS936 GHun -->
											<xsl:if test="PAIDATEND='1'">
												<xsl:attribute name="WhenFeeApplied">End</xsl:attribute>
											</xsl:if>	
											<xsl:attribute name="ElementIds"><xsl:value-of select="concat('EG', $LCSeqNo, 'Element1')"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
								</xsl:if>
								
								<!-- CM130 calls have COMBOVALIDATIONTYPE, NAME, AMOUNT and ADDEDTOLOAN nodes -->
								<xsl:if test="COMBOVALIDATIONTYPE">
									<xsl:if test="COMBOVALIDATIONTYPE!= 'STA' and COMBOVALIDATIONTYPE!= 'TID'">
										<xsl:element name="Fee">
											<xsl:attribute name="Amount"><xsl:value-of select="AMOUNT"/></xsl:attribute>
											<xsl:attribute name="Type"><xsl:value-of select="COMBOVALIDATIONTYPE"/></xsl:attribute>
											<xsl:attribute name="ChargedToAccount">
												<xsl:choose>
													<xsl:when test="ADDEDTOLOAN ='1'">ImmediateInterest</xsl:when>
													<xsl:otherwise>Paid</xsl:otherwise>
												</xsl:choose>
											</xsl:attribute>
											<!-- BMIDS936 GHun -->
											<xsl:if test="PAIDATEND='1'">
												<xsl:attribute name="WhenFeeApplied">End</xsl:attribute>
											</xsl:if>									
											<xsl:attribute name="ElementIds"><xsl:value-of select="concat('EG', $LCSeqNo, 'Element1')"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
								</xsl:if>								
							</xsl:if>
						</xsl:for-each>
						
						<!-- End Mortgage Element -->
					</xsl:element>
					<!-- End Control Element -->
				</xsl:element>
				<!-- End Inputs Element -->
			</xsl:element>
			<!-- End Request Element -->
		</xsl:element>
	</xsl:template>
	
	<!-- Create ParameterOverrides for each MortgageLenderParameter -->
	<xsl:template match="MORTGAGELENDERPARAMETERS">
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride01</xsl:attribute>
			<xsl:attribute name="Parameter">MonthsToCompletion</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="SHIFTINMONTHS"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride02</xsl:attribute>
			<xsl:attribute name="Parameter">CompletionTiming</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="POSITIONINMONTH"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride03</xsl:attribute>
			<xsl:attribute name="Parameter">CompletionDaysAdjust</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="SHIFTINDAYS"/></xsl:attribute>
		</xsl:element>
		<xsl:variable name="IsPaymentTimingZero" select="number(FIRSTPAYMENTDUEDATEIND)=0"></xsl:variable>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride04</xsl:attribute>
			<xsl:attribute name="Parameter">PaymentTiming</xsl:attribute>
			<xsl:attribute name="OverrideValue">
				<xsl:choose>
					<xsl:when test="$IsPaymentTimingZero">0</xsl:when>
					<xsl:otherwise>1</xsl:otherwise>
				</xsl:choose>
			</xsl:attribute>
		</xsl:element>
		<!-- CoolingOffPeriod and SpecifiedPaymentDay are not required when PaymentTiming = 0 -->
		<xsl:if test="not($IsPaymentTimingZero)">
			<xsl:element name="ParameterOverride">
				<xsl:attribute name="Id">ParameterOverride05</xsl:attribute>
				<xsl:attribute name="Parameter">CoolingOffPeriod</xsl:attribute>
				<xsl:attribute name="OverrideValue"><xsl:value-of select="COOLINGOFFPERIOD"/></xsl:attribute>
			</xsl:element>
			<xsl:element name="ParameterOverride">
				<xsl:attribute name="Id">ParameterOverride06</xsl:attribute>
				<xsl:attribute name="Parameter">SpecifiedPaymentDay</xsl:attribute>
				<xsl:attribute name="OverrideValue"><xsl:value-of select="SPECIFIEDPAYMENTDAY"/></xsl:attribute>
			</xsl:element>
		</xsl:if>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride07</xsl:attribute>
			<xsl:attribute name="Parameter">FirstMonth</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCOUNTINGSTARTMONTH"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride08</xsl:attribute>
			<xsl:attribute name="Parameter">DaysInYear</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="DAYSINYEAR"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride09</xsl:attribute>
			<xsl:attribute name="Parameter">AccruedBasis</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCINTERESTIND"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride10</xsl:attribute>
			<xsl:attribute name="Parameter">DaysIncludedInAccrued</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCRUEDDAYSINCLUDED"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride11</xsl:attribute>
			<xsl:attribute name="Parameter">AddedDaysAccrued</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCRUEDDAYSADDED"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride12</xsl:attribute>
			<xsl:attribute name="Parameter">AccruedTiming</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCINTERESTPAYABLEIND"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride13</xsl:attribute>
			<xsl:attribute name="Parameter">InterestChargeMonth</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="MONTHLYPAYMENTIND"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride14</xsl:attribute>
			<xsl:attribute name="Parameter">InterestChargeBasis</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="INTERESTCHARGEDIND"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride15</xsl:attribute>
			<xsl:attribute name="Parameter">APRMonth</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="APRMONTH"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride16</xsl:attribute>
			<xsl:attribute name="Parameter">WhenAccIntRounded</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="WHENACCINTROUNDED"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride17</xsl:attribute>
			<xsl:attribute name="Parameter">AccIntRoundingMultiple</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCINTERESTROUNDINGFACTOR"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride18</xsl:attribute>
			<xsl:attribute name="Parameter">AccIntRoundingDirection</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="ACCINTERESTROUNDINGDIRECTION"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride19</xsl:attribute>
			<xsl:attribute name="Parameter">WhenPaymentRounded</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="WHENPAYMENTROUNDED"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride20</xsl:attribute>
			<xsl:attribute name="Parameter">PaymentRoundingMultiple</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="PAYMENTROUNDINGFACTOR"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride21</xsl:attribute>
			<xsl:attribute name="Parameter">PaymentRoundingDirection</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="PAYMENTROUNDINGDIRECTION"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride22</xsl:attribute>
			<xsl:attribute name="Parameter">WhenBalanceRounded</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="WHENBALANCEROUNDED"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride23</xsl:attribute>
			<xsl:attribute name="Parameter">BalanceRoundingMultiple</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="BALANCEROUNDINGFACTOR"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride24</xsl:attribute>
			<xsl:attribute name="Parameter">BalanceRoundingDirection</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="BALANCEROUNDINGDIRECTION"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride25</xsl:attribute>
			<xsl:attribute name="Parameter">ChargeRoundingMultiple</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="CHARGEROUNDINGFACTOR"/></xsl:attribute>
		</xsl:element>
		<xsl:element name="ParameterOverride">
			<xsl:attribute name="Id">ParameterOverride26</xsl:attribute>
			<xsl:attribute name="Parameter">ChargeRoundingDirection</xsl:attribute>
			<xsl:attribute name="OverrideValue"><xsl:value-of select="CHARGEROUNDINGDIRECTION"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!-- End of MortgageLenderParameters template -->
</xsl:stylesheet>
