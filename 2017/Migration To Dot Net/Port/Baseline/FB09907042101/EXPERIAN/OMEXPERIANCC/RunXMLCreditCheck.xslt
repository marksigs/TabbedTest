<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: RunXMLCreditCheck

History:

Version Author		Date        Description
01.00   RFairlie	10/01/2005	MAR1008		Correct setting of RESIDSTATUS.
01.01	PCarter		17/01/2006	MAR1059		Change ToDate in CreateBlankSeqBlocks to be blank
01.02	GHun		20/01/2006	MAR1096		Don't set TYPE when creating a blank NAME block
01.03	RFairlie	09/02/2006	MAR1008		Correct setting of RESIDSTATUS - pick up TENANCY correctly and check after nature of occupancy.
01.04	JD			23/02/2006	MAR1180		If confirmed income is present use it as QualifyingRegularIncome. LDM changes not needed for Epsom
01.05	JD			24/02/2006	MAR1177 	set SEARCHFLAG purely based on global param, not address. LDM changes not needed for Epsom
01.06	JD			28/02/2006	MAR1175 	corrected path to ALIAS on set up of NAMESCOUNT.  LDM changes not needed for Epsom as already fixed
01.07   HMA			02/03/2006  MAR1333 	Allow for unknown gender
01.08	GHun		07/03/2006	MAR1372		Set AC01 function to reprocess if a 2nd initial credit check is run.  LDM changes not needed for Epsom
01.09	LDM			15/02/2006	EP6			New layout for for Epsom call. RESIDSTATUS changed to be more like 1.02. One shot is not now supported.
01.10	SAB			29/03/2006	EP306		Overrides the EARNEDINCOME total with INCOMESUMMARY.ALLOWABLEANNUALINCOME ID 
01.11	SAB			04/04/2006	EP321		Restricts the length of telephone numbers and correctly sets the default value for EMAILADDR to 'Z'
01.12	LDM			20/04/2006  EP358		If have a case with joint app but only joint app has a alias, pad the name block but do not pad the resy block
01.13	SAB			25/04/2006	EP437		Set GROSSANNUALINC correctly when ALLOWABLEANNUALINCOME is used
01.14	IK			25/04/2006	EP469,EP470	additional blocks for re-score & re-process when expired
01.15	IK			10/05/2006	EP530		fix EP469,EP470
01.16   LDM			15/05/2006  EP368		RESY Date data was incorrect in the Experian call. If the joint applicants current address was the same as the main applicant went wrong
01.17	LDM			25/05/2006  EP618		Force the address details by truncating, to be the right size for experian
01.18	IK			30/05/2006	EP480		set SEARCHTYPE depending on stage
01.19   LDM			31/05/2006  EP627		Apply mars changes to Epsom cut. (02.03	PEdney	28/03/2006	MAR1395) & (02.06	GHun	07/04/2006	MAR1409 Added some validation)
01.20   LDM			31/05/2006  EP635		Take out the validation on Term
01.21   LDM			01/06/2006  EP649		truncate title to 10, send a blank suffix
01.22	PB			07/06/2006	EP696		MAR1803 Experian CreditCheck must validate MaritalStatus
01.23   LDM			14/09/2006  EP1121    Make sure the correct blank name blocks are created for aliases
01.24   LDM			06/02/2007  EP2_750  Get dependant information from application node
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="Combo" select="/RESPONSE/COMBOVALUE"/>
	<xsl:variable name="GlobalParam" select="/RESPONSE/GLOBALPARAMETER"/>
	<xsl:variable name="Customer" select="/RESPONSE/CUSTOMER"/>
	<xsl:variable name="IsRescore" select="/RESPONSE/@RESCORE"/>
	<xsl:variable name="IsReprocess" select="/RESPONSE/@REPROCESS"/>
	<xsl:variable name="IsOneShot" select="/RESPONSE/@ONESHOT"/>
	<xsl:variable name="AppCreditCheck" select="/RESPONSE/APPLICATIONCREDITCHECK"/>
	<xsl:variable name="TotAliasAssoc1">
		<xsl:variable name="TotAliasAssoc1NoLimit" select="count(/RESPONSE/CUSTOMER[1]/ALIAS)"/>
		<xsl:choose>
			<xsl:when test="$TotAliasAssoc1NoLimit &gt; 3">3</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$TotAliasAssoc1NoLimit"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="TotAliasAssoc2">
		<xsl:variable name="TotAliasAssoc2NoLimit" select="count(/RESPONSE/CUSTOMER[2]/ALIAS)"/>
		<xsl:choose>
			<xsl:when test="$TotAliasAssoc2NoLimit &gt; 3">3</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$TotAliasAssoc2NoLimit"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
	<xsl:variable name="MaxAliasAssoc">
		<xsl:choose>
			<xsl:when test="number($TotAliasAssoc1) > number($TotAliasAssoc2)">
				<xsl:value-of select="$TotAliasAssoc1"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="$TotAliasAssoc2"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>

	<xsl:template match="/">
		<xsl:element name="EXPERIAN">
			<xsl:call-template name="Validation"/>
			<xsl:element name="ESERIES">
				<xsl:element name="FORM_ID">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianCCFormId']/@STRING"/>
				</xsl:element>
			</xsl:element>
			<xsl:variable name="Function">
				<xsl:choose>
					<xsl:when test="$AppCreditCheck">
						<xsl:choose>
							<xsl:when test="number($AppCreditCheck[1]/@DAYSSINCECHECK) &gt; number($GlobalParam[@NAME='CreditCheckPostArchiveDay']/@AMOUNT)">
								<xsl:value-of select="'0010'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:choose>
									<xsl:when test="$IsReprocess">0020</xsl:when>
									<xsl:otherwise>
										<xsl:choose>
											<xsl:when test="$IsRescore">0025</xsl:when>
											<xsl:otherwise>0010</xsl:otherwise>
										</xsl:choose>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>0010</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:if test="not($Function='0010')">
				<xsl:element name="EXP">
					<xsl:element name="EXPERIANREF">
							<xsl:value-of select="/RESPONSE/APPLICATIONCREDITCHECK[1]/@CREDITCHECKREFERENCENUMBER"/>
					</xsl:element>
				</xsl:element>
			</xsl:if>	
			<xsl:element name="AC01">
				<xsl:element name="CLIENTKEY">
					<xsl:choose>
						<xsl:when test="$IsOneShot">
							<xsl:value-of select="$GlobalParam[@NAME='ExperianCCClientKeyOneShot']/@STRING"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$GlobalParam[@NAME='ExperianCCClientKey']/@STRING"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:element>
				<xsl:element name="ENTRYPOINT">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianCCEntryPoint']/@STRING"/>
				</xsl:element>
				<xsl:element name="FUNCTION">
					<xsl:value-of select="$Function"/>
				</xsl:element>
			</xsl:element>
			<xsl:if test="$Function='0010'">
				<xsl:call-template name="TDP1block"/>
			</xsl:if>
			<xsl:if test="$Function='0010' or $Function='0020'">
				<xsl:call-template name="customerBlock"/>
			</xsl:if>
			<xsl:element name="AP19">
				<xsl:element name="AMOUNT">
					<xsl:variable name="Amount" select="number(/RESPONSE/LOANCOMPONENT/@AMOUNTREQUESTED)"/>
					<xsl:if test="$Amount">
						<xsl:value-of select="number($Amount)"/>
					</xsl:if>	
				</xsl:element>
				<xsl:element name="TERM">
					<xsl:variable name="Term" select="number(/RESPONSE/LOANCOMPONENT/@TERMINYEARS)"/>
					<xsl:if test="$Term">
						<xsl:value-of select="number($Term)"/>
					</xsl:if>	
				</xsl:element>
				<xsl:element name="APPTYPE">MG</xsl:element>
				<xsl:element name="PROPERTYVALUE">
					<xsl:variable name="Price" select="number(/RESPONSE/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"/>
					<xsl:if test="$Price">
						<xsl:value-of select="number($Price)"/>
					</xsl:if>	
					<xsl:if test="not($Price)">0</xsl:if>	
				</xsl:element>
				<xsl:element name="MORTGAGETYPE">Z</xsl:element>
			</xsl:element>
			<xsl:for-each select="$Customer">
				<xsl:element name="AM19">
					<xsl:call-template name="MaritalStatus">
						<xsl:with-param name="Value" select="CUSTOMERVERSION/@MARITALSTATUS"/>
					</xsl:call-template>
					<xsl:element name="DEPENDANTS">
						<xsl:call-template name="GetDependants"><!--EP2_750 -->
							<xsl:with-param name="CustPosition" select="position()"/>
						</xsl:call-template>					
					</xsl:element>
					<xsl:element name="RESIDSTATUS">
						<xsl:variable name="CurrentAddressType" select="$Combo[@GROUPNAME='CustomerAddressType'][COMBOVALIDATION[@VALIDATIONTYPE='H']]/@VALUEID"/>
						<xsl:variable name="CurrentAddress" select="ADDRESS[@ADDRESSTYPE=$CurrentAddressType]"/>
						<xsl:choose>
							<xsl:when test="not($CurrentAddress)">Z</xsl:when>
							<xsl:otherwise>
								<xsl:variable name="NatureOfOccupancy" select="ADDRESS[@ADDRESSTYPE=$CurrentAddressType]/@NATUREOFOCCUPANCY"/>
								<xsl:variable name="TenancyType" select="CURRENTADDRESS_TENANCY/@TENANCYTYPE"/>
								<xsl:variable name="NOCCombo" select="$Combo[@GROUPNAME='NatureOfOccupancy' and @VALUEID=$NatureOfOccupancy]/COMBOVALIDATION"/>
								<xsl:choose>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='O']">O</xsl:when>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='LP']">P</xsl:when>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='R']">X</xsl:when>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='RV']">
									<xsl:choose>
										<xsl:when test="$TenancyType">
											<xsl:choose>
												<xsl:when test="$Combo[@GROUPNAME='LandlordType' and @VALUEID=$TenancyType]/COMBOVALIDATION[@VALIDATIONTYPE='L']">C</xsl:when>
												<xsl:when test="$Combo[@GROUPNAME='LandlordType' and @VALUEID=$TenancyType]/COMBOVALIDATION[@VALIDATIONTYPE='P']">T</xsl:when>
												<xsl:otherwise>X</xsl:otherwise>
											</xsl:choose>
										</xsl:when>
										<xsl:otherwise>T</xsl:otherwise>
									</xsl:choose>
									</xsl:when>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='RB']">
									<xsl:choose>
										<xsl:when test="$TenancyType">
											<xsl:choose>
												<xsl:when test="$Combo[@GROUPNAME='LandlordType' and @VALUEID=$TenancyType]/COMBOVALIDATION[@VALIDATIONTYPE='L']">C</xsl:when>
												<xsl:when test="$Combo[@GROUPNAME='LandlordType' and @VALUEID=$TenancyType]/COMBOVALIDATION[@VALIDATIONTYPE='P']">T</xsl:when>
												<xsl:otherwise>X</xsl:otherwise>
											</xsl:choose>
										</xsl:when>
										<xsl:otherwise>C</xsl:otherwise>
									</xsl:choose>
									</xsl:when>
									<xsl:when test="$NOCCombo[@VALIDATIONTYPE='X']">X</xsl:when>
									<xsl:otherwise>Z</xsl:otherwise>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					
					<xsl:element name="CURRACCNTHELD">
						<xsl:choose>
							<xsl:when test="BANKCREDITCARD[@CARDTYPE=$Combo[@GROUPNAME='CreditCardType'][COMBOVALIDATION[@VALIDATIONTYPE='CA']]/@VALUEID]">Y</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="TIMEWITHBANK">
						<xsl:choose>
							<xsl:when test="string(CUSTOMERVERSION/@TIMEATBANKMONTHS) or string(CUSTOMERVERSION/@TIMEATBANKYEARS)">
								<xsl:choose>
									<xsl:when test="string(number(CUSTOMERVERSION/@TIMEATBANKYEARS))='NaN'">00</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="format-number(number(CUSTOMERVERSION/@TIMEATBANKYEARS), '00')"/>
									</xsl:otherwise>
								</xsl:choose>
								<xsl:choose>
									<xsl:when test="string(number(CUSTOMERVERSION/@TIMEATBANKMONTHS))='NaN'">00</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="format-number(number(CUSTOMERVERSION/@TIMEATBANKMONTHS), '00')"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise/>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="CHEQUECARDHELD">
						<xsl:choose>
							<xsl:when test="BANKCREDITCARD[@CARDTYPE=$Combo[@GROUPNAME='CreditCardType'][COMBOVALIDATION[@VALIDATIONTYPE='C']]/@VALUEID]">Y</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:variable name="EmploymentStatus" select="EMPLOYMENT/@EMPLOYMENTSTATUS"/>
					<xsl:variable name="OccupationType" select="EMPLOYMENT/@OCCUPATIONTYPE"/>
					<xsl:variable name="OTCombo" select="$Combo[@GROUPNAME='OccupationType' and @VALUEID=$OccupationType]/COMBOVALIDATION"/>
					<xsl:variable name="ESCombo" select="$Combo[@GROUPNAME='EmploymentStatus' and @VALUEID=$EmploymentStatus]/COMBOVALIDATION"/>
					<xsl:variable name="DerivedEmpStatus">
						<xsl:choose>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='SU']">O</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='U']">N</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='MG']">T</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='P']">M</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='S']">S</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='SS']">P</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='J']">J</xsl:when>
							<xsl:when test="$OTCombo[@VALIDATIONTYPE='O']">X</xsl:when>
							<xsl:otherwise>Z</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<xsl:element name="EMPSTATUS">
						<xsl:value-of select="$DerivedEmpStatus"/>
					</xsl:element>
					<xsl:element name="EMPTYPE">
						<!--<xsl:variable name="EmploymentType" select="EMPLOYMENT/@EMPLOYMENTTYPE"/>
						<xsl:variable name="ETCombo" select="$Combo[@GROUPNAME='EmploymentType' and @VALUEID=$EmploymentType]/COMBOVALIDATION"/>-->
						<xsl:choose>
							<!--<xsl:when test="($ESCombo[@VALIDATIONTYPE='EMP'] and $ETCombo[@VALIDATIONTYPE='ET']) or $ESCombo[@VALIDATIONTYPE='TP']">T</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='EMP'] and $ETCombo[@VALIDATIONTYPE='PF']">E</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='EMP'] and $ETCombo[@VALIDATIONTYPE='PP']">L</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='SELF'] and $DerivedEmpStatus='M'">P</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='SELF'] and not($DerivedEmpStatus='M')">N</xsl:when>-->
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='CON']">T</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='EMP']">E</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='SELF']">P</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='STU']">S</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='HM']">H</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='R']">R</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='N']">U</xsl:when>
							<xsl:when test="$ESCombo[@VALIDATIONTYPE='O']">X</xsl:when>
							<xsl:otherwise>Z</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="TIMEWITHEMP">
						<xsl:choose>
							<xsl:when test="not(EMPLOYMENT/@DATELEFTORCEASEDTRADING) and not(string(EMPLOYMENT/@DATESTARTEDORESTABLISHED))">
								<xsl:value-of select="EMPLOYMENT/@TIMEWITHEMPLOYER"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:text> </xsl:text>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="NUMINHOUSEHOLD">
						<xsl:value-of select="/RESPONSE/APPLICATION/@NUMINHOUSEHOLD"/>
					</xsl:element>
					<xsl:element name="CITIZENSHIP"/>
					<xsl:element name="CURRADDRTYPE"/>
					<xsl:element name="PREVADDRTYPE"/>
					<xsl:element name="PREVPREVADDRTYPE"/>
					<xsl:element name="HOMETELSTD">
						<xsl:choose>
							<xsl:when test="HOMETELEPHONE/@AREACODE">
								<xsl:value-of select="substring(HOMETELEPHONE/@AREACODE, 1, 6)"/>
							</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="HOMETELLOCAL">
						<xsl:value-of select="substring(HOMETELEPHONE/@TELEPHONENUMBER, 1, 10)"/>
					</xsl:element>
					<xsl:call-template name="CardsHeld">
						<xsl:with-param name="Number" select="1"/>
					</xsl:call-template>
					<xsl:element name="WORKSTD">
						<xsl:choose>
							<xsl:when test="WORKTELEPHONE/@AREACODE">
								<xsl:value-of select="substring(WORKTELEPHONE/@AREACODE, 1, 6)"/>
							</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="WORKLOCAL">
						<xsl:value-of select="substring(WORKTELEPHONE/@TELEPHONENUMBER, 1, 10)"/>
					</xsl:element>
					<xsl:element name="GROSSANNUALINC">
						<xsl:call-template name="SumBasicAnnualIncome">
							<xsl:with-param name="Income" select="EARNEDINCOME"/>
							<xsl:with-param name="AllowableIncome" select="INCOMESUMMARY/@ALLOWABLEANNUALINCOME"/>
							<xsl:with-param name="RunningTotal" select="'0'"/>
						</xsl:call-template>
					</xsl:element>
					<xsl:element name="MOBILETELCODE">
						<xsl:choose>
							<xsl:when test="MOBILETELEPHONE/@AREACODE">
								<xsl:value-of select="substring(MOBILETELEPHONE/@AREACODE, 1, 6)"/>
							</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="MOBILETELNUM">
						<xsl:choose>
							<xsl:when test="MOBILETELEPHONE/@TELEPHONENUMBER">
								<xsl:value-of select="substring(MOBILETELEPHONE/@TELEPHONENUMBER, 1, 10)"/>
							</xsl:when>
							<xsl:otherwise>N</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="EMAILADDR">
						<xsl:choose>
							<xsl:when test="CUSTOMERVERSION/@CONTACTEMAILADDRESS">
								<xsl:value-of select="substring(CUSTOMERVERSION/@CONTACTEMAILADDRESS, 1, 60)"/>
							</xsl:when>
							<xsl:otherwise>Z</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
					<xsl:element name="DRIVINGLICENCE">Q</xsl:element>
				</xsl:element>
			</xsl:for-each>
			<xsl:variable name="AgeMain" select="/RESPONSE/CUSTOMER[1]/CUSTOMERVERSION/@AGE"/>
			<xsl:variable name="AgeJoint" select="/RESPONSE/CUSTOMER[2]/CUSTOMERVERSION/@AGE"/>
			<xsl:variable name="TimeAtAddrMain" select="/RESPONSE/CUSTOMER[1]/ADDRESS/@TIMEATADDRESS"/>
			<xsl:variable name="TimeAtAddrJoint" select="/RESPONSE/CUSTOMER[2]/ADDRESS/@TIMEATADDRESS"/>
			<!-- Only create one AM90 block -->
			<xsl:element name="AM90">
				<xsl:variable name="NatureOfLoan" select="/RESPONSE/APPLICATIONFACTFIND/@NATUREOFLOAN"/>
				<xsl:variable name="NatureCombo" select="$Combo[@GROUPNAME='NatureOfLoan' and @VALUEID=$NatureOfLoan]/COMBOVALIDATION"/>
				<xsl:variable name="NatureOfLoanStatus">
					<xsl:choose>
						<xsl:when test="$NatureCombo[@VALIDATIONTYPE='RS']">R</xsl:when>
						<xsl:when test="$NatureCombo[@VALIDATIONTYPE='LT']">L</xsl:when>
						<xsl:when test="$NatureCombo[@VALIDATIONTYPE='BI']">B</xsl:when>
						<xsl:when test="$NatureCombo[@VALIDATIONTYPE='BR']">B</xsl:when>
					</xsl:choose>
				</xsl:variable>
				<xsl:element name="PRODUCTTYPE">
					<xsl:value-of select="$NatureOfLoanStatus"/>
				</xsl:element>
				<xsl:variable name="BuyerType" select="/RESPONSE/APPLICATION/@TYPEOFBUYER"/>
				<xsl:variable name="BuyerTypeCombo" select="$Combo[@GROUPNAME='TypeOfBuyerNewLoan' and @VALUEID=$BuyerType]/COMBOVALIDATION"/>
				<xsl:variable name="BuyerTypeStatus">
					<xsl:choose>
						<xsl:when test="$BuyerTypeCombo[@VALIDATIONTYPE='F']">F</xsl:when>
						<xsl:when test="$BuyerTypeCombo[@VALIDATIONTYPE='S']">H</xsl:when>
						<xsl:when test="$BuyerTypeCombo[@VALIDATIONTYPE='R']">R</xsl:when>
					</xsl:choose>
				</xsl:variable>
				<xsl:element name="BUYERTYPE">
					<xsl:value-of select="$BuyerTypeStatus"/>
				</xsl:element>
				<xsl:call-template name="LTV_Calc">
					<xsl:with-param name="AmountReq" select="/RESPONSE/LOANCOMPONENT/@AMOUNTREQUESTED"/>
					<xsl:with-param name="PurchPrice" select="/RESPONSE/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE"/>
				</xsl:call-template>
				<xsl:element name="AGEMAINAPP">
					<xsl:value-of select="$AgeMain"/>
				</xsl:element>
				<xsl:element name="AGEJOINTAPP">
					<xsl:value-of select="$AgeJoint"/>
				</xsl:element>
				<xsl:element name="TIMEADDRMAIN">
					<xsl:value-of select="$TimeAtAddrMain"/>
				</xsl:element>
				<xsl:element name="TIMEADDRJOINT">
					<xsl:value-of select="$TimeAtAddrJoint"/>
				</xsl:element>
				<xsl:element name="SELFCERTFLAG">
					<xsl:choose>
						<xsl:when test="/RESPONSE/APPLICATIONFACTFIND/@SELFCERTIND='1'">Y</xsl:when>
						<xsl:otherwise>N</xsl:otherwise>
					</xsl:choose>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateName">
		<xsl:param name="XML"/>
		<xsl:param name="IsAliasAssoc"/>
		<xsl:param name="DateOfBirthXML"/>
		<xsl:element name="NAME">
			<xsl:element name="TYPE">
				<xsl:choose>
					<xsl:when test="count($XML/@*) = 0"/>
					<xsl:when test="$IsAliasAssoc='AL'">
						<xsl:choose>
							<xsl:when test="$XML/../../@CUSTOMERORDER='1'">2</xsl:when>
							<xsl:otherwise>6</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$IsAliasAssoc='AS'">
						<xsl:choose>
							<xsl:when test="$XML/../../@CUSTOMERORDER='1'">3</xsl:when>
							<xsl:otherwise>7</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="$XML/../@CUSTOMERORDER='1'">1</xsl:when>
							<xsl:otherwise>5</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<!--	<xsl:element name="CODE"/> -->
			<xsl:element name="DATEOFBIRTH_DD">
				<xsl:choose>
					<xsl:when test="not($IsAliasAssoc) or $IsAliasAssoc='AL'">
						<xsl:value-of select="substring($DateOfBirthXML/@DATEOFBIRTH, 9, 2)"/>
					</xsl:when>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="DATEOFBIRTH_MM">
				<xsl:choose>
					<xsl:when test="not($IsAliasAssoc) or $IsAliasAssoc='AL'">
						<xsl:value-of select="substring($DateOfBirthXML/@DATEOFBIRTH, 6, 2)"/>
					</xsl:when>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="DATEOFBIRTH_CCYY">
				<xsl:choose>
					<xsl:when test="not($IsAliasAssoc) or $IsAliasAssoc='AL'">
						<xsl:value-of select="substring($DateOfBirthXML/@DATEOFBIRTH, 1, 4)"/>
					</xsl:when>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="TITLE">
				<xsl:value-of select="substring($XML/@TITLE, 1, 10)"/>
			</xsl:element>
			<xsl:element name="FORENAME">
				<xsl:value-of select="substring($XML/@FIRSTFORENAME, 1, 15)"/>
			</xsl:element>
			<xsl:element name="INITIALS">
				<xsl:value-of select="substring($XML/@SECONDFORENAME, 1, 15)"/>
			</xsl:element>
			<xsl:element name="SURNAME">
				<xsl:value-of select="substring($XML/@SURNAME, 1, 30)"/>
			</xsl:element>
			<xsl:element name="SUFFIX"/>
			<xsl:element name="SEX">
				<xsl:value-of select="translate($XML/@GENDER,'124', 'MF')"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateBlankName">
		<xsl:element name="NAME">
			<xsl:element name="TYPE"/>
			<xsl:element name="DATEOFBIRTH_DD"/>
			<xsl:element name="DATEOFBIRTH_MM"/>
			<xsl:element name="DATEOFBIRTH_CCYY"/>
			<xsl:element name="TITLE"/>
			<xsl:element name="FORENAME"/>
			<xsl:element name="INITIALS"/>
			<xsl:element name="SURNAME"/>
			<xsl:element name="SUFFIX"/>
			<xsl:element name="SEX"/>
		</xsl:element>
	</xsl:template>
	<!-- Create NAME records for aliases. Experian requires blank NAME records to be generated when no data exists -->
	<xsl:template name="CreateAliasName">
		<xsl:param name="MaxAlias"/>
		<xsl:param name="MaxAliasApplicant"/><!--Applicant with the most aliases(1or2), if they have the same each set to 2  -->
		<xsl:param name="CurrentApplicant"/>
		<xsl:param name="CurrentAlias"/>
		<xsl:variable name="XML" select="$Customer[$CurrentApplicant]/ALIAS[string(@ALIASSEQUENCENUMBER)][$CurrentAlias]/PERSON"/>
		<!-- Create the name block -->		
		<xsl:if test="count($XML/@*) = 0">
			<xsl:call-template name="CreateBlankName"/>
		</xsl:if>
		<xsl:if test="count($XML/@*) &gt; 0">
			<xsl:call-template name="CreateName">
				<xsl:with-param name="XML" select="$XML"/>
				<xsl:with-param name="IsAliasAssoc" select="$Customer[$CurrentApplicant]/ALIAS[string(@ALIASSEQUENCENUMBER)][$CurrentAlias]/@ALIASTYPE"/>
				<xsl:with-param name="DateOfBirthXML" select="$Customer[$CurrentApplicant]/CUSTOMERVERSION"/>
			</xsl:call-template>
		</xsl:if>
		<!-- we have created the name block by here; be it real or padded -->
		<xsl:choose>
		<xsl:when test="$CurrentApplicant = 1 and ($CurrentAlias &lt; $MaxAlias)">
		<!-- we have just delt with the 1st applicant, by creating a name block for this alias seq no. -->
		<!--If the current alias seq no is not the maximum alias seq no. Then need a name block for applicatnt 2, be it real or padded for this alias seq no-->
			<xsl:call-template name="CreateAliasName">
				<xsl:with-param name="MaxAlias" select="$MaxAlias"/>
				<xsl:with-param name="MaxAliasApplicant" select="$MaxAliasApplicant"/>
				<xsl:with-param name="CurrentApplicant" select="2"/>
				<xsl:with-param name="CurrentAlias" select="$CurrentAlias"/>
			</xsl:call-template>
		</xsl:when>
		<xsl:when test="$CurrentApplicant = 2 and ($CurrentAlias &lt; $MaxAlias)">
		<!-- we have just delt with the 2nd applicant, by creating a name block for this alias seq no.-->
		<!-- If the current alias seq no is not the maximum alias seq nor. Then need a name block for applicatnt 1 be it real or padded for the next alias seq no-->
			<xsl:call-template name="CreateAliasName">
				<xsl:with-param name="MaxAlias" select="$MaxAlias"/>
				<xsl:with-param name="MaxAliasApplicant" select="$MaxAliasApplicant"/>
				<xsl:with-param name="CurrentApplicant" select="1"/>
				<xsl:with-param name="CurrentAlias" select="$CurrentAlias + 1"/>
			</xsl:call-template>
		</xsl:when>
		<xsl:when test=" ($CurrentAlias = $MaxAlias) and $CurrentApplicant = 1 and $MaxAliasApplicant=2">
		<!-- we have just delt with the 1st applicant last alias(real or padded), by creating a name block for this alias seq no.-->
		<!-- If 2nd applicant has the most or the same number of aliases as the 1st applicant. Then create an name block for it, otherwise drop out -->
			<xsl:call-template name="CreateAliasName">
				<xsl:with-param name="MaxAlias" select="$MaxAlias"/>
				<xsl:with-param name="MaxAliasApplicant" select="$MaxAliasApplicant"/>
				<xsl:with-param name="CurrentApplicant" select="2"/>
				<xsl:with-param name="CurrentAlias" select="$CurrentAlias"/>
			</xsl:call-template>
		</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="MaritalStatus">
		<xsl:param name="Value"/>
		<xsl:variable name="MSCombo" select="$Combo[@GROUPNAME='MaritalStatus']"/>
		<xsl:element name="MARITALSTATUS">
			<xsl:choose>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='M']">M</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='S']">S</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='D']">D</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='W']">W</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='E']">E</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='C']">C</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='P']">X</xsl:when>
				<xsl:when test="$MSCombo[@VALUEID=$Value]/COMBOVALIDATION[@VALIDATIONTYPE='O']">O</xsl:when>
				<xsl:otherwise>Z</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CardsHeld">
		<xsl:param name="Number"/>
		<xsl:element name="{concat('CARDSHELD', $Number)}">
			<xsl:variable name="CardType" select="BANKCREDITCARD[$Number]/@CARDTYPE"/>
			<xsl:choose>
				<xsl:when test="not($CardType)">N</xsl:when>
				<xsl:otherwise>
					<xsl:variable name="CCTCombo" select="$Combo[@GROUPNAME='CreditCardType' and @VALUEID=$CardType]/COMBOVALIDATION"/>
					<xsl:choose>
						<xsl:when test="$CCTCombo[@VALIDATIONTYPE='VC']">V</xsl:when>
						<xsl:when test="$CCTCombo[@VALIDATIONTYPE='MC']">M</xsl:when>
						<xsl:when test="$CCTCombo[@VALIDATIONTYPE='SC']">S</xsl:when>
						<xsl:when test="$CCTCombo[@VALIDATIONTYPE='DC']">D</xsl:when>
						<xsl:when test="$CCTCombo[@VALIDATIONTYPE='AC']">A</xsl:when>
						<xsl:otherwise>O</xsl:otherwise>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
		<xsl:if test="$Number &lt; 6">
			<xsl:call-template name="CardsHeld">
				<xsl:with-param name="Number" select="number($Number) + 1"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateAddresses">
		<!-- 1st applicant, current address -->
		<xsl:apply-templates select="$Customer[1]/ADDRESS[1]"/>
		<!-- 1st applicant, 1st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[1]/ADDRESS[2]">
				<xsl:apply-templates select="$Customer[1]/ADDRESS[2]"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankUKAddress"/>
			</xsl:when>
		</xsl:choose>
		<!-- 1st applicant, 2st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[1]/ADDRESS[3]">
				<xsl:apply-templates select="$Customer[1]/ADDRESS[3]"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankUKAddress"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant, current address -->
		<xsl:choose>
			<xsl:when test="$Customer[2]/ADDRESS[1][not(@DUPLICATE)]">
				<xsl:apply-templates select="$Customer[2]/ADDRESS[1][not(@DUPLICATE)]"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankUKAddress"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant,  1st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[2]/ADDRESS[2][not(@DUPLICATE)]">
				<xsl:apply-templates select="$Customer[2]/ADDRESS[2][not(@DUPLICATE)]"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[position()>1][not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankUKAddress"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant,  2nd previous address -->
		<xsl:if test="$Customer[2]/ADDRESS[3][not(@DUPLICATE)]">
			<xsl:apply-templates select="$Customer[2]/ADDRESS[3][not(@DUPLICATE)]"/>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateResidency">
		<!-- 1st applicant, current address -->
		<xsl:apply-templates select="$Customer[1]/ADDRESS[1]" mode="RESY"/>
		<!-- 1st applicant, 1st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[1]/ADDRESS[2]">
				<xsl:apply-templates select="$Customer[1]/ADDRESS[2]" mode="RESY"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankResidency"/>
			</xsl:when>
		</xsl:choose>
		<!-- 1st applicant, 2st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[1]/ADDRESS[3]">
				<xsl:apply-templates select="$Customer[1]/ADDRESS[3]" mode="RESY"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankResidency"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant, current address -->
		<xsl:choose>
			<xsl:when test="$Customer[2]/ADDRESS[1][not(@DUPLICATE)]">
				<xsl:apply-templates select="$Customer[2]/ADDRESS[1][not(@DUPLICATE)]" mode="RESY"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankResidency"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant,  1st previous address -->
		<xsl:choose>
			<xsl:when test="$Customer[2]/ADDRESS[2][not(@DUPLICATE)]">
				<xsl:apply-templates select="$Customer[2]/ADDRESS[2][not(@DUPLICATE)]" mode="RESY"/>
			</xsl:when>
			<xsl:when test="$Customer[2]/ADDRESS[position()>1][not(@DUPLICATE)]">
				<xsl:call-template name="CreateBlankResidency"/>
			</xsl:when>
		</xsl:choose>
		<!-- 2nd applicant,  2nd previous address -->
		<xsl:if test="$Customer[2]/ADDRESS[3][not(@DUPLICATE)]">
			<xsl:apply-templates select="$Customer[2]/ADDRESS[3][not(@DUPLICATE)]" mode="RESY"/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="ADDRESS">
		<xsl:variable name="Country" select="@COUNTRY"/>
		<xsl:choose>
			<xsl:when test="@BFPO='1'">
				<xsl:call-template name="BritishForcesAddress">
					<xsl:with-param name="Address" select="."/>
				</xsl:call-template>
			</xsl:when>
			<xsl:when test="(not($Country)) or (not(string($Country))) or ($Combo[@GROUPNAME='Country' and @VALUEID=$Country]/COMBOVALIDATION[@VALIDATIONTYPE='UK'])">
				<xsl:call-template name="UKAddress">
					<xsl:with-param name="Address" select="."/>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="OverseasAddress">
					<xsl:with-param name="Address" select="."/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template match="ADDRESS" mode="RESY">
		<xsl:element name="RESY">
			<xsl:variable name="Id" select="@ID"/>
			<xsl:variable name="Position">
				<xsl:for-each select="../ADDRESS">
					<xsl:if test="@ID=$Id">
						<xsl:value-of select="position()-1"/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name="AddressInd">
				<xsl:choose>
					<xsl:when test="$Position=0">C</xsl:when>
					<xsl:when test="$Position=1">P</xsl:when>
					<xsl:when test="$Position=2">3</xsl:when>
				</xsl:choose>
			</xsl:variable>
			<xsl:variable name="Custno" select="number(../@CUSTOMERORDER)"/>
			<xsl:element name="ADDRESSNUMBER">
				<xsl:value-of select="concat(string($Position), ../@CUSTOMERORDER)"/>
			</xsl:element>
			<xsl:element name="ADDRESSINDICATOR">
					<xsl:value-of select="$AddressInd"/>
			</xsl:element>
			<xsl:variable name="NamesCountAlias" select="count(/RESPONSE/CUSTOMER/ALIAS[../ADDRESS/@ID=$Id or ../ADDRESS/@DUPLICATE=$Id])"/>
			<xsl:variable name="NamesCountMain" select="count(/RESPONSE/CUSTOMER[./ADDRESS/@ID=$Id or ./ADDRESS/@DUPLICATE=$Id])"/>
			<xsl:variable name="NamesCount" select="$NamesCountMain + $NamesCountAlias"/>
			<!-- names count is the number of valid names, not counting blank/padded name blocks -->
			<xsl:element name="NAMESCOUNT">
				<xsl:value-of select="$NamesCount"/>
			</xsl:element>
			<xsl:variable name="CurrentAddress" select="."/>
			<!-- SEQBLOCK for customers -->
			<xsl:call-template name="CreateMainJoint">
				<xsl:with-param name="CustomerOrder" select="1"/>
				<xsl:with-param name="CurrentAddress" select="."/>
				<xsl:with-param name="Id" select="$Id"/>
			</xsl:call-template>
			<!-- If the joint applicant has the same current address. we need to use the from and to dates from the duplicate address element -->			
			<xsl:choose>
				<xsl:when test="$AddressInd='C'">
					<xsl:call-template name="CreateMainJoint">
						<xsl:with-param name="CustomerOrder" select="2"/>
						<xsl:with-param name="CurrentAddress" select="$Customer[2]/ADDRESS[1]"/>
						<xsl:with-param name="Id" select="$Id"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:call-template name="CreateMainJoint">
						<xsl:with-param name="CustomerOrder" select="2"/>
						<xsl:with-param name="CurrentAddress" select="."/>
						<xsl:with-param name="Id" select="$Id"/>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
			<!-- SEQBLOCK for aliases -->
			<xsl:call-template name="CreateAliasSeq">
				<xsl:with-param name="Count" select="1"/>
				<xsl:with-param name="Id" select="$Id"/>
				<xsl:with-param name="MaxAliasAssoc" select="$MaxAliasAssoc"/>
			</xsl:call-template>
			<!-- Blank SEQBLOCKs for padding. 8 is the max to be sent. 2 for main and jont, the rest made up by alias's or blanks, all the blanks go at the end of this block-->
			<xsl:variable name="Customer1NameSeqs">
				<!-- the niumber of NAMESEQUENCE elements created for the main applicant -->
				<xsl:variable name="Customer1AssociatedWithAddress" select="$Customer[(ADDRESS/@ID=$Id or ADDRESS/@DUPLICATE=$Id) and @CUSTOMERORDER=1]"/>
				<xsl:choose>
					<xsl:when test="$Customer1AssociatedWithAddress">
						<xsl:value-of select="$TotAliasAssoc1 + 1"/>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:variable name="Customer2NameSeqs">
				<!-- the niumber of NAMESEQUENCE elements created for the joint applicant -->
				<xsl:variable name="Customer2AssociatedWithAddress" select="$Customer[(ADDRESS/@ID=$Id or ADDRESS/@DUPLICATE=$Id) and @CUSTOMERORDER=2]"/>
				<xsl:choose>
					<xsl:when test="$Customer2AssociatedWithAddress">
						<xsl:value-of select="$TotAliasAssoc2 + 1"/>
					</xsl:when>
					<xsl:otherwise>0</xsl:otherwise>
				</xsl:choose>
			</xsl:variable>
			<xsl:if test="(8 - ($Customer1NameSeqs + $Customer2NameSeqs)) &gt; 0">
				<xsl:call-template name="CreateBlankSeqBlocks">
					<xsl:with-param name="NumberRequired" select="8 - ($Customer1NameSeqs + $Customer2NameSeqs)"/>
				</xsl:call-template>
			</xsl:if>
			<xsl:element name="SEARCHFLAG">
				<xsl:choose>
					<xsl:when test="$IsOneShot">1</xsl:when>
				</xsl:choose>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="UKAddress">
		<xsl:param name="Address"/>
		<xsl:element name="ADDR">
			<xsl:element name="DETAILCODE"/>
			<xsl:element name="DATAITEM"/>
			<xsl:element name="DATAITEM"/>
			<xsl:element name="FLAT">
				<xsl:value-of select="substring($Address/@FLATNUMBER, 1, 16)"/>
			</xsl:element>
			<xsl:element name="HOUSENAME">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENAME, 1, 26)"/>
			</xsl:element>
			<xsl:element name="HOUSENUMBER">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENUMBER, 1, 10)"/>
			</xsl:element>
			<xsl:element name="STREET">
				<xsl:value-of select="substring($Address/@STREET, 1, 40)"/>
			</xsl:element>
			<xsl:element name="DISTRICT">
				<xsl:value-of select="substring($Address/@DISTRICT, 1, 30)"/>
			</xsl:element>
			<xsl:element name="TOWN">
				<xsl:value-of select="substring($Address/@TOWN, 1, 20)"/>
			</xsl:element>
			<xsl:element name="COUNTY">
				<xsl:value-of select="substring($Address/@COUNTY, 1, 20)"/>
			</xsl:element>
			<xsl:element name="POSTCODE">
				<xsl:value-of select="substring($Address/@POSTCODE, 1, 8)"/>
			</xsl:element>
			<!--<xsl:element name="DETAILCODE"/>-->
			<!--<xsl:element name="DATAITEM"/>-->
		</xsl:element>
	</xsl:template>
	<xsl:template name="BritishForcesAddress">
		<xsl:param name="Address"/>
		<xsl:element name="ADDR-BFPO">
			<xsl:element name="FlatNameOrNumber">
				<xsl:value-of select="substring($Address/@FLATNUMBER, 1, 16)"/>
			</xsl:element>
			<xsl:element name="HouseOrBuildingName">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENAME, 1, 26)"/>
			</xsl:element>
			<xsl:element name="HouseOrBuildingNumber">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENUMBER, 1, 10)"/>
			</xsl:element>
			<xsl:element name="Street">
				<xsl:value-of select="substring($Address/@STREET, 1, 40)"/>
			</xsl:element>
			<xsl:element name="District">
				<xsl:value-of select="substring($Address/@DISTRICT, 1, 30)"/>
			</xsl:element>
			<xsl:element name="TownOrCity">
				<xsl:value-of select="substring($Address/@TOWN, 1, 20)"/>
			</xsl:element>
			<xsl:element name="County">
				<xsl:value-of select="substring($Address/@COUNTY, 1, 20)"/>
			</xsl:element>
			<xsl:element name="PostCode"/>
			<xsl:element name="DETAILCODE"/>
			<xsl:element name="DATAITEM"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="OverseasAddress">
		<xsl:param name="Address"/>
		<xsl:element name="ADDR-OVERSEAS">
			<xsl:element name="FlatNameOrNumber">
				<xsl:value-of select="substring($Address/@FLATNUMBER, 1, 16)"/>
			</xsl:element>
			<xsl:element name="HouseOrBuildingName">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENAME, 1, 26)"/>
			</xsl:element>
			<xsl:element name="HouseOrBuildingNumber">
				<xsl:value-of select="substring($Address/@BUILDINGORHOUSENUMBER, 1, 10)"/>
			</xsl:element>
			<xsl:element name="Street">
				<xsl:value-of select="substring($Address/@STREET, 1, 40)"/>
			</xsl:element>
			<xsl:element name="District">
				<xsl:value-of select="substring($Address/@DISTRICT, 1, 30)"/>
			</xsl:element>
			<xsl:element name="TownOrCity">
				<xsl:value-of select="substring($Address/@TOWN, 1, 20)"/>
			</xsl:element>
			<xsl:element name="County">
				<xsl:value-of select="substring($Address/@COUNTY, 1, 20)"/>
			</xsl:element>
			<xsl:element name="DETAILCODE"/>
			<xsl:element name="DATAITEM"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="SumBasicAnnualIncome">
		<xsl:param name="Income"/>
		<xsl:param name="AllowableIncome"/>
		<xsl:param name="RunningTotal"/>
		<xsl:choose>
			<xsl:when test="not($Income)">
				<xsl:choose>
					<xsl:when test="number($RunningTotal) &gt; 0">
						<xsl:copy-of select="$RunningTotal"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="$AllowableIncome"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:variable name="CurrentTotal">
					<xsl:choose>
						<xsl:when test="not($Income[1]/@EARNEDINCOMEAMOUNT) or not($Income[1]/@PAYMENTFREQUENCYTYPE)">
							<xsl:value-of select="$RunningTotal"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$RunningTotal + ($Income[1]/@EARNEDINCOMEAMOUNT * $Income[1]/@PAYMENTFREQUENCYTYPE)"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<xsl:call-template name="SumBasicAnnualIncome">
					<xsl:with-param name="Income" select="$Income[position()>1]"/>
					<xsl:with-param name="AllowableIncome" select="$AllowableIncome"/>
					<xsl:with-param name="RunningTotal" select="$CurrentTotal"/>
				</xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="LTV_Calc">
		<xsl:param name="AmountReq"/>
		<xsl:param name="PurchPrice"/>
		<xsl:element name="LTTP">
			<xsl:choose>
				<xsl:when test="/RESPONSE/LOANCOMPONENT/@LTV">
					<xsl:variable name="LTVdb" select="floor(number(/RESPONSE/LOANCOMPONENT/@LTV)*100)"/>
					<xsl:choose>
						<xsl:when test="number($LTVdb) &gt; 99999">99999</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="$LTVdb"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
					<xsl:choose>
						<xsl:when test="(number($PurchPrice) &gt; 0) and (number($AmountReq) &gt; 0)">
							<xsl:variable name="LTV" select="floor(($AmountReq div $PurchPrice) * 10000)"/>
							<xsl:choose>
								<xsl:when test="number($LTV) &gt; 99999">99999</xsl:when>
								<xsl:otherwise>
									<xsl:value-of select="$LTV"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateBlankUKAddress">
		<xsl:element name="ADDR">
			<xsl:element name="DETAILCODE"/>
			<xsl:element name="DATAITEM"/>
			<xsl:element name="DATAITEM"/>
			<xsl:element name="FLAT"/>
			<xsl:element name="HOUSENAME"/>
			<xsl:element name="HOUSENUMBER"/>
			<xsl:element name="STREET"/>
			<xsl:element name="DISTRICT"/>
			<xsl:element name="TOWN"/>
			<xsl:element name="COUNTY"/>
			<xsl:element name="POSTCODE"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateBlankResidency">
		<xsl:element name="RESY">
			<xsl:element name="ADDRESSNUMBER"/>
			<xsl:element name="ADDRESSINDICATOR"/>
			<xsl:element name="NAMESCOUNT"/>
			<xsl:call-template name="CreateBlankSeqBlocks">
				<xsl:with-param name="NumberRequired" select="8"/>
			</xsl:call-template>
			<xsl:element name="SEARCHFLAG"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateMainJoint">
		<xsl:param name="CustomerOrder"/>
		<xsl:param name="CurrentAddress"/>
		<xsl:param name="Id"/>
		<xsl:variable name="Customer" select="$Customer[(ADDRESS/@ID=$Id or ADDRESS/@DUPLICATE=$Id) and @CUSTOMERORDER=$CustomerOrder]"/>
		<xsl:choose>
			<xsl:when test="$Customer and $CurrentAddress">
				<xsl:element name="NAMESEQUENCE">
					<xsl:value-of select="concat('0', $CustomerOrder)"/>
				</xsl:element>
				<xsl:element name="DATEFROM_DD">
					<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDIN, 9, 2)"/>
				</xsl:element>
				<xsl:element name="DATEFROM_MM">
					<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDIN, 6, 2)"/>
				</xsl:element>
				<xsl:element name="DATEFROM_CCYY">
					<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDIN, 1, 4)"/>
				</xsl:element>
				<xsl:choose>
					<xsl:when test="$CurrentAddress/@DATEMOVEDOUT">
						<xsl:element name="DATETO_DD">
							<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDOUT, 9, 2)"/>
						</xsl:element>
						<xsl:element name="DATETO_MM">
							<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDOUT, 6, 2)"/>
						</xsl:element>
						<xsl:element name="DATETO_CCYY">
							<xsl:value-of select="substring($CurrentAddress/@DATEMOVEDOUT, 1, 4)"/>
						</xsl:element>
					</xsl:when>
					<xsl:otherwise>
						<xsl:element name="DATETO_DD">
							<xsl:value-of select="substring(/RESPONSE/TODAY/@DATE, 9, 2)"/>
						</xsl:element>
						<xsl:element name="DATETO_MM">
							<xsl:value-of select="substring(/RESPONSE/TODAY/@DATE, 6, 2)"/>
						</xsl:element>
						<xsl:element name="DATETO_CCYY">
							<xsl:value-of select="substring(/RESPONSE/TODAY/@DATE, 1, 4)"/>
						</xsl:element>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
	<!--		<xsl:otherwise>
				<xsl:call-template name="CreateBlankSeqBlocks">
					<xsl:with-param name="NumberRequired" select="1"/>
				</xsl:call-template>
			</xsl:otherwise> -->
		</xsl:choose>
	</xsl:template>
	<xsl:template name="CreateAliasSeq">
		<xsl:param name="Count"/>
		<xsl:param name="Id"/>
		<xsl:param name="MaxAliasAssoc"/>
		<xsl:if test="$MaxAliasAssoc &gt; 0">	
			<xsl:call-template name="CreateAliasSeqInside">
				<xsl:with-param name="Count" select="$Count"/>
				<xsl:with-param name="CustomerOrder" select="1"/>
				<xsl:with-param name="Id" select="$Id"/>
			</xsl:call-template>
			<xsl:call-template name="CreateAliasSeqInside">
				<xsl:with-param name="Count" select="$Count"/>
				<xsl:with-param name="CustomerOrder" select="2"/>
				<xsl:with-param name="Id" select="$Id"/>
			</xsl:call-template>
			<xsl:if test="$Count &lt; $MaxAliasAssoc">
				<xsl:call-template name="CreateAliasSeq">
					<xsl:with-param name="Count" select="$Count+1"/>
					<xsl:with-param name="Id" select="$Id"/>
					<xsl:with-param name="MaxAliasAssoc" select="$MaxAliasAssoc"/>
				</xsl:call-template>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateAliasSeqInside">
		<xsl:param name="Count"/>
		<xsl:param name="CustomerOrder"/>
		<xsl:param name="Id"/>
		<xsl:variable name="Customer" select="$Customer[(ADDRESS/@ID=$Id or ADDRESS/@DUPLICATE=$Id) and @CUSTOMERORDER=$CustomerOrder]"/>
		<xsl:variable name="CustomerAliasCount" select="count(/RESPONSE/CUSTOMER[@CUSTOMERORDER=$CustomerOrder]/ALIAS[../ADDRESS/@ID=$Id or ../ADDRESS/@DUPLICATE=$Id])"/>
		<xsl:choose>
		<!-- no padding only create a alias entry if there is a genuine alias -->
			<xsl:when test="$Customer and (($CustomerAliasCount &gt; $Count) or ($CustomerAliasCount = $Count))">
				<xsl:element name="NAMESEQUENCE">
					<xsl:value-of select="concat($Count, $CustomerOrder)"/>
				</xsl:element>
				<xsl:element name="DATEFROM_DD"/>
				<xsl:element name="DATEFROM_MM"/>
				<xsl:element name="DATEFROM_CCYY"/>
				<xsl:element name="DATETO_DD"/>
				<xsl:element name="DATETO_MM"/>
				<xsl:element name="DATETO_CCYY"/>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="CreateBlankAliasAssocSeqBlocks">
		<xsl:param name="Count"/>
		<xsl:param name="CustomerOrder"/>
		<xsl:param name="MaxAliasAssoc"/>
		<!--<xsl:element name="SEQBLOCK">-->
		<xsl:element name="NAMESEQUENCE">
			<xsl:choose>
				<xsl:when test="$CustomerOrder and $Count">
					<xsl:value-of select="concat($Count, $CustomerOrder)"/>
				</xsl:when>
			</xsl:choose>
		</xsl:element>
		<xsl:element name="DATEFROM_DD"/>
		<xsl:element name="DATEFROM_MM"/>
		<xsl:element name="DATEFROM_CCYY"/>
		<xsl:element name="DATETO_DD"/>
		<xsl:element name="DATETO_MM"/>
		<xsl:element name="DATETO_CCYY"/>
		<!--</xsl:element>-->
		<xsl:if test="$Count &lt; $MaxAliasAssoc">
			<xsl:call-template name="CreateBlankAliasAssocSeqBlocks">
				<xsl:with-param name="Count" select="$Count+1"/>
				<xsl:with-param name="CustomerOrder" select="$CustomerOrder"/>
				<xsl:with-param name="MaxAliasAssoc" select="$MaxAliasAssoc"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateBlankSeqBlocks">
		<xsl:param name="NumberRequired"/>
		<!--<xsl:element name="SEQBLOCK">-->
		<xsl:element name="NAMESEQUENCE"/>
		<xsl:element name="DATEFROM_DD"/>
		<xsl:element name="DATEFROM_MM"/>
		<xsl:element name="DATEFROM_CCYY"/>
		<xsl:element name="DATETO_DD"/>
		<xsl:element name="DATETO_MM"/>
		<xsl:element name="DATETO_CCYY"/>
		<!--</xsl:element>-->
		<xsl:if test="$NumberRequired &gt; 1">
			<xsl:call-template name="CreateBlankSeqBlocks">
				<xsl:with-param name="NumberRequired" select="$NumberRequired -1"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="TDP1block">
		<xsl:element name="TPD1">
			<xsl:element name="SPAOPTOUT">
				<xsl:choose>
					<xsl:when test="APPLICATIONFACTFIND/@OPTOUTINDICATOR='1'">Y</xsl:when>
					<xsl:otherwise>N</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="TRANSASSOC">
				<xsl:choose>
					<xsl:when test="$GlobalParam[@NAME='ExperianCCTransAssoc']/@BOOLEAN='0'">N</xsl:when>
					<xsl:otherwise>Y</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="HHO">
				<xsl:choose>
					<xsl:when test="$GlobalParam[@NAME='ExperianCCHHO']/@BOOLEAN='0'">N</xsl:when>
					<xsl:otherwise>Y</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<xsl:element name="OPTOUTCUTOFF">
				<xsl:if test="$GlobalParam[@NAME='ExperianUseTPDAlertScore']/@BOOLEAN='1'">
					<xsl:value-of select="$GlobalParam[@NAME='ExperianTPDAlertScore']/@AMOUNT"/>
				</xsl:if>
			</xsl:element>
			<xsl:element name="DBTYPE"/>
			<xsl:element name="SEARCHTYPE">
				<xsl:variable name="thisStageSeq"><xsl:value-of select="number(/RESPONSE/APPLICATIONSTAGE/@STAGESEQUENCENO)"/></xsl:variable>
				<xsl:variable name="DIPStageSeq"><xsl:value-of select="number(/RESPONSE/GLOBALPARAMETER[@NAME='DIPStageSequenceNumber']/@AMOUNT)"/></xsl:variable>
				<xsl:choose>
					<xsl:when test="$thisStageSeq > $DIPStageSeq">C</xsl:when>
					<xsl:otherwise>Q</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="customerBlock">
		<xsl:for-each select="$Customer/CUSTOMERVERSION">
			<xsl:call-template name="CreateName">
				<xsl:with-param name="XML" select="."/>
				<xsl:with-param name="IsAliasAssoc"/>
				<xsl:with-param name="DateOfBirthXML" select="."/>
			</xsl:call-template>
		</xsl:for-each>
		<!-- if have one applicant with aliases need to create blank name block for joint applicant -->
		<xsl:if test="(number($MaxAliasAssoc) &gt; 0) and (count(/RESPONSE/CUSTOMER) = 1) ">
			<xsl:call-template name="CreateBlankName"/>
		</xsl:if>
		<xsl:if test="$Customer/ALIAS[string(@ALIASSEQUENCENUMBER)]">
			<xsl:variable name="MaxAlias">
				<xsl:for-each select="$Customer">
					<xsl:sort order="descending" data-type="number" select="count(ALIAS[string(@ALIASSEQUENCENUMBER)])"/>
					<xsl:if test="position()=1">
						<xsl:value-of select="count(ALIAS[string(@ALIASSEQUENCENUMBER)])"/>
					</xsl:if>
				</xsl:for-each>
			</xsl:variable>
			<xsl:variable name="MaxAliasApplicant">
			<!-- Which applicant has the most aliases(applicant 1 or 2). If they have the same num of aliases set to 2 -->
				<xsl:variable name="NoOfAliasForApp1" select="count($Customer[1]/ALIAS)"/>
				<xsl:variable name="NoOfAliasForApp2" select="count($Customer[2]/ALIAS)"/>
				<xsl:choose>
					<xsl:when test="$NoOfAliasForApp1 &gt; $NoOfAliasForApp2">
						<xsl:value-of select="1"/>
					</xsl:when>
					<xsl:when test="($NoOfAliasForApp1 &lt; $NoOfAliasForApp2) or ($NoOfAliasForApp1 = $NoOfAliasForApp2)">
						<xsl:value-of select="2"/>
					</xsl:when>
				</xsl:choose>
			</xsl:variable>
			<xsl:call-template name="CreateAliasName">
				<xsl:with-param name="MaxAlias" select="$MaxAlias"/>
				<xsl:with-param name="MaxAliasApplicant" select="$MaxAliasApplicant"/>
				<xsl:with-param name="CurrentApplicant" select="1"/>
				<xsl:with-param name="CurrentAlias" select="1"/>
			</xsl:call-template>
		</xsl:if>
		<xsl:call-template name="CreateAddresses"/>
		<xsl:call-template name="CreateResidency"/>
	</xsl:template>
	<xsl:template name="Validation">
		<xsl:element name="VALIDATION">
			<xsl:if test="/RESPONSE/CUSTOMER/CUSTOMERVERSION[not(@DATEOFBIRTH) or @DATEOFBIRTH='']">
				<xsl:element name="ERROR">Date of birth</xsl:element>
			</xsl:if>
			<xsl:if test="not(/RESPONSE/CUSTOMER/ADDRESS[@ID='1' and @DATEMOVEDIN])">
				<xsl:element name="ERROR">Current address - date moved in</xsl:element>
			</xsl:if>
			<xsl:if test="not(/RESPONSE/CUSTOMER/INCOMESUMMARY)">
				<xsl:element name="ERROR">Income</xsl:element>
			</xsl:if>
			<xsl:if test="not(/RESPONSE/LOANCOMPONENT[@AMOUNTREQUESTED])">
				<xsl:element name="ERROR">Amount requested</xsl:element>
			</xsl:if>
			<xsl:if test="not(/RESPONSE/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)">
				<xsl:element name="ERROR">Purchase Price</xsl:element>
			</xsl:if>
			<xsl:if test="/RESPONSE/CUSTOMER/CUSTOMERVERSION[not(@MARITALSTATUS) or @MARITALSTATUS='']">
				<xsl:element name="ERROR">Marital Status</xsl:element>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<!--EP2_750 -->
	<xsl:template name="GetDependants">
		<xsl:param name="CustPosition"/>
		<xsl:choose>
			<xsl:when test="$CustPosition = 1">
				<xsl:variable name="NumDep" select="/RESPONSE/APPLICATION/@NUMBEROFDEPENDANTS"/>
				<xsl:choose>
					<xsl:when test="$NumDep and number($NumDep) &lt;= 7">
						<xsl:value-of select="$NumDep"/>
					</xsl:when>
					<xsl:when test="$NumDep and number($NumDep) &gt; 7">7</xsl:when>
					<xsl:otherwise>8</xsl:otherwise>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>8</xsl:otherwise>
		</xsl:choose>
	</xsl:template>	
</xsl:stylesheet>
