<?xml version="1.0" encoding="UTF-8"?>
<!-- FullApplicationSummaryData.xsl
	 Purpose:	Translates XML data retrieved from a call to CRUD from AP040.asp
						to a more 'user friendly' data format as defined in ApplicationBO.doc (see Appendix C).  It also performs
						additional data processing such as calculating averages.  The CRUD call was originally going to be from ApplicationBO.GetFullApplicationSummaryData()
						so this is why the spec is contained within ApplicationBO.doc under GetFullApplicationSummaryData().
	Author:		Lee Hudson
	History:		LH 8/9/2006 	- Created
	LH 	27/9/2006 	- EP1168 - intermediary GUID now returned in XML for call to DC011.asp from AP040/AP040P.asp
	INR	11/03/2007	EP2_1885 Use TERMINMONTHS instead of TERMINYEARS
    PSC	16/03/2007	EP2_1726 Include new introducer structure
!-->
<!--<?altova_samplexml Test.xml?-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:my-scripts">
	<xsl:output method="xml" version="2.0" encoding="UTF-8" indent="no"/>
	<msxsl:script language="JavaScript" implements-prefix="user"><![CDATA[
		var vValuation;
						
		function SetValuation(pValuation)
		{
			vValuation = pValuation;
			return '';
		}
				
		function GetValuation()
		{
			return vValuation;
		}

		function CalculateAverage(pYr1Amount, pYr2Amount, pYr3Amount)
		{
			var vAvgAmount = '';
			var vDivideBy = 0;
			
			pYr1Amount = Number(pYr1Amount);
			pYr2Amount = Number(pYr2Amount);
			pYr3Amount = Number(pYr3Amount);
									
			if (pYr1Amount > 0)
			{
				vDivideBy = vDivideBy + 1;
			}

			if (pYr2Amount > 0)
			{
				vDivideBy = vDivideBy + 1;
			}
			
			if (pYr3Amount > 0)
			{
				vDivideBy = vDivideBy + 1;
			}
			
			if (vDivideBy > 0)
				vAvgAmount = Math.round((pYr1Amount + pYr2Amount + pYr3Amount) / vDivideBy);
				
			return vAvgAmount;
		}

		function CalculateAge(pDOB)
		{
		var vAge;
		
		//current date
		today = new Date();
		vCurrentYear = today.getYear() ;
		vCurrentMonth = today.getMonth() + 1;
		vCurrentDay = today.getDate();
		
		//20/01/1960
		DOB_year = pDOB.substring(6,10) - 0;
		DOB_month = pDOB.substring(3,5) - 0;
		DOB_day = pDOB.substring(0,2) - 0;
		
			if ( vCurrentMonth < DOB_month)
			{
				vAge = vCurrentYear - DOB_year - 1;
			}
			else if ( vCurrentMonth > DOB_month)
			{
				vAge = vCurrentYear - DOB_year;
			}
			else if ( vCurrentMonth == DOB_month)
			{
				if ( vCurrentDay < DOB_day)
				{
				vAge = vCurrentYear - DOB_year - 1;
				}
				else if ( vCurrentDay > DOB_day)
				{
				vAge = vCurrentYear - DOB_year;
				}
				else if ( vCurrentDay == DOB_day)
				{
				vAge = vCurrentYear - DOB_year;
				}
			}
			return vAge;
		}
				
		function CalculateElapsedTime(pDateMovedIn, pDateMovedOut) {
		/*
		//Purpose: returns elapsed time in months between the date moved in and the date moved out or 
		//between the date moved in and the system date (if the date moved out is null).
				
		   dateType is a numeric integer from 1 to 4, representing
		   the type of dateString passed, as defined above.
		
		   Returns string containing the age in years & months
		   in the format yyy years mm months
		*/
	

			var vDateMovedIn = new Date(pDateMovedIn.substring(6,10),
									pDateMovedIn.substring(3,5)-1,
									pDateMovedIn.substring(0,2));
			
			//if date moved in exists
			if (pDateMovedIn.substring(6,10) != '')
			{
				var vDateMovedIn_year = vDateMovedIn.getFullYear();
				var vDateMovedIn_month = vDateMovedIn.getMonth();
				var vDateMovedIn_date = vDateMovedIn.getDate();
		
		
				if (pDateMovedOut != '')
				{
					var vDateMovedOut = new Date(pDateMovedOut.substring(6,10),
										pDateMovedOut.substring(3,5)-1,
										pDateMovedOut.substring(0,2));
			
					var vDateMovedOut_year = vDateMovedOut.getFullYear();
					var vDateMovedOut_month = vDateMovedOut.getMonth();
					var vDateMovedOut_date = vDateMovedOut.getDate();
				}
				else  //no date moved out
				{		
					//get current date: current date replaces moved out date
					var now = new Date();
				
					var vDateMovedOut_year = now.getFullYear();
					var vDateMovedOut_month = now.getMonth();
					var vDateMovedOut_date = now.getDate();
				}
				
				vDateDiffInYears = vDateMovedOut_year - vDateMovedIn_year;
				
				
				if (vDateMovedOut_month >= vDateMovedIn_month)
					var vDateDiffInMonths = vDateMovedOut_month - vDateMovedIn_month;
				else {
					vDateDiffInYears--;
					var vDateDiffInMonths = 12 + vDateMovedOut_month - vDateMovedIn_month;
				}
			
				if (vDateMovedOut_date >= vDateMovedIn_date)
					var vDateDiffInDays = vDateMovedOut_date - vDateMovedIn_date;
				else {
					vDateDiffInMonths--;
					var vDateDiffInDays = 31 + vDateMovedOut_date - vDateMovedIn_date;
			
					if (vDateDiffInMonths < 0) {
						vDateDiffInMonths = 11;
						vDateDiffInYears--; 
					}
				}
				sReturn = vDateDiffInYears + " yrs " + vDateDiffInMonths + " mths"
			}
			else
				sReturn = '';
				
			return sReturn;
		}

	]]></msxsl:script>
	<!--  Get full application summary details !-->
	<xsl:template match="/">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:for-each select="RESPONSE/APPLICATION">
				<xsl:variable name="vAppNumber">
					<xsl:value-of select="@APPLICATIONNUMBER"/>
				</xsl:variable>
				
				<xsl:for-each select="APPLICATIONFACTFIND/CUSTOMERROLE">				
					<xsl:element name="CUSTOMER">
						<!--Write customer details-->
						<xsl:call-template name="WriteCustomerDetails">
							<xsl:with-param name="pAppNumber">
								<xsl:value-of select="$vAppNumber"/>
							</xsl:with-param>
						</xsl:call-template>
						<!--Write customer contact details-->
						<xsl:call-template name="WriteCustomerContactDetails">
							<xsl:with-param name="pAppNumber">
								<xsl:value-of select="$vAppNumber"/>
							</xsl:with-param>
						</xsl:call-template>
						<xsl:for-each select="CUSTOMERVERSION/EMPLOYMENT">
							<!--Write customer employment details-->
							<xsl:call-template name="WriteCustomerEmploymentDetails">
								<xsl:with-param name="pAppNumber">
									<xsl:value-of select="$vAppNumber"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:for-each>
						<xsl:for-each select="CUSTOMERADDRESS">
							<!--Write customer address details-->
							<xsl:call-template name="WriteCustomerAddressDetails">
								<xsl:with-param name="pAppNumber">
									<xsl:value-of select="$vAppNumber"/>
								</xsl:with-param>
							</xsl:call-template>
						</xsl:for-each>
					</xsl:element>
				</xsl:for-each>
				
				<!--Write mortgage property address details-->
				<xsl:call-template name="WriteMortgagePropertyAddressDetails">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>

				<!--Write introducer details-->
				<xsl:call-template name="WriteIntroducerDetails">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>

				<!--Write mortgage details-->
				<xsl:call-template name="WriteMortgageDetails">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>
				
				<!--Write application details-->
				<xsl:call-template name="WriteApplicationDetails">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>

				<!--Write credit details-->
				<xsl:call-template name="WriteCreditDetails">
					<xsl:with-param name="pAppNumber">
						<xsl:value-of select="$vAppNumber"/>
					</xsl:with-param>
				</xsl:call-template>

				<!--Write third party details-->
				<xsl:call-template name="WriteThirdPartyDetails">
				</xsl:call-template>
				
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--START: WriteApplicationDetails-->
	<xsl:template name="WriteApplicationDetails">
		<xsl:element name="APPLICATION">
			<xsl:attribute name="NATUREOFLOAN"><xsl:value-of select="APPLICATIONFACTFIND/@NATUREOFLOAN"/></xsl:attribute>
			<xsl:attribute name="NATUREOFLOANTEXT"><xsl:value-of select="APPLICATIONFACTFIND/@NATUREOFLOAN_TEXT"/></xsl:attribute>
			<xsl:attribute name="TYPEOFBUYER"><xsl:value-of select="@TYPEOFBUYER"/></xsl:attribute>
			<xsl:attribute name="TYPEOFBUYERTEXT"><xsl:value-of select="@TYPEOFBUYER_TEXT"/></xsl:attribute>
			<xsl:attribute name="DECLAREDINCOMEMULTIPLE"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@INCOMEMULTIPLE"/></xsl:attribute>
			<xsl:attribute name="CONFIRMEDINCOMEMULTIPLE"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@CONFIRMEDCALCULATEDINCMULTIPLE"/></xsl:attribute>
			<xsl:attribute name="DIPDATE"><xsl:value-of select="@APPLICATIONDATE"/></xsl:attribute>
			<xsl:attribute name="APPLICATIONRECEIVEDDATE"><xsl:value-of select="APPLICATIONFACTFIND/@APPLICATIONRECEIVEDDATE"/></xsl:attribute>
			<xsl:attribute name="APPLICATIONINCOMESTATUS"><xsl:value-of select="APPLICATIONFACTFIND/@APPLICATIONINCOMESTATUS_TEXT"/></xsl:attribute>
			<xsl:attribute name="LEVELOFADVICE"><xsl:value-of select="NULL"/></xsl:attribute>
			<xsl:variable name="vInstructionSequenceNumber">
				<xsl:value-of select="APPLICATIONFACTFIND/VALUERINSTRUCTION[@VALUATIONSTATUS_TYPE_C='true']/@INSTRUCTIONSEQUENCENO"/>
			</xsl:variable>
			<xsl:variable name="vPresentValuation">
				<xsl:value-of select="APPLICATIONFACTFIND/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO=$vInstructionSequenceNumber]/@PRESENTVALUATION"/>
			</xsl:variable>
			<xsl:attribute name="PRESENTVALUATION"><xsl:value-of select="$vPresentValuation"/></xsl:attribute>
			<xsl:attribute name="POSTWORKSVALUATION"><xsl:value-of select="APPLICATIONFACTFIND/VALNREPVALUATION/@POSTWORKSVALUATION"/></xsl:attribute>
			<xsl:variable name="vPostWorksValuation">
				<xsl:value-of select="APPLICATIONFACTFIND/VALNREPVALUATION/@POSTWORKSVALUATION"/>
			</xsl:variable>
			<!-- Calculate retention amount -->
			<xsl:variable name="vRetentionRoads">
				<xsl:value-of select="APPLICATIONFACTFIND/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO=$vInstructionSequenceNumber]/@RETENTIONSROADS"/>
			</xsl:variable>
			<xsl:variable name="vRetentionWorks">
				<xsl:value-of select="APPLICATIONFACTFIND/VALNREPVALUATION[@INSTRUCTIONSEQUENCENO=$vInstructionSequenceNumber]/@RETENTIONWORKS"/>
			</xsl:variable>
			<xsl:if test="$vRetentionRoads != '' and $vRetentionWorks != '' ">
				<xsl:attribute name="RETENTIONAMOUNT"><xsl:value-of select="number($vRetentionRoads+$vRetentionWorks)"/></xsl:attribute>
			</xsl:if>

			<!-- Please NOTE: Var vValuation is only used as a mechanism for calling SetValuation(), one of the drawbacks of XSL... This variables values are redundant-->
			<xsl:choose>
				<xsl:when test="APPLICATIONFACTFIND/NEWPROPERTY/@CHANGEOFPROPERTY = '1' ">
					<xsl:variable name="vValuation">
						<xsl:value-of select="user:SetValuation(APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"/>
					</xsl:variable>
				</xsl:when>
				<xsl:otherwise>
					<!-- not change in property -->
					<xsl:choose>
						<xsl:when test="($vRetentionRoads > 0 or $vRetentionWorks > 0) and $vPostWorksValuation > 0">
							<xsl:variable name="vValuation">
								<xsl:value-of select="user:SetValuation($vPostWorksValuation)"/>
							</xsl:variable>
						</xsl:when>
						<xsl:otherwise>
							<!-- no retention or postworks valuation -->
							<xsl:choose>
								<xsl:when test="APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPE_N='true'">
									<!-- new loan? -->
									<xsl:choose>
										<!-- &lt; = less than -->
										<xsl:when test="$vPresentValuation > 0 and $vPresentValuation &lt; APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE">
											<xsl:variable name="vValuation">
												<xsl:value-of select="user:SetValuation($vPresentValuation)"/>
											</xsl:variable>
										</xsl:when>
										<xsl:otherwise>
											<xsl:variable name="vValuation">
												<xsl:value-of select="user:SetValuation(APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"/>
											</xsl:variable>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:when>
								<xsl:otherwise>
									<!-- not new loan, but re-mortgage, further advance, SPL or POE or TOE -->
									<xsl:choose>
										<xsl:when test="APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPE_R='true' and APPLICATIONFACTFIND/NEWPROPERTY/@VALUATIONTYPE_TYPE_AU='true'">
											<!-- remortgage (R) and hometrack (AU) -->
											<xsl:choose>
												<xsl:when test="$vPresentValuation &lt; APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE">
													<xsl:variable name="vValuation">
														<xsl:value-of select="user:SetValuation($vPresentValuation)"/>
													</xsl:variable>
												</xsl:when>
												<xsl:otherwise>
													<!-- purchase price is lower -->
													<xsl:variable name="vValuation">
														<xsl:value-of select="user:SetValuation(APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"/>
													</xsl:variable>
												</xsl:otherwise>
											</xsl:choose>
										</xsl:when>
										<xsl:otherwise>
											<!-- not re-mortgage and hometrack -->
											<xsl:choose>
												<xsl:when test="$vPresentValuation &gt; 0">
													<!-- > 0 -->
													<xsl:variable name="vValuation">
														<xsl:value-of select="user:SetValuation($vPresentValuation)"/>
													</xsl:variable>
												</xsl:when>
												<xsl:otherwise>
													<!-- present valuation is 0 or not entered -->
													<xsl:variable name="vValuation">
														<xsl:value-of select="user:SetValuation(APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE)"/>
													</xsl:variable>
												</xsl:otherwise>
											</xsl:choose>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
			<xsl:attribute name="VALUATION"><xsl:value-of select="user:GetValuation()"/></xsl:attribute>
			<xsl:attribute name="APPLICATIONPRIORITY"><xsl:value-of select="NULL"/></xsl:attribute>
			<xsl:attribute name="REGULATIONINDICATOR"><xsl:value-of select="NULL"/></xsl:attribute>
			<xsl:attribute name="OPTEDOUT"><xsl:value-of select="NULL"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--END: WriteApplicationDetails-->
	<!--START: WriteCustomerDetails-->
	<xsl:template name="WriteCustomerDetails">
		<xsl:element name="CUSTOMERDETAILS">
			<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERROLETYPE"><xsl:value-of select="@CUSTOMERROLETYPE"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERROLETEXT"><xsl:value-of select="@CUSTOMERROLETYPE_TEXT"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERORDER"><xsl:value-of select="@CUSTOMERORDER"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERSURNAME"><xsl:value-of select="CUSTOMERVERSION/@SURNAME"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERFIRSTFORENAME"><xsl:value-of select="CUSTOMERVERSION/@FIRSTFORENAME"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERTITLE"><xsl:value-of select="CUSTOMERVERSION/@TITLE"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERTITLETEXT"><xsl:value-of select="CUSTOMERVERSION/@TITLE_TEXT"/></xsl:attribute>
			<xsl:attribute name="MARITALSTATUS"><xsl:value-of select="CUSTOMERVERSION/@MARITALSTATUS"/></xsl:attribute>
			<xsl:attribute name="MARITALSTATUSTEXT"><xsl:value-of select="CUSTOMERVERSION/@MARITALSTATUS_TEXT"/></xsl:attribute>
			<xsl:variable name="vDateOfBirth">
				<xsl:value-of select="CUSTOMERVERSION/@DATEOFBIRTH"/>
			</xsl:variable>
			<xsl:attribute name="DATEOFBIRTH"><xsl:value-of select="$vDateOfBirth"/></xsl:attribute>
			<xsl:attribute name="AGE"><xsl:value-of select="user:CalculateAge(string($vDateOfBirth))"/></xsl:attribute>
			<xsl:attribute name="GENDER"><xsl:value-of select="CUSTOMERVERSION/@GENDER"/></xsl:attribute>
			<xsl:attribute name="GENDERTEXT"><xsl:value-of select="CUSTOMERVERSION/@GENDER_TEXT"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--END: WriteCustomerDetails-->
	<!--START: WriteMortgagePropertyAddressDetails-->
	<xsl:template name="WriteMortgagePropertyAddressDetails">
		<xsl:param name="pAppNumber"/>
		<xsl:element name="MORTGAGEPROPERTYADDRESS">
			<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@BUILDINGORHOUSENAME"/></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:attribute>
			<xsl:attribute name="FLATNUMBER"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@FLATNUMBER"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@STREET"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@DISTRICT"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@TOWN"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@COUNTY"/></xsl:attribute>
			<xsl:attribute name="COUNTRY"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@COUNTRY"/></xsl:attribute>
			<xsl:attribute name="COUNTRY_TEXT"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@COUNTRY_TEXT"/></xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="APPLICATIONFACTFIND/NEWPROPERTY/NEWPROPERTYADDRESS/MORTGAGEPROPERTYADDRESS/@POSTCODE"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--END: WriteMortgagePropertyAddressDetails-->
	<!--START: WriteCustomerContactDetails-->
	<xsl:template name="WriteCustomerContactDetails">
		<xsl:element name="CUSTOMERCONTACTDETAILS">
			<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="CUSTOMERVERSION/@CONTACTEMAILADDRESS"/></xsl:attribute>
			<xsl:attribute name="CORRESPONDENCESALUTATION"><xsl:value-of select="CUSTOMERVERSION/@CORRESPONDENCESALUTATION"/></xsl:attribute>
			<xsl:for-each select="CUSTOMERTELEPHONENUMBER">
				<xsl:element name="CUSTOMERTELEPHONE">
					<xsl:attribute name="TELEPHONESEQUENCENUMBER"><xsl:value-of select="@TELEPHONESEQUENCENUMBER"/></xsl:attribute>
					<xsl:attribute name="TELEPHONENUMBER"><xsl:value-of select="@TELEPHONENUMBER"/></xsl:attribute>
					<xsl:attribute name="USAGE"><xsl:value-of select="@USAGE"/></xsl:attribute>
					<xsl:attribute name="CONTACTTIME"><xsl:value-of select="@CONTACTTIME"/></xsl:attribute>
					<xsl:attribute name="EXTENTIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:attribute>
					<xsl:attribute name="PREFERREDMETHODOFCONTACT"><xsl:value-of select="@PREFERREDMETHODOFCONTACT"/></xsl:attribute>
					<xsl:attribute name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:attribute>
					<xsl:attribute name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--END: WriteCustomerContactDetails-->
	<!--START: WriteCustomerEmploymentDetails-->
	<xsl:template name="WriteCustomerEmploymentDetails">
		<xsl:element name="CUSTOMEREMPLOYMENT">
			<xsl:variable name="vSequenceNumber">
				<xsl:value-of select="@EMPLOYMENTSEQUENCENUMBER"/>
			</xsl:variable>
			<xsl:attribute name="EMPLOYMENTSEQUENCENUMBER"><xsl:value-of select="$vSequenceNumber"/></xsl:attribute>
			<xsl:attribute name="EMPLOYMENTSTATUS"><xsl:value-of select="@EMPLOYMENTSTATUS"/></xsl:attribute>
			<xsl:attribute name="EMPLOYMENTSTATUSTEXT"><xsl:value-of select="@EMPLOYMENTSTATUS_TEXT"/></xsl:attribute>
			<xsl:attribute name="OCCUPATION"><xsl:value-of select="@JOBTITLE"/></xsl:attribute>
			<xsl:attribute name="OCCUPATIONTYPE"><xsl:value-of select="@OCCUPATIONTYPE"/></xsl:attribute>
			<xsl:attribute name="OCCUPATIONTEXT"><xsl:value-of select="@OCCUPATIONTYPE_TEXT"/></xsl:attribute>
			<xsl:variable name="vDateStarted">
				<xsl:value-of select="@DATESTARTEDORESTABLISHED"/>
			</xsl:variable>
			<xsl:variable name="vDateEnded">
				<xsl:value-of select="@DATELEFTORCEASEDTRADING"/>
			</xsl:variable>
			<xsl:attribute name="DATESTARTED"><xsl:value-of select="$vDateStarted"/></xsl:attribute>
			<xsl:attribute name="DATEENDED"><xsl:value-of select="$vDateEnded"/></xsl:attribute>
			<xsl:attribute name="TIMEEMPLOYED"><xsl:value-of select="user:CalculateElapsedTime(string($vDateStarted), string($vDateEnded))"/></xsl:attribute>
			<xsl:variable name="vMainStatus">
				<xsl:value-of select="@MAINSTATUS"/>
			</xsl:variable>
			<xsl:attribute name="MAINSTATUS"><xsl:value-of select="$vMainStatus"/></xsl:attribute>
			<xsl:choose>
				<xsl:when test="$vMainStatus = 1">
					<xsl:attribute name="MAINSTATUSTEXT"><xsl:value-of select="'Yes'"/></xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="MAINSTATUSTEXT"><xsl:value-of select="NULL"/></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<xsl:choose>
				<xsl:when test="@EMPLOYMENTSTATUS_TYPE_E='true' or @EMPLOYMENTSTATUS_TYPE_R='true'">
					<!-- employed or retired? -->
					<xsl:attribute name="GROSSINCOME"><xsl:value-of select="EARNEDINCOME[@EMPLOYMENTSEQUENCENUMBER=$vSequenceNumber]/@EARNEDINCOMEAMOUNT"/></xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:choose>
						<xsl:when test="@EMPLOYMENTSTATUS_TYPE_S='true' or @EMPLOYMENTSTATUS_TYPE_C='true'">
							<!--self-employed or contract? -->
							<xsl:variable name="vYr1Amount">
								<xsl:value-of select="NETPROFIT[@EMPLOYMENTSEQUENCENUMBER=$vSequenceNumber]/@YEAR1AMOUNT"/>
							</xsl:variable>
							<xsl:variable name="vYr2Amount">
								<xsl:value-of select="NETPROFIT[@EMPLOYMENTSEQUENCENUMBER=$vSequenceNumber]/@YEAR2AMOUNT"/>
							</xsl:variable>
							<xsl:variable name="vYr3Amount">
								<xsl:value-of select="NETPROFIT[@EMPLOYMENTSEQUENCENUMBER=$vSequenceNumber]/@YEAR3AMOUNT"/>
							</xsl:variable>
							<xsl:attribute name="GROSSINCOME"><xsl:value-of select="user:CalculateAverage(string($vYr1Amount), string($vYr2Amount), string($vYr3Amount))"/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="GROSSINCOME"><xsl:value-of select="0"/></xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<!--END: WriteCustomerEmploymentDetails-->

	<!--START: WriteCustomerAddressDetails-->
	<xsl:template name="WriteCustomerAddressDetails">
		<xsl:element name="CUSTOMERADDRESS">
			<xsl:variable name="vDateMovedIn"><xsl:value-of select="@DATEMOVEDIN"/></xsl:variable>
			<xsl:variable name="vDateMovedOut"><xsl:value-of select="@DATEMOVEDOUT"/></xsl:variable>
			<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="$vDateMovedIn"/></xsl:attribute>
			<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="$vDateMovedOut"/></xsl:attribute>
			<xsl:attribute name="TIMEATADDRESS"><xsl:value-of select="user:CalculateElapsedTime(string($vDateMovedIn), string($vDateMovedOut))"/></xsl:attribute>
			<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="@NATUREOFOCCUPANCY"/></xsl:attribute>
			<xsl:attribute name="NATUREOFOCCUPANCYTEXT"><xsl:value-of select="@NATUREOFOCCUPANCY_TEXT"/></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENAME"><xsl:value-of select="ADDRESS/@BUILDINGORHOUSENAME"/></xsl:attribute>
			<xsl:attribute name="BUILDINGORHOUSENUMBER"><xsl:value-of select="ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:attribute>
			<xsl:attribute name="FLATNUMBER"><xsl:value-of select="ADDRESS/@FLATNUMBER"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="ADDRESS/@STREET"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="ADDRESS/@DISTRICT"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="ADDRESS/@TOWN"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="ADDRESS/@COUNTY"/></xsl:attribute>
			<xsl:attribute name="COUNTRY"><xsl:value-of select="ADDRESS/@COUNTRY"/></xsl:attribute>
			<xsl:attribute name="COUNTRY_TEXT"><xsl:value-of select="ADDRESS/@COUNTRY_TEXT"/></xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="ADDRESS/@POSTCODE"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<!--END: WriteCustomerAddressDetails-->

	<!--START: WriteIntroducerDetails-->
	<xsl:template name="WriteIntroducerDetails">
		<xsl:for-each select="APPLICATIONFACTFIND/APPLICATIONINTRODUCER">
			<xsl:variable name="APPLICATIONINTRODUCERSEQNO" select="@APPLICATIONINTRODUCERSEQNO"/>
			<xsl:for-each select=".//MORTGAGECLUBNETWORKASSOCIATION">
				<xsl:element name="INTERMEDIARY">
					<xsl:attribute name="APPLICATIONINTRODUCERSEQNO"><xsl:value-of select="$APPLICATIONINTRODUCERSEQNO"/></xsl:attribute>
					<xsl:call-template name="mortgageClub"/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select=".//PRINCIPALFIRM">
				<xsl:element name="INTERMEDIARY">
					<xsl:attribute name="APPLICATIONINTRODUCERSEQNO"><xsl:value-of select="$APPLICATIONINTRODUCERSEQNO"/></xsl:attribute>
					<xsl:call-template name="principalFirm"/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select=".//ARFIRM">
				<xsl:element name="INTERMEDIARY">
					<xsl:attribute name="APPLICATIONINTRODUCERSEQNO"><xsl:value-of select="$APPLICATIONINTRODUCERSEQNO"/></xsl:attribute>
					<xsl:call-template name="arFirm"/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select=".//INTRODUCER">
				<xsl:element name="INTERMEDIARY">
					<xsl:attribute name="APPLICATIONINTRODUCERSEQNO"><xsl:value-of select="$APPLICATIONINTRODUCERSEQNO"/></xsl:attribute>
					<xsl:call-template name="introducer"/>
				</xsl:element>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>
	<!--END: WriteIntroducerDetails-->
	
<!--START: WriteMortgageDetails-->
	<xsl:template name="WriteMortgageDetails">
		<xsl:element name="MORTGAGEDETAILS">
			<xsl:attribute name="TOTALLOANAMOUNT"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@TOTALLOANAMOUNT"/></xsl:attribute>
			<xsl:attribute name="AMOUNTREQUESTED"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@AMOUNTREQUESTED"/></xsl:attribute>
			<xsl:attribute name="PURCHASEPRICE"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@PURCHASEPRICEORESTIMATEDVALUE"/></xsl:attribute>
			<xsl:attribute name="LTV"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@LTV"/></xsl:attribute>
			<xsl:attribute name="TOTALNETMONTHLYCOST"><xsl:value-of select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@TOTALNETMONTHLYCOST"/></xsl:attribute>
			<xsl:for-each select="APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/LOANCOMPONENT">
				<xsl:element name="LOANCOMPONENT">
					<xsl:attribute name="SEQUENCE"><xsl:value-of select="@LOANCOMPONENTSEQUENCENUMBER"/></xsl:attribute>
					<xsl:attribute name="MORTGAGEPRODUCTCODE"><xsl:value-of select="@MORTGAGEPRODUCTCODE"/></xsl:attribute>
					<xsl:attribute name="PRODUCTTYPE"><xsl:value-of select="MORTGAGEPRODUCT/INTERESTRATETYPE/@RATETYPE"/></xsl:attribute>
					<xsl:attribute name="PRODUCTDESCRIPTION"><xsl:value-of select="MORTGAGEPRODUCT/MORTGAGEPRODUCTLANGUAGE/@PRODUCTNAME"/></xsl:attribute>
					<xsl:attribute name="RESOLVEDRATE"><xsl:value-of select="@RESOLVEDRATE"/></xsl:attribute>
					<xsl:attribute name="ADJUSTEDRATE"><xsl:value-of select="@MANUALADJUSTMENTPERCENT"/></xsl:attribute>
					<xsl:attribute name="LOANAMOUNT"><xsl:value-of select="@TOTALLOANCOMPONENTAMOUNT"/></xsl:attribute>
					<xsl:attribute name="TERMINYEARS"><xsl:value-of select="@TERMINYEARS"/></xsl:attribute>
					<xsl:attribute name="TERMINMONTHS"><xsl:value-of select="@TERMINMONTHS"/></xsl:attribute>
					<xsl:attribute name="REPAYMENTMETHOD"><xsl:value-of select="@REPAYMENTMETHOD"/></xsl:attribute>
					<xsl:attribute name="REPAYMENTMETHODTEXT"><xsl:value-of select="@REPAYMENTMETHOD_TEXT"/></xsl:attribute>
					<xsl:attribute name="APR"><xsl:value-of select="@APR"/></xsl:attribute>
					<xsl:attribute name="NETMONTHLYCOST"><xsl:value-of select="@NETMONTHLYCOST"/></xsl:attribute>
					<xsl:attribute name="FINALRATEMONTHLYCOST"><xsl:value-of select="@FINALRATEMONTHLYCOST"/></xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--END: WriteMortgageDetails-->
	
	<!--START: WriteCreditDetails-->
	<xsl:template name="WriteCreditDetails">
		<!-- if an application credit check record exists -->
		<xsl:if test="APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/@APPLICATIONNUMBER != '' ">
			<xsl:element name="CREDITDETAILS">
				<xsl:attribute name="EXPERIANREFERENCE"><xsl:value-of select="APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/@CREDITCHECKREFERENCENUMBER"/></xsl:attribute>
				<xsl:variable name="vEligibilityInd">
					<xsl:value-of select="APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/BESPOKEDECISION/@ELIGIBILITYINDICATOR"/>
				</xsl:variable>
				<xsl:attribute name="ELIGIBILITYINDICATOR"><xsl:value-of select="$vEligibilityInd"/></xsl:attribute>
				
				<!-- Get eligibility indicator text from combo table-->
				<xsl:for-each select="COMBOGROUP[@GROUPNAME='ProductScheme']/COMBOVALUE/COMBOVALIDATION[@VALIDATIONTYPE=$vEligibilityInd]">
					<xsl:attribute name="ELIGIBILITYINDICATOR_TEXT"><xsl:value-of select="../@VALUENAME"/></xsl:attribute>
				</xsl:for-each>
				
				<xsl:attribute name="DELPHITOTAL"><xsl:value-of select="APPLICATIONFACTFIND/APPLICATIONCREDITCHECK/CREDITCHECKSCORE/@SCORE"/></xsl:attribute>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<!--END: WriteCreditDetails-->
		
	
	<!--START: WriteThirdPartyDetails-->
	<xsl:template name="WriteThirdPartyDetails">
		<xsl:element name="APPLICATIONTHIRDPARTYLIST">	
			<xsl:for-each select="APPLICATIONTHIRDPARTY">
				<xsl:element name="APPLICATIONTHIRDPARTY">
					<xsl:element name="APPLICATIONNUMBER"><xsl:value-of select="@APPLICATIONNUMBER"/></xsl:element>
					<xsl:element name="APPLICATIONFACTFINDNUMBER"><xsl:value-of select="../APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER"/></xsl:element>
					<xsl:choose>
						<xsl:when test="THIRDPARTY/@THIRDPARTYGUID != '' ">
								<xsl:attribute name="TYPE"><xsl:value-of select="'THIRDPARTY'"></xsl:value-of></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="NAMEANDADDRESSDIRECTORY/@DIRECTORYGUID != '' ">
									<xsl:attribute name="TYPE"><xsl:value-of select="'NAMEANDADDRESSDIRECTORY'"></xsl:value-of></xsl:attribute>
										<xsl:element name="CONTACTDETAILSGUID"><xsl:value-of select="@CONTACTDETAILSGUID"/></xsl:element>
								</xsl:when>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
																			
					<xsl:element name="DIRECTORYGUID"><xsl:value-of select="@DIRECTORYGUID"/></xsl:element>
					<xsl:element name="THIRDPARTYGUID"><xsl:value-of select="@THIRDPARTYGUID"/></xsl:element>	
					
					<xsl:if test="THIRDPARTY/@THIRDPARTYGUID != '' ">
						<xsl:element name="THIRDPARTY">				
							<xsl:element name="THIRDPARTYGUID"><xsl:value-of select="THIRDPARTY/@THIRDPARTYGUID"/></xsl:element>
							<xsl:element name="THIRDPARTYTYPE"><xsl:value-of select="THIRDPARTY/@THIRDPARTYTYPE"/></xsl:element>
							<xsl:element name="THIRDPARTYTYPE_TEXT"><xsl:value-of select="THIRDPARTY/@THIRDPARTYTYPE_TEXT"/></xsl:element>
							<xsl:element name="COMPANYNAME"><xsl:value-of select="THIRDPARTY/@COMPANYNAME"/></xsl:element>
							<xsl:element name="ADDRESSGUID"><xsl:value-of select="THIRDPARTY/@ADDRESSGUID"/></xsl:element>
							
							<xsl:if test="THIRDPARTY/@THIRDPARTYTYPE = 3 "> <!-- if bank/building society -->
								<xsl:element name="ACCOUNTNUMBER"><xsl:value-of select="@ACCOUNTNUMBER"/></xsl:element>
								<xsl:element name="REPAYMENTBANKACCOUNTINDICATOR"><xsl:value-of select="@REPAYMENTBANKACCOUNTINDICATOR"/></xsl:element>	
							</xsl:if>
							
							<xsl:element name="ADDRESS">
								<xsl:element name="ADDRESSGUID"><xsl:value-of select="THIRDPARTY/ADDRESS/@ADDRESSGUID"/></xsl:element>
								<xsl:element name="BUILDINGORHOUSENAME"><xsl:value-of select="THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME"/></xsl:element>
								<xsl:element name="BUILDINGORHOUSENUMBER"><xsl:value-of select="THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:element>
								<xsl:element name="STREET"><xsl:value-of select="THIRDPARTY/ADDRESS/@STREET"/></xsl:element>
								<xsl:element name="DISTRICT"><xsl:value-of select="THIRDPARTY/ADDRESS/@DISTRICT"/></xsl:element>
								<xsl:element name="TOWN"><xsl:value-of select="THIRDPARTY/ADDRESS/@TOWN"/></xsl:element>
								<xsl:element name="COUNTY"><xsl:value-of select="THIRDPARTY/ADDRESS/@COUNTY"/></xsl:element>
								<xsl:element name="COUNTRY"><xsl:value-of select="THIRDPARTY/ADDRESS/@COUNTRY"/></xsl:element>
								<xsl:element name="COUNTRY_TEXT"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTRY_TEXT"/></xsl:element>
								<xsl:element name="POSTCODE"><xsl:value-of select="THIRDPARTY/ADDRESS/@POSTCODE"/></xsl:element>
							</xsl:element>
							<xsl:element name="CONTACTDETAILS">
								<xsl:element name="CONTACTDETAILSGUID"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@CONTACTDETAILSGUID"/></xsl:element>
								<xsl:element name="CONTACTFORENAME"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@CONTACTFORENAME"/></xsl:element>
								<xsl:element name="CONTACTSURNAME"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@CONTACTSURNAME"/></xsl:element>
								<xsl:element name="CONTACTTITLE"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@CONTACTTITLE"/></xsl:element>
								<xsl:element name="CONTACTTYPE"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@CONTACTTYPE"/></xsl:element>
								<xsl:element name="EMAILADDRESS"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@EMAILADDRESS"/></xsl:element>
								<xsl:element name="FAXNUMBER"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@FAXNUMBER"/></xsl:element>
								<xsl:element name="TELEPHONEEXTENSIONNUMBER"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@TELEPHONEEXTENSIONNUMBER"/></xsl:element>
								<xsl:element name="TELEPHONENUMBER"><xsl:value-of select="THIRDPARTY/CONTACTDETAILS/@TELEPHONENUMBER"/></xsl:element>
							
								<xsl:for-each select="THIRDPARTY/CONTACTTELEPHONEDETAILS">
									<xsl:if test="@USAGE != '' ">
										<xsl:element name="CONTACTTELEPHONEDETAILS">
											<xsl:element name="CONTACTDETAILSGUID"><xsl:value-of select="@CONTACTDETAILSGUID"/></xsl:element>
											<xsl:element name="TELEPHONESEQNUM"><xsl:value-of select="@TELEPHONESEQNUM"/></xsl:element>
											<xsl:element name="USAGE"><xsl:value-of select="@USAGE"/></xsl:element>
											<xsl:element name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:element>
											<xsl:element name="TELENUMBER"><xsl:value-of select="@TELENUMBER"/></xsl:element>
											<xsl:element name="EXTENSIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:element>
											<xsl:element name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:element>
										</xsl:element>
									</xsl:if>
								</xsl:for-each>

							</xsl:element>
						</xsl:element>	
					</xsl:if>
					<xsl:if test="NAMEANDADDRESSDIRECTORY/@DIRECTORYGUID != '' ">
						<xsl:element name="THIRDPARTY">
							<xsl:element name="THIRDPARTYTYPE"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/@NAMEANDADDRESSTYPE"/></xsl:element>
							<xsl:element name="THIRDPARTYTYPE_TEXT"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/@NAMEANDADDRESSTYPE_TEXT"/></xsl:element>
							<xsl:element name="COMPANYNAME"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/@COMPANYNAME"/></xsl:element>
							<xsl:element name="ADDRESSGUID"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/@ADDRESSGUID"/></xsl:element>
							<xsl:element name="ADDRESS">
								<xsl:element name="ADDRESSGUID"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@ADDRESSGUID"/></xsl:element>
								<xsl:element name="BUILDINGORHOUSENAME"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENAME"/></xsl:element>
								<xsl:element name="BUILDINGORHOUSENUMBER"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@BUILDINGORHOUSENUMBER"/></xsl:element>
								<xsl:element name="STREET"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@STREET"/></xsl:element>
								<xsl:element name="DISTRICT"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@DISTRICT"/></xsl:element>
								<xsl:element name="TOWN"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@TOWN"/></xsl:element>
								<xsl:element name="COUNTY"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTY"/></xsl:element>
								<xsl:element name="COUNTRY"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTRY"/></xsl:element>
								<xsl:element name="COUNTRY_TEXT"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@COUNTRY_TEXT"/></xsl:element>
								<xsl:element name="POSTCODE"><xsl:value-of select="NAMEANDADDRESSDIRECTORY/ADDRESS/@POSTCODE"/></xsl:element>
							</xsl:element>
						
							<xsl:element name="CONTACTDETAILS">
								<xsl:element name="CONTACTDETAILSGUID"><xsl:value-of select="CONTACTDETAILS/@CONTACTDETAILSGUID"/></xsl:element>
								<xsl:element name="CONTACTFORENAME"><xsl:value-of select="CONTACTDETAILS/@CONTACTFORENAME"/></xsl:element>
								<xsl:element name="CONTACTSURNAME"><xsl:value-of select="CONTACTDETAILS/@CONTACTSURNAME"/></xsl:element>
								<xsl:element name="CONTACTTITLE"><xsl:value-of select="CONTACTDETAILS/@CONTACTTITLE"/></xsl:element>
								<xsl:element name="CONTACTTYPE"><xsl:value-of select="CONTACTDETAILS/@CONTACTTYPE"/></xsl:element>
								<xsl:element name="EMAILADDRESS"><xsl:value-of select="CONTACTDETAILS/@EMAILADDRESS"/></xsl:element>
								<xsl:element name="FAXNUMBER"><xsl:value-of select="CONTACTDETAILS/@FAXNUMBER"/></xsl:element>
								<xsl:element name="TELEPHONEEXTENSIONNUMBER"><xsl:value-of select="CONTACTDETAILS/@TELEPHONEEXTENSIONNUMBER"/></xsl:element>
								<xsl:element name="TELEPHONENUMBER"><xsl:value-of select="CONTACTDETAILS/@TELEPHONENUMBER"/></xsl:element>

								<xsl:for-each select="CONTACTTELEPHONEDETAILS">
									<xsl:if test="@USAGE != '' ">
										<xsl:element name="CONTACTTELEPHONEDETAILS">
											<xsl:element name="CONTACTDETAILSGUID"><xsl:value-of select="@CONTACTDETAILSGUID"/></xsl:element>
											<xsl:element name="TELEPHONESEQNUM"><xsl:value-of select="@TELEPHONESEQNUM"/></xsl:element>
											<xsl:element name="USAGE"><xsl:value-of select="@USAGE"/></xsl:element>
											<xsl:element name="AREACODE"><xsl:value-of select="@AREACODE"/></xsl:element>
											<xsl:element name="TELENUMBER"><xsl:value-of select="@TELENUMBER"/></xsl:element>
											<xsl:element name="EXTENSIONNUMBER"><xsl:value-of select="@EXTENSIONNUMBER"/></xsl:element>
											<xsl:element name="COUNTRYCODE"><xsl:value-of select="@COUNTRYCODE"/></xsl:element>
										</xsl:element>
									</xsl:if>
								</xsl:for-each>
							</xsl:element>
							
						</xsl:element>	
					</xsl:if>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!--END: WriteThirdPartyDetails-->
	<xsl:template name="mortgageClub">
		<xsl:attribute name="CLUBNETWORKASSOCIATIONID"><xsl:value-of select="@CLUBNETWORKASSOCIATIONID"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="@MORTGAGECLUBNETWORKASSOCNAME"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="(@PACKAGERINDICATOR = '1')">
				<xsl:attribute name="TYPEVALIDATION">PA</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="TYPEVALIDATION">MC</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:call-template name="toAddressLines"/>
		<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(@TELEPHONEAREACODE, @TELEPHONENUMBER)"/></xsl:attribute>
		<xsl:attribute name="FAX"><xsl:value-of select="concat(@FAXAREACODE, @FAXNUMBER)"/></xsl:attribute>
	</xsl:template>
	<xsl:template name="principalFirm">
		<xsl:attribute name="PRINCIPALFIRMID"><xsl:value-of select="@PRINCIPALFIRMID"/></xsl:attribute>
		<xsl:attribute name="FSAREFNUMBER"><xsl:value-of select="@FSAREF"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="@PRINCIPALFIRMNAME"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="(@PACKAGERINDICATOR = '1')">
				<xsl:attribute name="TYPEVALIDATION">PKF</xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="TYPEVALIDATION">PFN</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:call-template name="toListingStatus"/>
		<xsl:attribute name="STATUS"><xsl:value-of select="@FIRMSTATUS"/></xsl:attribute>
		<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(@TELEPHONEAREACODE, @TELEPHONENUMBER)"/></xsl:attribute>
		<xsl:attribute name="FAX"><xsl:value-of select="concat(@FAXAREACODE, @FAXNUMBER)"/></xsl:attribute>
		<xsl:call-template name="addressLines"/>
	</xsl:template>
<xsl:template name="arFirm">
		<xsl:attribute name="ARFIRMID"><xsl:value-of select="@ARFIRMID"/></xsl:attribute>
		<xsl:attribute name="FSAREFNUMBER"><xsl:value-of select="@FSAARFIRMREF"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="@ARFIRMNAME"/></xsl:attribute>
		<xsl:attribute name="TYPEVALIDATION">ARF</xsl:attribute>
		<xsl:call-template name="toListingStatus"/>
		<xsl:attribute name="STATUS"><xsl:value-of select="@FIRMSTATUS"/></xsl:attribute>
		<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(@TELEPHONEAREACODE, @TELEPHONENUMBER)"/></xsl:attribute>
		<xsl:attribute name="FAX"><xsl:value-of select="concat(@FAXAREACODE, @FAXNUMBER)"/></xsl:attribute>
		<xsl:call-template name="addressLines"/>
	</xsl:template>
	<xsl:template name="introducer">
		<xsl:attribute name="INTRODUCERID"><xsl:value-of select="@INTRODUCERID"/></xsl:attribute>
		<xsl:attribute name="NAME"><xsl:value-of select="concat(ORGANISATIONUSER/@USERFORENAME,' ', ORGANISATIONUSER/@USERSURNAME)"/></xsl:attribute>
		<xsl:attribute name="TYPE"><xsl:value-of select="@INTRODUCERTYPE"/></xsl:attribute>
		<!-- EP2_863 -->
		<xsl:if test="INTRODUCERFIRM[@TRADINGAS]">
			<xsl:attribute name="TRADINGAS"><xsl:value-of select="INTRODUCERFIRM/@TRADINGAS"/></xsl:attribute>
		</xsl:if>
		<!-- EP2_863  ends-->
		<xsl:choose>
			<xsl:when test="(@INTRODUCERTYPE = '1')">
				<xsl:choose>
					<xsl:when test="(@ARINDICATOR = '1')">
						<xsl:attribute name="TYPEVALIDATION">IBAR</xsl:attribute>
					</xsl:when>
					<xsl:when test="(@ARINDICATOR = '0')">
						<xsl:attribute name="TYPEVALIDATION">IBDA</xsl:attribute>
					</xsl:when>
				</xsl:choose>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="TYPEVALIDATION">IP</xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:call-template name="toListingStatus"/>
		<xsl:attribute name="STATUS"><xsl:value-of select="@INTRODUCERSTATUS"/></xsl:attribute>
		<xsl:call-template name="toAddressLines"/>
		<xsl:attribute name="EMAILADDRESS"><xsl:value-of select="@EMAILADDRESS"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=2]">
				<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=2]/@COUNTRYCODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=2]/@AREACODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=2]/@TELENUMBER)"></xsl:value-of>
				</xsl:attribute>
			</xsl:when>
			<xsl:when test="ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=3]">
				<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=3]/@COUNTRYCODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=3]/@AREACODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=3]/@TELENUMBER)"></xsl:value-of>
				</xsl:attribute>
			</xsl:when>	
			<xsl:when test="ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=1]">
				<xsl:attribute name="TELEPHONE"><xsl:value-of select="concat(ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=1]/@COUNTRYCODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=1]/@AREACODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=1]/@TELENUMBER)"></xsl:value-of>
				</xsl:attribute>				
			</xsl:when>			
		</xsl:choose>
		<xsl:if test="ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=4]">
			<xsl:attribute name="FAX"><xsl:value-of select="concat(ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=4]/@COUNTRYCODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=4]/@AREACODE, ' ', ORGANISATIONUSER/ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS/CONTACTTELEPHONEDETAILS[@USAGE=4]/@TELENUMBER)"></xsl:value-of></xsl:attribute>
		</xsl:if>	
		<xsl:apply-templates select="*"/>
	</xsl:template>
	<xsl:template name="addressLines">
		<xsl:attribute name="ADDRESSLINE1"><xsl:value-of select="@ADDRESSLINE1"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE2"><xsl:value-of select="@ADDRESSLINE2"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE3"><xsl:value-of select="@ADDRESSLINE3"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE4"><xsl:value-of select="@ADDRESSLINE4"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE5"><xsl:value-of select="@ADDRESSLINE5"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE6"><xsl:value-of select="@ADDRESSLINE6"/></xsl:attribute>
		<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
	</xsl:template>
	<xsl:template name="toAddressLines">
		<xsl:attribute name="ADDRESSLINE1"><xsl:value-of select="concat(@FLATNUMBER,' ',@BUILDINGORHOUSENUMBER)"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE2"><xsl:value-of select="@BUILDINGORHOUSENAME"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE3"><xsl:value-of select="@STREET"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE4"><xsl:value-of select="@DISTRICT"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE5"><xsl:value-of select="@TOWN"/></xsl:attribute>
		<xsl:attribute name="ADDRESSLINE6"><xsl:value-of select="@COUNTY"/></xsl:attribute>
		<xsl:attribute name="POSTCODE"><xsl:value-of select="@POSTCODE"/></xsl:attribute>
	</xsl:template>
	<xsl:template name="toListingStatus">
		<xsl:choose>
			<xsl:when test="(@LISTINGSTATUS = '0')">
				<xsl:attribute name="LISTSTATUSVALIDATION">NA</xsl:attribute>
			</xsl:when>
			<xsl:when test="(@LISTINGSTATUS = '1')">
				<xsl:attribute name="LISTSTATUSVALIDATION">BL</xsl:attribute>
			</xsl:when>
			<xsl:when test="(@LISTINGSTATUS = '2')">
				<xsl:attribute name="LISTSTATUSVALIDATION">GL</xsl:attribute>
			</xsl:when>
			<xsl:when test="(@LISTINGSTATUS = '3')">
				<xsl:attribute name="LISTSTATUSVALIDATION">WL</xsl:attribute>
			</xsl:when>
		</xsl:choose>
	</xsl:template>
	
	
</xsl:stylesheet>
