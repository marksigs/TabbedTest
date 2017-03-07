<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	
	<!-- MAR535 -->
	<xsl:variable name="NewPropertyAddressCompare" select="/RESPONSE/NEWPROPERTYADDRESS/@ADDRESSCOMPARE"/>
	
	<xsl:template match="RESPONSE">
		<xsl:element name="RequestInsuranceRequestType">
			<xsl:apply-templates/>

			<xsl:variable name="MortgageAccountOnProperty" select="MORTGAGEACCOUNT[@ADDRESSCOMPARE=$NewPropertyAddressCompare]"/>
			
			<xsl:variable name="PreviousLender" select="$MortgageAccountOnProperty[@REMORTGAGEINDICATOR='1' and (not(@SECONDCHARGEINDICATOR) or @SECONDCHARGEINDICATOR='0')]"/>
			<xsl:if test="$PreviousLender">
				<xsl:element name="PreviousLender">
					<xsl:element name="Name">
						<xsl:value-of select="$PreviousLender/@COMPANYNAME"/>
					</xsl:element>
					<xsl:element name="AccountNumber">
						<xsl:value-of select="$PreviousLender/@ACCOUNTNUMBER"/>
					</xsl:element>
					<xsl:element name="RedemptionAmount">
						<xsl:choose>
							<xsl:when test="string(number($PreviousLender/@TOTALCOLLATERALBALANCE))='NaN'">0</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="number($PreviousLender/@TOTALCOLLATERALBALANCE) * 100"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
				</xsl:element>
			</xsl:if>

			<xsl:for-each select="$MortgageAccountOnProperty[@SECONDCHARGEINDICATOR='1']">
				<xsl:element name="SecondCharge">
					<xsl:element name="Name">
						<xsl:value-of select="@COMPANYNAME"/>
					</xsl:element>
					<xsl:element name="AccountNumber">
						<xsl:value-of select="@ACCOUNTNUMBER"/>
					</xsl:element>
					<xsl:element name="RedemptionAmount">
						<xsl:choose>
							<xsl:when test="string(number(@TOTALCOLLATERALBALANCE))='NaN'">0</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="number(@TOTALCOLLATERALBALANCE) * 100"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
				</xsl:element>
			</xsl:for-each>

			<xsl:element name="UnencumberedFlag">
				<xsl:choose>
					<xsl:when test="not($MortgageAccountOnProperty)">Y</xsl:when>
					<xsl:otherwise>N</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
		
		</xsl:element>
	</xsl:template>
		
	<xsl:template match="NEWPROPERTY">
		<!-- MAR535 -->
		<xsl:element name="Conveyancer">
			<xsl:element name="Name">
				<xsl:choose>
					<xsl:when test="string(@PROPERTYLOCATIONSCOTLAND)">RA Direct</xsl:when>
					<xsl:otherwise>First Title</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
			<!--The conveyancer address is Not Applicable -->
		</xsl:element>
	</xsl:template>

	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:if test="@EARLIESTCOMPLETIONDATE">
			<xsl:element name="CompletionDate"><xsl:value-of select="@EARLIESTCOMPLETIONDATE"/></xsl:element>
		</xsl:if>
	</xsl:template>

	<xsl:template match="NEWPROPERTYADDRESS">
		<xsl:element name="Property">
			<xsl:element name="Address">
				<xsl:element name="FlatNameOrNumber"><xsl:value-of select="@FLATNUMBER"/></xsl:element>
				<xsl:element name="HouseOrBuildingName"><xsl:value-of select="@BUILDINGORHOUSENAME"/></xsl:element>
				<xsl:element name="HouseOrBuildingNumber"><xsl:value-of select="@BUILDINGORHOUSENUMBER"/></xsl:element>
				<xsl:element name="Street"><xsl:value-of select="@STREET"/></xsl:element>
				<xsl:element name="District"><xsl:value-of select="@DISTRICT"/></xsl:element>
				<xsl:element name="TownOrCity"><xsl:value-of select="@TOWN"/></xsl:element>
				<xsl:element name="County"><xsl:value-of select="@COUNTY"/></xsl:element>
				<xsl:element name="PostCode"><xsl:value-of select="@POSTCODE"/></xsl:element>
				<xsl:element name="AddressType">H</xsl:element>
			</xsl:element>
			<!-- MAR535 -->
			<!--<xsl:variable name="TTCombo" select="$Combo[@GROUPNAME='PropertyTenure' and @VALUEID=$TenureType]"/>-->
			<xsl:element name="Tenure">
				<!--<xsl:value-of select="$TTCombo/@VALUENAME"/></xsl:element>-->
				<xsl:value-of select="/RESPONSE/NEWPROPERTY/@TENURETYPENAME"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>

	<xsl:template match="CUSTOMER">
		<xsl:element name="Borrower">
			<xsl:for-each select="/RESPONSE/CUSTOMERVERSION">
				<xsl:element name="Name">
					<xsl:element name="Title">
						<xsl:if test="string(@TITLE)">
							<xsl:value-of select="substring(@TITLE, 1, 10)"/>
						</xsl:if>
					</xsl:element>
					<xsl:element name="Forename">
						<xsl:variable name="FirstName">
							<xsl:if test="string(@FIRSTFORENAME)">
								<xsl:value-of select="@FIRSTFORENAME"/>
							</xsl:if>
							<xsl:if test="string(@SECONDFORENAME)">
								<xsl:value-of select="concat(' ', @SECONDFORENAME)"/>
							</xsl:if>
							<xsl:if test="string(@OTHERFORENAMES)">
								<xsl:value-of select="concat(' ', @OTHERFORENAMES)"/>
							</xsl:if>
						</xsl:variable>
						<xsl:value-of select="$FirstName"/>
					</xsl:element>
					<xsl:element name="Surname">
						<xsl:value-of select="@SURNAME"/>
					</xsl:element>
				</xsl:element>
			</xsl:for-each>

			<!-- Only include the address if the current address is different to the NewPropertyAddress -->
			<xsl:variable name="CurrentAddress" select="CUSTOMERVERSION/ADDRESS[@ADDRESSTYPE='Current' and @ADDRESSCOMPARE != $NewPropertyAddressCompare]"/>
			<xsl:if test="$CurrentAddress">
				<!-- If a correspondence address exists, then it overrides the current address -->
				<xsl:variable name="CorrespondenceAddress" select="CUSTOMERVERSION/ADDRESS[@ADDRESSTYPE='Correspondence']"/>
				<xsl:choose>
					<xsl:when test="$CorrespondenceAddress">
						<xsl:call-template name="CorrespondAddress"><xsl:with-param name="Address" select="$CorrespondenceAddress"/></xsl:call-template>
					</xsl:when>
					<xsl:otherwise>
						<xsl:call-template name="CorrespondAddress"><xsl:with-param name="Address" select="$CurrentAddress"/></xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
			
			<xsl:variable name="HomePhone" select="/RESPONSE/CUSTOMERHOMETELEPHONENUMBER"/>
			<xsl:if test="$HomePhone">
				<xsl:element name="HomePhone">
					<xsl:element name="AreaCode"><xsl:value-of select="$HomePhone/@AREACODE"/></xsl:element>
					<xsl:element name="LocalNumber"><xsl:value-of select="$HomePhone/@TELEPHONENUMBER"/></xsl:element>
					<xsl:element name="Extension"><xsl:value-of select="$HomePhone/@EXTENSIONNUMBER"/></xsl:element>
					<xsl:element name="PhoneType">R</xsl:element>
				</xsl:element>
			</xsl:if>
			
			<xsl:variable name="WorkPhone" select="/RESPONSE/CUSTOMERWORKTELEPHONENUMBER"/>
			<xsl:if test="$WorkPhone">
				<xsl:element name="WorkPhone">
					<xsl:element name="AreaCode"><xsl:value-of select="$WorkPhone/@AREACODE"/></xsl:element>
					<xsl:element name="LocalNumber"><xsl:value-of select="$WorkPhone/@TELEPHONENUMBER"/></xsl:element>
					<xsl:element name="Extension"><xsl:value-of select="$WorkPhone/@EXTENSIONNUMBER"/></xsl:element>
					<xsl:element name="PhoneType">B</xsl:element>
				</xsl:element>
			</xsl:if>
			
			<xsl:variable name="MobilePhone" select="/RESPONSE/CUSTOMERMOBILETELEPHONENUMBER"/>
			<xsl:if test="$MobilePhone">
				<xsl:element name="MobilePhone">
					<xsl:element name="AreaCode"><xsl:value-of select="$MobilePhone/@AREACODE"/></xsl:element>
					<xsl:element name="LocalNumber"><xsl:value-of select="$MobilePhone/@TELEPHONENUMBER"/></xsl:element>
					<xsl:element name="Extension"><xsl:value-of select="$MobilePhone/@EXTENSIONNUMBER"/></xsl:element>
					<xsl:element name="PhoneType">C</xsl:element>
				</xsl:element>
			</xsl:if>
			
			<xsl:element name="EmailAddress"><xsl:value-of select="CUSTOMERVERSION/@CONTACTEMAILADDRESS"/></xsl:element>	
		</xsl:element>
	</xsl:template>
 
	<xsl:template match="MORTGAGESUBQUOTE">
		<xsl:element name="NewLender">
			<xsl:element name="Name">ING Direct</xsl:element>
			<xsl:element name="OfficeLocation">
				<xsl:value-of select="/RESPONSE/REGION/@REGIONNAME"/>
			</xsl:element> 
			<xsl:element name="ApplicationNo">
				<xsl:value-of select="/RESPONSE/APPLICATIONFACTFIND/@APPLICATIONNUMBER"/>
			</xsl:element>
			<xsl:element name="MortgageAdvance">
				<!-- MAR535 Mortgage advance should be in pence -->
				<xsl:variable name="MortgageAdvance" select="@AMOUNTREQUESTED"/>
				<xsl:value-of select="number($MortgageAdvance) * 100"/> 
			</xsl:element>
		</xsl:element>
	</xsl:template>
	
	<xsl:template name="CorrespondAddress">
		<xsl:param name="Address"/>
		<xsl:element name="Address">
			<xsl:element name="FlatNameOrNumber"><xsl:value-of select="$Address/@FLATNUMBER"/></xsl:element>
			<xsl:element name="HouseOrBuildingName"><xsl:value-of select="$Address/@BUILDINGORHOUSENAME"/></xsl:element>
			<xsl:element name="HouseOrBuildingNumber"><xsl:value-of select="$Address/@BUILDINGORHOUSENUMBER"/></xsl:element>
			<xsl:element name="Street"><xsl:value-of select="$Address/@STREET"/></xsl:element>
			<xsl:element name="District"><xsl:value-of select="$Address/@DISTRICT"/></xsl:element>
			<xsl:element name="TownOrCity"><xsl:value-of select="$Address/@TOWN"/></xsl:element>
			<xsl:element name="County"><xsl:value-of select="$Address/@COUNTY"/></xsl:element>
			<xsl:element name="PostCode"><xsl:value-of select="$Address/@POSTCODE"/></xsl:element>
			<xsl:element name="AddressType">M</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
