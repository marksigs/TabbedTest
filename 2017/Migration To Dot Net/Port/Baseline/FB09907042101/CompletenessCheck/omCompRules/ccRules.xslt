<?xml version="1.0" encoding="UTF-8"?>
<!--
	IK 18/04/2006 EP333
	IK 27/04/2006 EP328
	IK 02/06/2006 EP661
	PE 21/06/2006 EP817 - modified rule 527 to check that the NEWPROPERTYINDICATOR is set to 1 instead of just checking its presence.
	IK 05/02/2007 EP2_1212 - modify rule 517 to test for EMP & EMPT
	LH 07/02/2007 EP2_1247 - Changed rule 508 according to spec change.
	AShaw 07/02/2007 EP385 - Rule 508 should return to one of two different forms, dependant on where error is. 
	LH 12/02/2007 EP2_1315 - Changed rule 508 according to spec change.
	AShaw 12/02/2007 EP1293 - Rules 532 & 533 Only run if Application.IsLegalRepToBeUsed = 1. 
	OS	08/03/2007	EP2_1618 - Removed Rule 518 
	AShaw 15/03/2007 EP2_1691 - Rules 501-516 need to run for Guarantors as well as Applicants. 
	PSC	28/03/2007 EP2_1904 - Disable rule 512
	
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="http://vertex.com/omiga4">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<msxsl:script language="JScript" implements-prefix="user"><![CDATA[
		var addressError = 0;
		function initAddressFlag()
		{
			addressError = 0;
			return(addressError);
		}
		function setAddressFlag()
		{
			addressError = 1;
			return(addressError);
		}
		function getAddressFlag()
		{
			return(addressError);
		}
		function addressAge(movedIn,age)
		{
			var agedDate = new Date(parseInt(movedIn.substr(6,4),10),parseInt(movedIn.substr(3,2),10) -1,parseInt(movedIn.substr(0,2),10));
			var now = new Date();
			agedDate.setFullYear(agedDate.getFullYear() + parseInt(age));
			if(agedDate < now) return(1);
			return(0);
		}
	]]></msxsl:script>
	<xsl:template match="REQUEST">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:call-template name="aipRules"/>
			<xsl:if test="APPLICATION/APPLICATIONFACTFIND[NEWPROPERTY]">
				<xsl:for-each select="APPLICATION/APPLICATIONFACTFIND">
					<xsl:call-template name="fmaRules"/>
				</xsl:for-each>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<xsl:template name="aipRules">
		<xsl:for-each select="APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION">
			<xsl:for-each select="CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H]">
				<xsl:call-template name="rule500"/>
				<xsl:call-template name="rule501"/>
			</xsl:for-each>
			<xsl:call-template name="rule502"/>
			<xsl:for-each select="CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H]">
				<xsl:call-template name="rule505">
					<xsl:with-param name="prefix">Current Address : </xsl:with-param>
				</xsl:call-template>
			</xsl:for-each>
			<xsl:for-each select="CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_P]">
				<xsl:call-template name="rule503"/>
				<xsl:call-template name="rule504"/>
				<xsl:call-template name="rule505">
					<xsl:with-param name="prefix">Previous Address : </xsl:with-param>
				</xsl:call-template>
			</xsl:for-each>
			<xsl:call-template name="rule506"/>
			<xsl:call-template name="rule507"/>
			<xsl:call-template name="rule508"/>
			<xsl:call-template name="rule509"/>
			<xsl:for-each select="EMPLOYMENT[not(@DATELEFTORCEASEDTRADING)]">
				<xsl:call-template name="rule511"/>
				<xsl:if test="@EMPLOYMENTSTATUS_TYPE_EMP or @EMPLOYMENTSTATUS_TYPE_CON or @EMPLOYMENTSTATUS_TYPE_SELF">
					<!--<xsl:call-template name="rule512"/>-->
					<xsl:call-template name="rule513"/>
					<xsl:call-template name="rule514"/>
					<xsl:call-template name="rule515"/>
					<xsl:call-template name="rule516"/>
				</xsl:if>
			</xsl:for-each>
		</xsl:for-each>
		<xsl:for-each select="APPLICATION/APPLICATIONFACTFIND">
			<xsl:call-template name="rule517"/>
			<!--xsl:call-template name="rule518"/-->
			<xsl:call-template name="rule519"/>
			<xsl:call-template name="rule520"/>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="rule500">
		<xsl:if test="not(@DATEMOVEDIN)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">500</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule501">
		<xsl:call-template name="addressRules">
			<xsl:with-param name="ruleRef">501</xsl:with-param>
		</xsl:call-template>
	</xsl:template>
	<xsl:template name="rule502">
		<xsl:if test="user:addressAge(string(CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_H]/@DATEMOVEDIN),string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='AddressValidationYears']/@AMOUNT)) = '0'">
			<xsl:choose>
				<xsl:when test="not(CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_P][@DATEMOVEDIN])">
					<xsl:call-template name="ruleResponse">
						<xsl:with-param name="ruleRef">502</xsl:with-param>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:for-each select="CUSTOMERADDRESS[@ADDRESSTYPE_TYPE_P][@DATEMOVEDIN]">
						<xsl:if test="position() =1">
							<xsl:attribute name="test_{position()}"><xsl:value-of select="@DATEMOVEDIN"/></xsl:attribute>
							<xsl:attribute name="res"><xsl:value-of select="user:addressAge(string(@DATEMOVEDIN),string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='AddressValidationYears']/@AMOUNT))"/></xsl:attribute>
							<xsl:if test="user:addressAge(string(@DATEMOVEDIN),string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='AddressValidationYears']/@AMOUNT)) = '0'">
								<xsl:call-template name="ruleResponse">
									<xsl:with-param name="ruleRef">502</xsl:with-param>
								</xsl:call-template>
							</xsl:if>
						</xsl:if>
					</xsl:for-each>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule503">
		<xsl:call-template name="addressRules">
			<xsl:with-param name="ruleRef">503</xsl:with-param>
		</xsl:call-template>
	</xsl:template>
	<xsl:template name="rule504">
		<xsl:if test="not(@DATEMOVEDIN)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">504</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule505">
		<xsl:param name="prefix"/>
		<xsl:if test="not(@NATUREOFOCCUPANCY)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">505</xsl:with-param>
				<xsl:with-param name="prefix" select="$prefix"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule506">
		<xsl:if test="not(@DATEOFBIRTH)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">506</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule507">
		<xsl:if test="not(@MARITALSTATUS)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">507</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule508">
		<xsl:if test="string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='DCDependantsAtApplicationLevel']/@BOOLEAN) = '1'">
			<xsl:if test="not(//REQUEST/APPLICATION[@NUMBEROFDEPENDANTS]) or (//REQUEST/APPLICATION[@NUMBEROFDEPENDANTS &lt; '0'])">
				<xsl:call-template name="ruleResponse">
					<xsl:with-param name="ruleRef">508</xsl:with-param>
				</xsl:call-template>
			</xsl:if>
		</xsl:if>
		<xsl:if test="string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='DCDependantsAtApplicationLevel']/@BOOLEAN) = '0'">
			<xsl:if test="not(//REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION[@NUMBEROFDEPENDANTS]) or //REQUEST/APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION[@NUMBEROFDEPENDANTS &lt; '0']">
				<xsl:call-template name="ruleResponse">
					<xsl:with-param name="ruleRef">508</xsl:with-param>
				</xsl:call-template>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule509">
		<xsl:if test="not(CUSTOMERTELEPHONENUMBER[@USAGE_TYPE_H or @USAGE_TYPE_W or @USAGE_TYPE_M][@AREACODE][@TELEPHONENUMBER])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">509</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule511">
		<xsl:if test="not(@EMPLOYMENTSTATUS)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">511</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule512">
		<xsl:if test="not(@EMPLOYMENTTYPE)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">512</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule513">
		<xsl:if test="(@EMPLOYMENTSTATUS_TYPE_EMP or @EMPLOYMENTSTATUS_TYPE_SELF or @EMPLOYMENTSTATUS_TYPE_CON) and not(@DATESTARTEDORESTABLISHED)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">513</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule514">
		<xsl:if test="not(NAMEANDADDRESSDIRECTORY/@COMPANYNAME) and not(THIRDPARTY/@COMPANYNAME)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">514</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule515">
		<xsl:choose>
			<xsl:when test="not(NAMEANDADDRESSDIRECTORY/ADDRESS) and not(THIRDPARTY/ADDRESS)">
				<xsl:call-template name="ruleResponse">
					<xsl:with-param name="ruleRef">515</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="NAMEANDADDRESSDIRECTORY/ADDRESS">
						<xsl:for-each select="NAMEANDADDRESSDIRECTORY[ADDRESS]">
							<xsl:call-template name="addressRules">
								<xsl:with-param name="ruleRef">515</xsl:with-param>
							</xsl:call-template>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<xsl:for-each select="THIRDPARTY[ADDRESS]">
							<xsl:call-template name="addressRules">
								<xsl:with-param name="ruleRef">515</xsl:with-param>
							</xsl:call-template>
						</xsl:for-each>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule516">
		<xsl:if test="not(NAMEANDADDRESSDIRECTORY/CONTACTTELEPHONEDETAILS/@TELENUMBER) and not(THIRDPARTY/CONTACTTELEPHONEDETAILS/@TELENUMBER)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">516</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule517">
		<xsl:choose>
			<xsl:when test="CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_NOEMPLOYER]"/>
			<!-- EP2_1212 -->
			<xsl:when test="CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_EMP or @EMPLOYMENTSTATUS_TYPE_EMPT or @EMPLOYMENTSTATUS_TYPE_SELF or @EMPLOYMENTSTATUS_TYPE_CON][@NETMONTHLYINCOME]"/>
			<!-- EP328 -->
			<xsl:when test="CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_EMP or @EMPLOYMENTSTATUS_TYPE_EMPT or @EMPLOYMENTSTATUS_TYPE_SELF or @EMPLOYMENTSTATUS_TYPE_CON]/EARNEDINCOME[@EARNEDINCOMEAMOUNT != '0']"/>
			<xsl:when test="CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/EMPLOYMENT[@MAINSTATUS='1'][@EMPLOYMENTSTATUS_TYPE_EMP or @EMPLOYMENTSTATUS_TYPE_EMPT or @EMPLOYMENTSTATUS_TYPE_SELF or @EMPLOYMENTSTATUS_TYPE_CON]/NETPROFIT[@YEAR1AMOUNT != '0' or @YEAR2AMOUNT != '0' or @YEAR3AMOUNT != '0']"/>
			<!-- EP328 ends -->
			<xsl:otherwise>
				<xsl:call-template name="ruleResponse"><xsl:with-param name="ruleRef">517</xsl:with-param></xsl:call-template>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule518">
		<xsl:if test="not(CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION[@TIMEATBANKYEARS or @TIMEATBANKMONTHS])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">518</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule519">
		<xsl:if test="@TYPEOFAPPLICATION_TYPE_R">
			<xsl:choose>
				<xsl:when test="(CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT[@DIRECTORYGUID or @THIRDPARTYGUID][MORTGAGEACCOUNT[@REMORTGAGEINDICATOR='1']][MORTGAGELOAN[@MONTHLYREPAYMENT][@OUTSTANDINGBALANCE]])"/>
				<xsl:when test="(CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT[@DIRECTORYGUID or @THIRDPARTYGUID][MORTGAGEACCOUNT[@REMORTGAGEINDICATOR='1'][@TOTALMONTHLYCOST][@TOTALCOLLATERALBALANCE]])"/>
				<xsl:otherwise>
					<xsl:call-template name="ruleResponse">
						<xsl:with-param name="ruleRef">519</xsl:with-param>
					</xsl:call-template>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule520">
		<xsl:if test="FINANCIALSUMMARY[@LOANLIABILITYINDICATOR='1']">
			<xsl:choose>
				<xsl:when test="not(CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT[@DIRECTORYGUID or @THIRDPARTYGUID]/LOANSLIABILITIES)">
					<xsl:call-template name="ruleResponse">
						<xsl:with-param name="ruleRef">520</xsl:with-param>
					</xsl:call-template>
				</xsl:when>
				<xsl:otherwise>
					<xsl:if test="CUSTOMERROLE[@CUSTOMERROLETYPE_TYPE_A]/CUSTOMERVERSION/ACCOUNTRELATIONSHIP/ACCOUNT/LOANSLIABILITIES[not(@MONTHLYREPAYMENT) or not(@TOTALOUTSTANDINGBALANCE)]">
						<xsl:call-template name="ruleResponse">
							<xsl:with-param name="ruleRef">520</xsl:with-param>
						</xsl:call-template>
					</xsl:if>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
	<xsl:template name="fmaRules">
		<xsl:call-template name="rule521"/>
		<xsl:call-template name="rule522"/>
		<xsl:call-template name="rule523"/>
		<xsl:call-template name="rule524"/>
		<xsl:call-template name="rule525"/>
		<xsl:call-template name="rule526"/>
		<xsl:call-template name="rule527"/>
		<xsl:call-template name="rule528"/>
		<xsl:call-template name="rule529"/>
		<!--
		<xsl:call-template name="rule530"/>
		<xsl:call-template name="rule531"/>
		-->
		<xsl:if test="../@ISLEGALREPTOBEUSED = '1'">
			<xsl:call-template name="rule532"/>
		</xsl:if>
		<xsl:if test="../@ISLEGALREPTOBEUSED = '1'">
			<xsl:call-template name="rule533"/>
		</xsl:if>
		<xsl:call-template name="rule534"/>
		<xsl:call-template name="rule535"/>
		<xsl:call-template name="rule536"/>
		<xsl:call-template name="rule537"/>
		<xsl:call-template name="rule538"/>
	</xsl:template>
	<xsl:template name="rule521">
		<xsl:if test="not(NEWPROPERTY/@PROPERTYLOCATION)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">521</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule522">
		<xsl:choose>
			<xsl:when test="not(NEWPROPERTYADDRESS/ADDRESS)">
				<xsl:call-template name="ruleResponse">
					<xsl:with-param name="ruleRef">522</xsl:with-param>
				</xsl:call-template>
			</xsl:when>
			<xsl:otherwise>
				<xsl:for-each select="NEWPROPERTYADDRESS[ADDRESS]">
					<xsl:call-template name="addressRules">
						<xsl:with-param name="ruleRef">522</xsl:with-param>
					</xsl:call-template>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule523">
		<xsl:if test="NEWPROPERTYADDRESS/ADDRESS[@POSTCODE] and (not(NEWPROPERTYADDRESS/ADDRESS[@PAFINDICATOR]) or NEWPROPERTYADDRESS/ADDRESS[@PAFINDICATOR='0'])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">523</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule524">
		<xsl:if test="not(NEWPROPERTY/@TYPEOFPROPERTY)">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">524</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule525">
		<xsl:if test="not(NEWPROPERTYROOMTYPE[@ROOMTYPE_TYPE_BD][@NUMBEROFROOMS])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">525</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule526">
		<xsl:if test="not(NEWPROPERTY[@TENURETYPE])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">526</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule527">
		<xsl:if test="(NEWPROPERTY[@NEWPROPERTYINDICATOR='1'][not(@HOUSEBUILDERSGUARANTEE)])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">527</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule528">
		<xsl:choose>
			<xsl:when test="NEWPROPERTYVENDOR/*/CONTACTDETAILS[@CONTACTFORENAME or @CONTACTSURNAME]"/>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="NEWPROPERTYADDRESS[@ACCESSCONTACTNAME]"/>
					<xsl:otherwise>
						<xsl:call-template name="ruleResponse">
							<xsl:with-param name="ruleRef">528</xsl:with-param>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule529">
		<xsl:choose>
			<xsl:when test="NEWPROPERTYVENDOR/*/CONTACTTELEPHONEDETAILS[@TELENUMBER]"/>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="NEWPROPERTYADDRESS[@ACCESSTELEPHONENUMBER]"/>
					<xsl:otherwise>
						<xsl:call-template name="ruleResponse">
							<xsl:with-param name="ruleRef">529</xsl:with-param>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule530">
		<xsl:if test="NEWPROPERTY[not(@ANYOTHERRESIDENTSINDICATOR)]">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">530</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule531">
		<xsl:if test="NEWPROPERTY[@ANYOTHERRESIDENTSINDICATOR]">
			<xsl:if test="not(OTHERRESIDENT/PERSON[not(@TITLE) and not(@FIRSTFORENAME) and not(@SURNAME)])">
				<xsl:call-template name="ruleResponse">
					<xsl:with-param name="ruleRef">531</xsl:with-param>
				</xsl:call-template>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule532">
		<xsl:if test="not(APPLICATIONLEGALREP/THIRDPARTY[@COMPANYNAME]) and not(APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY[@COMPANYNAME])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">532</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule533">
		<xsl:choose>
			<xsl:when test="APPLICATIONLEGALREP/*/ADDRESS[@STREET][@TOWN]"/>
			<xsl:otherwise>
				<xsl:choose>
					<xsl:when test="APPLICATIONLEGALREP/CONTACTTELEPHONEDETAILS[@TELENUMBER]"/>
					<xsl:when test="APPLICATIONLEGALREP/*/CONTACTTELEPHONEDETAILS[@TELENUMBER]"/>
					<xsl:otherwise>
						<xsl:call-template name="ruleResponse">
							<xsl:with-param name="ruleRef">533</xsl:with-param>
						</xsl:call-template>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<xsl:template name="rule534">
		<xsl:if test="not(APPLICATIONBANKBUILDINGSOC/THIRDPARTY[@THIRDPARTYBANKSORTCODE]) and not(APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY[@NAMEANDADDRESSBANKSORTCODE])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">534</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule535">
		<xsl:if test="not(APPLICATIONBANKBUILDINGSOC[@ACCOUNTNUMBER])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">535</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule536">
		<xsl:if test="not(APPLICATIONBANKBUILDINGSOC[@ACCOUNTNAME])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">536</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule537">
		<xsl:if test="not(NEWPROPERTY[@VALUATIONTYPE])">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">537</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="rule538">
		<xsl:if test="FEEPAYMENT[@FEETYPE_TYPE_VAL][@PAYMENTEVENT_TYPE_AW]/PAYMENTRECORD[@PAYMENTMETHOD_TYPE_CD]">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef">538</xsl:with-param>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="addressRules">
		<xsl:param name="ruleRef"/>
		<xsl:param name="prefix"/>
		<xsl:if test="user:initAddressFlag()"/>
		<xsl:choose>
			<xsl:when test="not(ADDRESS)">
				<xsl:if test="user:setAddressFlag()"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:for-each select="ADDRESS">
					<xsl:choose>
						<xsl:when test="@POSTCODE">
							<xsl:choose>
								<xsl:when test="@BUILDINGORHOUSENUMBER"/>
								<xsl:when test="@BUILDINGORHOUSENAME"/>
								<xsl:when test="@FLATNUMBER and @BUILDINGORHOUSENAME"/>
								<xsl:otherwise>
									<xsl:if test="user:setAddressFlag()"/>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test="not(@DISTRICT or @TOWN)">
									<xsl:if test="user:setAddressFlag()"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:choose>
										<xsl:when test="@BUILDINGORHOUSENUMBER and @STREET"/>
										<xsl:when test="@BUILDINGORHOUSENAME and @STREET"/>
										<xsl:when test="@FLATNUMBER and @BUILDINGORHOUSENAME"/>
										<xsl:otherwise>
											<xsl:if test="user:setAddressFlag()"/>
										</xsl:otherwise>
									</xsl:choose>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:if test="user:getAddressFlag() = 1">
			<xsl:call-template name="ruleResponse">
				<xsl:with-param name="ruleRef" select="$ruleRef"/>
				<xsl:with-param name="prefix" select="$prefix"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="ruleResponse">
		<xsl:param name="ruleRef"/>
		<xsl:param name="prefix"/>
		<xsl:element name="COMPLETIONRULE">
			<xsl:attribute name="RULEID"><xsl:value-of select="$ruleRef"/></xsl:attribute>
			<xsl:if test="$ruleRef=500">
				<xsl:attribute name="DESCRIPTION">Current Address: Missing date of taking up residence at current address for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=501">
				<xsl:attribute name="DESCRIPTION">Current Address: Insufficient information in current address for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=502">
				<xsl:attribute name="DESCRIPTION">Address: Addresses missing for the last <xsl:value-of select="//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='AddressValidationYears']/@AMOUNT"/> years, for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=503">
				<xsl:attribute name="DESCRIPTION">Previous Address: Insufficient information in a previous address for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=504">
				<xsl:attribute name="DESCRIPTION">Previous Address: Missing date of taking up residence for address in the Address Summary for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=505">
				<xsl:attribute name="DESCRIPTION"><xsl:value-of select="$prefix"/>Nature of Occupancy missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=506">
				<xsl:attribute name="DESCRIPTION">Personal Details: Date of Birth Missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=507">
				<xsl:attribute name="DESCRIPTION">Personal Details: Marital Status missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=508">
				<xsl:attribute name="DESCRIPTION">Number of dependants not specified.</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=509">
				<xsl:attribute name="DESCRIPTION">Telephone Number: Minimum of one Phone Number not specified for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=511">
				<xsl:attribute name="DESCRIPTION">Personal Details: Employment Status missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=512">
				<xsl:attribute name="DESCRIPTION">Personal Details: Employment Type missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=513">
				<xsl:attribute name="DESCRIPTION">Employment: Date started Current Employment missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=514">
				<xsl:attribute name="DESCRIPTION">Employment: Employer name missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=515">
				<xsl:attribute name="DESCRIPTION">Employment: Employer Address Incomplete for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=516">
				<xsl:attribute name="DESCRIPTION">Employment: Employer telephone number missing for <xsl:call-template name="custName"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=517">
				<xsl:attribute name="DESCRIPTION">Employment: Basic Annual Income / Net Profit required</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=518">
				<xsl:attribute name="DESCRIPTION">Personal Details: Time at Bank must be entered for at least one Applicant</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=519">
				<xsl:attribute name="DESCRIPTION">Application: Existing Remortgage Account and Loan details not added</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=520">
				<xsl:attribute name="DESCRIPTION">Application: Loans/Liabilities not added</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=521">
				<xsl:attribute name="DESCRIPTION">Property: Please select the property location </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=522">
				<xsl:attribute name="DESCRIPTION">Property: Please enter valid property address </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=523">
				<xsl:attribute name="DESCRIPTION">Property: PostCode has not been validated. Please amend </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=524">
				<xsl:attribute name="DESCRIPTION">Property: Please select the property type </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=525">
				<xsl:attribute name="DESCRIPTION">Property: Please enter the number of bedrooms </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=526">
				<xsl:attribute name="DESCRIPTION">Property: Please select the property tenure </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=527">
				<xsl:attribute name="DESCRIPTION">Property: Please select the type of building standards indemnity that applies to the newly-built property </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=528">
				<xsl:attribute name="DESCRIPTION">Property: Please enter the name of the Estate Agent or other contact for the Valuer </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=529">
				<xsl:attribute name="DESCRIPTION">Property: Please enter the telephone number of the Estate Agent or other contact for the Valuer </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=530">
				<xsl:attribute name="DESCRIPTION">Property: Please indicate whether there will be other residents aged 17 years or over </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=531">
				<xsl:attribute name="DESCRIPTION">Property: Please enter Title, First Name and Last Name for each other resident aged 17 years or over </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=532">
				<xsl:attribute name="DESCRIPTION">Solicitor Details: Please enter the firm’s name for the Solicitor or Licenced Conveyancer </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=533">
				<xsl:attribute name="DESCRIPTION">Solicitor Details: Please enter the contact telephone number or address for the Solicitor or Licenced Conveyancer</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=534">
				<xsl:attribute name="DESCRIPTION">Payment Account: Please enter Payment Bank Sort Code</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=535">
				<xsl:attribute name="DESCRIPTION">Payment Account: Please enter Payment Bank Account Number</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=536">
				<xsl:attribute name="DESCRIPTION">Payment Account: Please enter Payment Bank Account Holder Name</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=537">
				<xsl:attribute name="DESCRIPTION">Property: Please select a valuation scheme </xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=538">
				<xsl:attribute name="DESCRIPTION">Application: Valuation fee credit card payment not complete</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=500">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=501">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=502">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=503">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=504">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=505">
				<xsl:attribute name="SCREENID">DC060</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=506">
				<xsl:attribute name="SCREENID">DC030</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=507">
				<xsl:attribute name="SCREENID">DC030</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=508">
				<xsl:if test="string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='DCDependantsAtApplicationLevel']/@BOOLEAN) = '0'">
					<xsl:attribute name="SCREENID">DC030</xsl:attribute>
				</xsl:if>
				<xsl:if test="string(//REQUEST/APPLICATION/GLOBALPARAMETER[@NAME='DCDependantsAtApplicationLevel']/@BOOLEAN) = '1'">
					<xsl:attribute name="SCREENID">DC010</xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="$ruleRef=509">
				<xsl:attribute name="SCREENID">DC030</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=511">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=512">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=513">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=514">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=515">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=516">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=517">
				<xsl:attribute name="SCREENID">DC160</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=518">
				<xsl:attribute name="SCREENID">DC030</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=519">
				<xsl:attribute name="SCREENID">DC085</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=520">
				<xsl:attribute name="SCREENID">DC085</xsl:attribute>
			</xsl:if>
			<!-- IK 02/06/2006 EP661 -->
			<xsl:if test="$ruleRef=521">
				<xsl:attribute name="SCREENID">DC200</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=522">
				<xsl:attribute name="SCREENID">DC210</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=523">
				<xsl:attribute name="SCREENID">DC210</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=524">
				<xsl:attribute name="SCREENID">DC220</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=525">
				<xsl:attribute name="SCREENID">DC220</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=526">
				<xsl:attribute name="SCREENID">DC220</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=527">
				<xsl:attribute name="SCREENID">DC220</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=528">
				<xsl:attribute name="SCREENID">DC210</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=529">
				<xsl:attribute name="SCREENID">DC215</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=530">
				<xsl:attribute name="SCREENID">DC201</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=531">
				<xsl:attribute name="SCREENID">DC235</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=532">
				<xsl:attribute name="SCREENID">DC280</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=533">
				<xsl:attribute name="SCREENID">DC280</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=534">
				<xsl:attribute name="SCREENID">DC270</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=535">
				<xsl:attribute name="SCREENID">DC270</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=536">
				<xsl:attribute name="SCREENID">DC270</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=537">
				<xsl:attribute name="SCREENID">DC200</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=538">
				<xsl:attribute name="SCREENID">PP010</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=506">
				<xsl:attribute name="COSTMODELLINGIND">Y</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=511">
				<xsl:attribute name="COSTMODELLINGIND">Y</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=512">
				<xsl:attribute name="COSTMODELLINGIND">Y</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=517">
				<xsl:attribute name="COSTMODELLINGIND">Y</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=521">
				<xsl:attribute name="COSTMODELLINGIND">Y</xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=500">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=501">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=502">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=503">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=504">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=505">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=506">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=507">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=508">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=509">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=511">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=512">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=513">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=514">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=515">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$ruleRef=516">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<xsl:template name="custName">
		<xsl:choose>
			<xsl:when test="name()='CUSTOMERVERSION'">
				<xsl:value-of select="@FIRSTFORENAME"/>&#160;<xsl:value-of select="@SURNAME"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:value-of select="ancestor::CUSTOMERVERSION/@FIRSTFORENAME"/>&#160;<xsl:value-of select="ancestor::CUSTOMERVERSION/@SURNAME"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
