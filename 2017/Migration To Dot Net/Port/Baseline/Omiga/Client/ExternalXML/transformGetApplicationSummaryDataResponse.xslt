<?xml version="1.0" encoding="UTF-8"?>
<!-- created for MAR421 ING Defect 731 -->
<?altova_samplexml C:\omiga4Projects\baseline\code\CRUD\omResponse.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:element name="APPLICATION">
			<xsl:element name="APPLICATIONNUMBER">
				<xsl:value-of select="@APPLICATIONNUMBER"/>
			</xsl:element>
			<xsl:element name="APPLICATIONFACTFINDNUMBER">
				<xsl:value-of select="@APPLICATIONFACTFINDNUMBER"/>
			</xsl:element>
			<xsl:apply-templates select="APPLICATION"/>
			<xsl:if test="@APPLICATIONRECOMMENDEDDATE">
				<xsl:element name="APPLICATIONRECOMMENDEDDATE">
					<xsl:value-of select="@APPLICATIONRECOMMENDEDDATE"/>
				</xsl:element>
			</xsl:if>
			<xsl:if test="@APPLICATIONAPPROVALDATE">
				<xsl:element name="APPLICATIONAPPROVALDATE">
					<xsl:value-of select="@APPLICATIONAPPROVALDATE"/>
				</xsl:element>
			</xsl:if>
			<xsl:if test="@OPTOUTINDICATOR">
				<xsl:element name="OPTOUTINDICATOR">
					<xsl:value-of select="@OPTOUTINDICATOR"/>
				</xsl:element>
			</xsl:if>
			<xsl:if test="@TYPEOFAPPLICATION">
				<xsl:element name="TYPEOFAPPLICATION">
					<xsl:value-of select="@TYPEOFAPPLICATION"/>
				</xsl:element>
			</xsl:if>
			<xsl:apply-templates select="APPLICATIONSTAGE[last()]"/>
			<xsl:apply-templates select="REPORTONTITLE"/>
			<xsl:apply-templates select="DISBURSEMENTPAYMENT[last()]"/>
			<xsl:apply-templates select="VALUERINSTRUCTION[last()]"/>
			<xsl:apply-templates select="QUOTATION"/>
			<xsl:apply-templates select="CUSTOMERROLE"/>
			<xsl:apply-templates select="NEWPROPERTY"/>
			<xsl:apply-templates select="APPLICATIONLEGALREP"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATION">
		<xsl:if test="@TYPEOFBUYER">
			<xsl:element name="TYPEOFBUYER">
				<xsl:value-of select="@TYPEOFBUYER"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@APPLICATIONDATE">
			<xsl:element name="APPLICATIONDATE">
				<xsl:value-of select="@APPLICATIONDATE"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="APPLICATIONSTAGE">
		<xsl:element name="APPLICATIONSTAGE">
			<xsl:attribute name="STAGENUMBER"><xsl:value-of select="@STAGENUMBER"/></xsl:attribute>
			<xsl:attribute name="STAGENAME"><xsl:value-of select="@STAGENAME"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<xsl:template match="DISBURSEMENTPAYMENT">
		<xsl:if test="@COMPLETIONDATE">
			<xsl:if test="string(number(@PAYMENTSTATUS))!='97' and string(number(@PAYMENTSTATUS))!='98' and string(number(@PAYMENTSTATUS))!='99'">
			<xsl:element name="ADVANCEDATE">
				<xsl:value-of select="@COMPLETIONDATE"/>
			</xsl:element>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template match="REPORTONTITLE">
		<xsl:if test="@COMPLETIONDATE">
			<xsl:element name="COMPLETIONDATE">
				<xsl:value-of select="@COMPLETIONDATE"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="QUOTATION">
		<xsl:if test="@QUOTATIONNUMBER">
			<xsl:element name="QUOTENUMBER">
				<xsl:value-of select="@QUOTATIONNUMBER"/>
			</xsl:element>
			<xsl:apply-templates select="MORTGAGESUBQUOTE"/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="MORTGAGESUBQUOTE">
		<xsl:if test="@AMOUNTREQUESTED">
			<xsl:element name="AMOUNTREQUESTED">
				<xsl:value-of select="@AMOUNTREQUESTED"/>
			</xsl:element>
			<xsl:apply-templates select="LOANCOMPONENT"/>
		</xsl:if>
	</xsl:template>
	<xsl:template match="LOANCOMPONENT">
		<xsl:element name="LOANCOMPONENT">
			<xsl:element name="LOANCOMPONENTSEQUENCENUMBER">
				<xsl:value-of select="@LOANCOMPONENTSEQUENCENUMBER"/>
			</xsl:element>
			<xsl:element name="MORTGAGEPRODUCTCODE">
				<xsl:value-of select="@MORTGAGEPRODUCTCODE"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERROLE">
		<xsl:element name="APPLICANT">
			<xsl:element name="CUSTOMERNUMBER">
				<xsl:value-of select="@CUSTOMERNUMBER"/>
			</xsl:element>
			<xsl:element name="CUSTOMERVERSIONNUMBER">
				<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>
			</xsl:element>
			<xsl:element name="CUSTOMERORDER">
				<xsl:value-of select="@CUSTOMERORDER"/>
			</xsl:element>
			<xsl:element name="CUSTOMERROLETYPE">
				<xsl:value-of select="@CUSTOMERROLETYPE"/>
			</xsl:element>
			<xsl:apply-templates select="CUSTOMERVERSION"/>
			<xsl:apply-templates select="CUSTOMER"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERVERSION">
		<xsl:if test="@CUSTOMERSTATUS">
			<xsl:element name="CUSTOMERSTATUS">
				<xsl:value-of select="@CUSTOMERSTATUS"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@TITLE">
			<xsl:element name="TITLE">
				<xsl:value-of select="@TITLE"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@TITLEOTHER">
			<xsl:element name="TITLEOTHER">
				<xsl:value-of select="@TITLEOTHER"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@FIRSTFORENAME">
			<xsl:element name="FIRSTFORENAME">
				<xsl:value-of select="@FIRSTFORENAME"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@SECONDFORENAME">
			<xsl:element name="SECONDFORENAME">
				<xsl:value-of select="@SECONDFORENAME"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@OTHERFORENAMES">
			<xsl:element name="OTHERFORENAMES">
				<xsl:value-of select="@OTHERFORENAMES"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@SURNAME">
			<xsl:element name="SURNAME">
				<xsl:value-of select="@SURNAME"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@DATEOFBIRTH">
			<xsl:element name="DATEOFBIRTH">
				<xsl:value-of select="@DATEOFBIRTH"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@MARITALSTATUS">
			<xsl:element name="MARITALSTATUS">
				<xsl:value-of select="@MARITALSTATUS"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@MEMBEROFSTAFF">
			<xsl:element name="MEMBEROFSTAFF">
				<xsl:value-of select="@MEMBEROFSTAFF"/>
			</xsl:element>
		</xsl:if>
		<xsl:apply-templates select="EMPLOYMENT"/>
	</xsl:template>
	<xsl:template match="CUSTOMER">
		<xsl:if test="@OTHERSYSTEMCUSTOMERNUMBER">
			<xsl:element name="OTHERSYSTEMCUSTOMERNUMBER">
				<xsl:value-of select="@OTHERSYSTEMCUSTOMERNUMBER"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="EMPLOYMENT">
		<xsl:if test="@OCCUPATIONTYPE">
			<xsl:element name="OCCUPATIONTYPE">
				<xsl:value-of select="@OCCUPATIONTYPE"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@EMPLOYMENTSTATUS">
			<xsl:element name="EMPLOYMENTSTATUS">
				<xsl:value-of select="@EMPLOYMENTSTATUS"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@DATESTARTEDORESTABLISHED">
			<xsl:element name="DATESTARTEDORESTABLISHED">
				<xsl:value-of select="@DATESTARTEDORESTABLISHED"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@NETMONTHLYINCOME">
			<xsl:element name="NETMONTHLYINCOME">
				<xsl:value-of select="@NETMONTHLYINCOME"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="NEWPROPERTY">
		<xsl:element name="NEWPROPERTY">
			<xsl:if test="@VALUATIONTYPE">
				<xsl:element name="VALUATIONTYPE">
					<xsl:value-of select="@VALUATIONTYPE"/>
				</xsl:element>
			</xsl:if>
			<xsl:apply-templates select="NEWPROPERTYADDRESS"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="NEWPROPERTYADDRESS">
		<xsl:apply-templates select="ADDRESS"/>
		<xsl:if test="@COUNTRYCODE">
			<xsl:element name="COUNTRYCODE">
				<xsl:value-of select="@COUNTRYCODE"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@AREACODE">
			<xsl:element name="AREACODE">
				<xsl:value-of select="@AREACODE"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@TELEPHONENUMBER">
			<xsl:element name="TELEPHONENUMBER">
				<xsl:value-of select="@ACCESSTELEPHONENUMBER"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="ADDRESS">
		<xsl:if test="@FLATNUMBER">
			<xsl:element name="FLATNUMBER">
				<xsl:value-of select="@FLATNUMBER"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@BUILDINGORHOUSENAME">
			<xsl:element name="BUILDINGORHOUSENAME">
				<xsl:value-of select="@BUILDINGORHOUSENAME"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@BUILDINGORHOUSENUMBER">
			<xsl:element name="BUILDINGORHOUSENUMBER">
				<xsl:value-of select="@BUILDINGORHOUSENUMBER"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@STREET">
			<xsl:element name="STREET">
				<xsl:value-of select="@STREET"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@DISTRICT">
			<xsl:element name="DISTRICT">
				<xsl:value-of select="@DISTRICT"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@TOWN">
			<xsl:element name="TOWN">
				<xsl:value-of select="@TOWN"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@COUNTY">
			<xsl:element name="COUNTY">
				<xsl:value-of select="@COUNTY"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@COUNTRY">
			<xsl:element name="COUNTRY">
				<xsl:value-of select="@COUNTRY"/>
			</xsl:element>
		</xsl:if>
		<xsl:if test="@POSTCODE">
			<xsl:element name="POSTCODE">
				<xsl:value-of select="@POSTCODE"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="APPLICATIONLEGALREP">
		<xsl:element name="APPLICATIONLEGALREP">
			<xsl:attribute name="DIRECTORYGUID"><xsl:value-of select="@DIRECTORYGUID"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<xsl:template match="VALUERINSTRUCTION">
		<xsl:if test="@APPOINTMENTDATE">
			<xsl:element name="APPOINTMENTDATE">
				<xsl:value-of select="@APPOINTMENTDATE"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template match="MESSAGE">
		<xsl:copy-of select="."/>
	</xsl:template>
	<xsl:template match="ERROR">
		<xsl:copy-of select="."/>
	</xsl:template>
</xsl:stylesheet>
