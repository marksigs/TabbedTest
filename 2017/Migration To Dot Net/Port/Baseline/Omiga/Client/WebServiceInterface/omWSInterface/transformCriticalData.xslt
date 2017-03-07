<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: transformCriticalData

History:

Version Author     Date       Description
01.00   RFairlie   04/01/2005 MAR981 WS26 (SubmitStopAndSaveAiP) is not returning the UnderWriterDecision
                              in the response - need to send task ids as if they were global parameters.
01.01   RFairlie   11/01/2005 MAR1024 Ensure GLOBALPARAMETER structure is in step with critical data rules.
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="REQUEST">
		<xsl:element name="REQUEST">
			<xsl:for-each select="@*"><xsl:copy/></xsl:for-each>
			<xsl:attribute name="CREDITCHECK_TEST">1</xsl:attribute>
			<xsl:apply-templates select="BEFORE"/>
			<xsl:apply-templates select="AFTER"/>
			<xsl:element name="GLOBALPARAMETER">
				<xsl:attribute name="NAME">TMReprocessCreditCheck</xsl:attribute>
				<xsl:attribute name="STRING">Exp_XMLReprocess</xsl:attribute>
				</xsl:element>
			<xsl:element name="GLOBALPARAMETER">
				<xsl:attribute name="NAME">TMRescoreCreditCheck</xsl:attribute>
				<xsl:attribute name="STRING">Exp_XMLRescore</xsl:attribute>
				</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template match="BEFORE">
		<xsl:element name="BEFORE">
			<xsl:apply-templates select="APPLICATION/APPLICATIONFACTFIND"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="AFTER">
		<xsl:element name="AFTER">
			<xsl:apply-templates select="APPLICATION/APPLICATIONFACTFIND"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:element name="APPLICATION">
			<xsl:attribute name="DIRECTINDIRECTBUSINESS">2</xsl:attribute>
			<xsl:for-each select="@APPLICATIONNUMBER"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@APPLICATIONFACTFINDNUMBER"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@PURCHASEPRICEORESTIMATEDVALUE"><xsl:copy/></xsl:for-each>
			<xsl:variable name="ActiveQuote" select="@ACTIVEQUOTENUMBER"/>
			<xsl:for-each select="QUOTATION[@QUOTATIONNUMBER=$ActiveQuote]">
				<xsl:element name="QUOTATION">
					<xsl:apply-templates select="MORTGAGESUBQUOTE"/>
				</xsl:element>
			</xsl:for-each>
			<xsl:apply-templates select="CUSTOMERROLE[@CUSTOMERROLETYPE='1']"/>
			<xsl:apply-templates select="APPLICATIONBANKBUILDINGSOC"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERROLE">
		<xsl:element name="APPLICATIONCUSTOMERROLE">
			<xsl:for-each select="@*"><xsl:copy/></xsl:for-each>
			<xsl:apply-templates select="CUSTOMER/CUSTOMERVERSION"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERVERSION">
		<xsl:element name="CUSTOMER">
			<xsl:for-each select="../../@CUSTOMERNUMBER"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="../../@CUSTOMERVERSIONNUMBER"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@SURNAME"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@FIRSTFORENAME"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@SECONDFORENAME"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@TITLE"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@DATEOFBIRTH"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@MARITALSTATUS"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@NUMBEROFDEPENDANTS"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@TIMEATBANKYEARS"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@TIMEATBANKMONTHS"><xsl:copy/></xsl:for-each>
			<xsl:apply-templates select="CUSTOMERADDRESS[@ADDRESSTYPE='1']"/>
			<xsl:apply-templates select="EMPLOYMENT"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERADDRESS">
		<xsl:copy>
			<xsl:for-each select="@NATUREOFOCCUPANCY"><xsl:copy/></xsl:for-each>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="EMPLOYMENT">
		<xsl:element name="INCOME">
			<xsl:element name="EMPLOYMENT">
				<xsl:for-each select="@EMPLOYMENTSEQUENCENUMBER"><xsl:copy/></xsl:for-each>
				<xsl:for-each select="@EMPLOYMENTSTATUS"><xsl:copy/></xsl:for-each>
				<xsl:for-each select="@OCCUPATIONTYPE"><xsl:copy/></xsl:for-each>
				<xsl:for-each select="@DATESTARTEDORESTABLISHED"><xsl:copy/></xsl:for-each>
				<xsl:for-each select="@EMPLOYMENTTYPE"><xsl:copy/></xsl:for-each>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATIONBANKBUILDINGSOC">
		<xsl:element name="APPLICATIONBANKBUILDINGSOC">
				<xsl:for-each select="@TIMEATBANK"><xsl:copy/></xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MORTGAGESUBQUOTE">
		<xsl:element name="MORTGAGESUBQUOTE">
			<xsl:for-each select="@MORTGAGESUBQUOTENUMBER"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@AMOUNTREQUESTED"><xsl:copy/></xsl:for-each>
			<xsl:apply-templates select="LOANCOMPONENT"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="LOANCOMPONENT">
		<xsl:copy>
			<xsl:for-each select="@TERMINYEARS"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="@TERMINMONTHS"><xsl:copy/></xsl:for-each>
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>
