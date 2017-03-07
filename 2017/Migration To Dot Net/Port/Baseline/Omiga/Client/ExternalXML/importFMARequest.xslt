<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Archive                  $Archive: /Epsom Phase2/2 INT Code/ExternalXML/importFMARequest.xslt $
Workfile                $Workfile: importFMARequest.xslt $
Current Version   	$Revision: 4 $
Last Modified       	$Modtime: 14/04/07 16:52 $
Modified By          	$Author: Mheys $

Prog					Date				Description
IK						09/01/2007	created for EP2_702 - create ACCOUNTRELATIONSHIP records
IK						12/01/2007	EP2_700 - create OTHERARREARSACCOUNT records
IK						12/01/2007	EP2_756 - security address for MORTGAGEACCOUNT
IK						10/04/2007	EP2_1678 - allow DECLINEDMORTGAGE
MCh					13/04/2007	EP2_2186 - Custom ACCOUNT processing to relate TENANCY THIRPARTYGUID to ACCOUNT for ARREARSHISTORY
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:element name="USERHISTORY">
				<xsl:attribute name="CRUD_OP">CREATE</xsl:attribute>
				<xsl:attribute name="USERID">
					<xsl:value-of select="ancestor::REQUEST/@USERID"/>
				</xsl:attribute>
				<xsl:attribute name="UNITID">
					<xsl:value-of select="ancestor::REQUEST/@UNITID"/>
				</xsl:attribute>
			</xsl:element>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="DECISION"/>
	<xsl:template match="CUSTOMERVERSION">
		<xsl:element name="CUSTOMERVERSION">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="CUSTOMERTELEPHONENUMBER"/>
			<xsl:apply-templates select="CUSTOMERADDRESS">
				<xsl:sort select="@ADDRESSTYPE"/>
				<xsl:sort data-type="number" select="concat(substring(@DATEMOVEDIN,7), substring(@DATEMOVEDIN,4,2), substring(@DATEMOVEDIN,1,2))"/>
			</xsl:apply-templates>
			<xsl:apply-templates select="ALIAS"/>
			<xsl:apply-templates select="EMPLOYMENT"/>
			<xsl:apply-templates select="ACCOUNT"/>
			<xsl:apply-templates select="BANKCREDITCARD"/>
			<xsl:apply-templates select="CUSTOMERWRAPUPDETAILS"/>
			<xsl:apply-templates select="INCOMESUMMARY"/>
			<!-- EP2_1678 -->
			<xsl:apply-templates select="DECLINEDMORTGAGE"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="PAYMENTRECORD">
		<xsl:element name="PAYMENTRECORD">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="not(@PAYMENTSEQUENCENUMBER)">
				<xsl:attribute name="PAYMENTSEQUENCENUMBER">2</xsl:attribute>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	<xsl:template match="FEEPAYMENT">
		<xsl:element name="FEEPAYMENT">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="not(@PAYMENTSEQUENCENUMBER)">
				<xsl:attribute name="PAYMENTSEQUENCENUMBER">2</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATIONBANKBUILDINGSOC">
		<xsl:element name="APPLICATIONBANKBUILDINGSOC">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="not(@PROPOSEDPAYMENTMETHOD)">
				<xsl:attribute name="PROPOSEDPAYMENTMETHOD">1</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>

	<!-- EP2_2186 MCh-->
	<xsl:template match="ACCOUNT[@CRUD_OP='CREATE'][ARREARSHISTORY[@DESCRIPTIONOFLOAN='4']]">
		<xsl:element name="ACCOUNT">
			<xsl:copy-of select="@*[name()!='CUSTOMERNUMBER'][name()!='CUSTOMERVERSIONNUMBER'][name()!='CUSTOMERADDRESSSEQUENCENUMBER']"/> 
			<xsl:attribute name="THIRDPARTYGUID">xpath://CUSTOMERADDRESS[@CUSTOMERNUMBER='<xsl:value-of select="@CUSTOMERNUMBER"/>'][@CUSTOMERVERSIONNUMBER='<xsl:value-of select="@CUSTOMERVERSIONNUMBER"/>'][@CUSTOMERADDRESSSEQUENCENUMBER='<xsl:value-of select="@CUSTOMERADDRESSSEQUENCENUMBER"/>'] /TENANCY/THIRDPARTY/@THIRDPARTYGUID</xsl:attribute>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
	<!-- EP2_2186 ends -->	
	
	<!-- EP2_700 -->
	<xsl:template match="ARREARSHISTORY[@CRUD_OP='CREATE'][@DESCRIPTIONOFLOAN='4']">
		<xsl:element name="ARREARSHISTORY">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:element name="OTHERARREARSACCOUNT">
				<xsl:attribute name="DESCRIPTION">RentArrears</xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<!-- EP2_700 ends -->

	<!-- EP2_756 -->
	<xsl:template match="MORTGAGEACCOUNT[ADDRESS]">
		<xsl:element name="MORTGAGEACCOUNT">
			<xsl:for-each select="@*"><xsl:copy/></xsl:for-each>
			<xsl:for-each select="ADDRESS">
				<xsl:element name="ADDRESS">
					<xsl:for-each select="@*"><xsl:copy/></xsl:for-each>

					<xsl:if test="@CRUD_OP='CREATE' or ../@CRUD_OP='CREATE' or ../../@CRUD_OP='CREATE'">
						<xsl:for-each select="../../ACCOUNTRELATIONSHIP">
							<xsl:element name="CUSTOMERADDRESS">
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="@CUSTOMERNUMBER"/></xsl:attribute>
								<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
								<xsl:attribute name="ADDRESSTYPE">5</xsl:attribute>
							</xsl:element>
						</xsl:for-each>
					</xsl:if>

				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!-- EP2_756 ends -->

	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
