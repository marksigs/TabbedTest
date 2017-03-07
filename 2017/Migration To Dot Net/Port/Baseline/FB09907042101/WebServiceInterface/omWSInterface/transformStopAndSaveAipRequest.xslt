<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\xml\defect 1864\wsInterfaceRequest.xml?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog				Date				Description
Peter Edney	10/03/2006		MAR1357 - Addedd LOANREPAYMENTINDICATOR attribute to LOANSLIABILITIES node.
Helen Aldred   15/03/2006      MAR1385 - Added CustomerRoleType to AccountRelationship
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:msg="http://Request.SubmitStopAndSaveAiP.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:variable name="firstPass">
		<xsl:for-each select="*">
			<xsl:call-template name="stripNS"/>
		</xsl:for-each>
	</xsl:variable>
	<xsl:template name="stripNS">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="stripNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/">
		<xsl:for-each select="msxsl:node-set($firstPass)/REQUEST">
			<xsl:element name="REQUEST">
				<xsl:attribute name="OPERATION">StopAndSaveAiP</xsl:attribute>
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<xsl:apply-templates select="*"/>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:element name="USERHISTORY">
				<xsl:attribute name="USERID"><xsl:value-of select="ancestor::REQUEST/@USERID"/></xsl:attribute>
				<xsl:attribute name="UNITID"><xsl:value-of select="ancestor::REQUEST/@UNITID"/></xsl:attribute>
			</xsl:element>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
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
			<xsl:element name="INCOMESUMMARY"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ACCOUNT">
		<xsl:element name="ACCOUNTRELATIONSHIP">
			<xsl:if test="@ACCOUNTGUID">
				<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="@ACCOUNTGUID"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="../../../@CUSTOMERNUMBER">
				<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="../../../@CUSTOMERNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="../../../@CUSTOMERVERSIONNUMBER">
				<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="../../../@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
			</xsl:if>
			<xsl:choose>
				<xsl:when test="../../../@CUSTOMERROLETYPE">
					<xsl:attribute name="CUSTOMERROLETYPE"><xsl:value-of select="../../../@CUSTOMERROLETYPE"/></xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="CUSTOMERROLETYPE">1</xsl:attribute>
				</xsl:otherwise>				
			</xsl:choose>
			<xsl:element name="ACCOUNT">
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<xsl:apply-templates select="MORTGAGEACCOUNT"/>
				<xsl:apply-templates select="LOANSLIABILITIES"/>
				<xsl:apply-templates select="ACCOUNTRELATIONSHIP"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MORTGAGEACCOUNT">
		<xsl:apply-templates select="*"/>
		<xsl:element name="MORTGAGEACCOUNT">
			<xsl:if test="../@ACCOUNTGUID">
				<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="../@ACCOUNTGUID"/></xsl:attribute>
			</xsl:if>
			<xsl:for-each select="@*">
				<xsl:choose>
					<xsl:when test="name()='CURRENTADDRESSINDICATOR' and .='1'">
						<xsl:if test="not(../@ADDRESSGUID)">
							<xsl:attribute name="ADDRESSGUID">xpath:../../../CUSTOMERADDRESS[@ADDRESSTYPE='1']/@ADDRESSGUID</xsl:attribute>
						</xsl:if>
					</xsl:when>
					<xsl:otherwise>
						<xsl:copy/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="LOANSLIABILITIES">
		<xsl:apply-templates select="*"/>
		<xsl:element name="LOANSLIABILITIES">
			<xsl:if test="../@ACCOUNTGUID">
				<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="../@ACCOUNTGUID"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="@TOBEREPAIDINDICATOR">
				<xsl:attribute name="LOANREPAYMENTINDICATOR"><xsl:value-of select="@TOBEREPAIDINDICATOR"/></xsl:attribute>
			</xsl:if>			
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ACCOUNTRELATIONSHIP">
		<xsl:element name="OTHERCUSTOMERACCOUNTRELATIONSHIP">
			<xsl:if test="../@ACCOUNTGUID">
				<xsl:attribute name="ACCOUNTGUID"><xsl:value-of select="../@ACCOUNTGUID"/></xsl:attribute>
			</xsl:if>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CUSTOMERADDRESS">
		<xsl:element name="CUSTOMERADDRESS">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:if test="string-length(@DATEMOVEDIN) = 7">
				<xsl:attribute name="DATEMOVEDIN">01/<xsl:value-of select="@DATEMOVEDIN"/></xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="*"/>
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
	<xsl:template match="*">
		<xsl:element name="{name()}">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates select="*"/>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
