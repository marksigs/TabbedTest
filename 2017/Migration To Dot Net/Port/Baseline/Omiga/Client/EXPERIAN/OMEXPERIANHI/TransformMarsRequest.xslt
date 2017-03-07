<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" indent="no"/>

	<xsl:template match="/">
		<xsl:copy>
		  <xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>

	<!-- Rename the root element and add additional child elements -->
	<xsl:template match="FraudCheckRequestType">
		<xsl:element name="EXPERIAN">
			<xsl:element name="ESERIES">
				<xsl:element name="FORM_ID"><xsl:value-of select="Form_ID"/></xsl:element>
			</xsl:element>
			<xsl:if test="Exp_ExperianRef[text() != '0']">
				<xsl:element name="EXP">
					<xsl:element name="EXPERIANREF"><xsl:value-of select="Exp_ExperianRef"/></xsl:element>
				</xsl:element>			
			</xsl:if>
			<xsl:element name="AC01">
				<xsl:element name="CLIENTKEY"><xsl:value-of select="ClientKey"/></xsl:element>
				<xsl:element name="ENTRYPOINT"><xsl:value-of select="EntryPoint"/></xsl:element>
				<xsl:element name="FUNCTION"><xsl:value-of select="Function"/></xsl:element>
			</xsl:element>
			<xsl:apply-templates select="child::*"/>	
		</xsl:element>
	</xsl:template>

	<!-- START: Remove redundant elements -->
	<xsl:template match="Form_ID|Form_ID/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="Exp_ExperianRef|Exp_ExperianRef/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ClientKey|ClientKey/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="EntryPoint|EntryPoint/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="Function|Function/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ClientDevice|ClientDevice/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="TellerID|TellerID/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="TellerPwd|TellerPwd/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ProxyID|ProxyID/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ProxyPwd|ProxyPwd/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="Operator|Operator/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ProductType|ProductType/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="SessionID|SessionID/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="CommunicationChannel|CommunicationChannel/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="CommunicationDirection|CommunicationDirection/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ServiceName|ServiceName/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<xsl:template match="ServiceName|ServiceName/text()">
		<xsl:apply-templates/>	
	</xsl:template>
	<!-- END: Remove redundant elements -->
	
	<xsl:template match="*">
		<xsl:element name="{translate(local-name(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')}">
			<!-- go process attributes and children -->
				<xsl:apply-templates select="@*|node()"/>
			</xsl:element>
	</xsl:template>
	<xsl:template match="@*">
		<xsl:if test="local-name() != 'nil'">
			<xsl:attribute name="{translate(local-name(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')}">
			  <xsl:value-of select="."/>
			</xsl:attribute>
		</xsl:if>
	</xsl:template>
</xsl:stylesheet>

