<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/transformSubmitAiPResponse.xslt $
Workfile:             $Workfile: transformSubmitAiPResponse.xslt $
Current Version   $Revision: 5 $
Last Modified      $Modtime: 8/12/06 14:47 $
Modified By        $Author: Dbarraclough $

Copyright				Copyright © 2006 Vertex Financial Services
Description			epsom specific SubmitAiP response transformation

History:

Author		Date				    Description
IK				14/03/2006		EP2 epsom specific changes
IK				19/09/2006		remove Epsom identifier
IK				10/10/2006		address targetting response
IK				25/10/2006		associate with EP2_1
IK				25/10/2006		associate with EP2_10 - address targetting
IK				25/11/2006		EP2_202 product cascade
IK				05/12/2006		EP2_287 shopping list
================================================================================-->
<?altova_samplexml C:\omiga4Projects\projectEpsom\xml\ingestion\decisionResponse.xml?>
<xsl:stylesheet version="2.0" xmlns="http://Response.SubmitAiP.Omiga.vertex.co.uk" xmlns:OmigaAiPData="http://OmigaAiPData.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:choose>
				<xsl:when test="ERROR">
					<xsl:element name="RESPONSESTATUS">Error</xsl:element>
					<xsl:element name="ERRORCODE">
						<xsl:value-of select="ERROR/NUMBER"/>
					</xsl:element>
					<xsl:element name="ERRORTEXT">
						<xsl:value-of select="ERROR/SOURCE"/>
						<xsl:choose>
							<xsl:when test="ERROR/DESCRIPTION">.<xsl:value-of select="ERROR/DESCRIPTION"/>
							</xsl:when>
							<xsl:otherwise>.Unspecified Error</xsl:otherwise>
						</xsl:choose>
					</xsl:element>
				</xsl:when>
				<xsl:otherwise>
					<xsl:element name="RESPONSESTATUS">Success</xsl:element>
					<xsl:apply-templates/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template match="TARGETINGDATA">
		<xsl:element name="TARGETINGDATA">
			<xsl:for-each select="ADDRESSTARGETING">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="ADDRESSTARGETING">
					<xsl:value-of select="."/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="ADDRESSMAP">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="ADDRESSMAP">
					<xsl:for-each select="@*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
					<xsl:for-each select="ADDRESS">
						<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="ADDRESS">
							<xsl:for-each select="@*">
								<xsl:copy-of select="."/>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="CCN1LIST">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="CCN1LIST">
					<xsl:for-each select="CCN1">
						<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="CCN1">
							<xsl:for-each select="@*">
								<xsl:copy-of select="."/>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="TARGETDISPLAY">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="TARGETDISPLAY">
					<xsl:value-of select="."/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="DECLAREADDRESSLIST">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="DECLAREADDRESSLIST">
					<xsl:for-each select="DECLAREADDRESS">
						<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="DECLAREADDRESS">
							<xsl:for-each select="@*">
								<xsl:copy-of select="."/>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="ADDRESSTARGET">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="ADDRESSTARGET">
					<xsl:for-each select="@*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="EXP">
				<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="EXP">
					<xsl:for-each select="EXPERIANREF">
						<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="EXPERIANREF">
							<xsl:value-of select="."/>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATIONFACTFIND">
		<xsl:element namespace="http://OmigaAiPData.Omiga.vertex.co.uk" name="APPLICATIONFACTFIND">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="CCN1LIST">
	</xsl:template>
	<xsl:template match="TARGETDISPLAY">
		<xsl:element name="OmigaAiPData:TARGETDISPLAY">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="INGESTIONDECISION">
		<xsl:element name="DECISION">
			<xsl:choose>
				<xsl:when test="@RISKASSESSMENTDECISION">
					<xsl:value-of select="@RISKASSESSMENTDECISION"/>
				</xsl:when>
				<xsl:otherwise>20</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
		<!-- temporary fix, PRODUCTSCHEME not mandatory -->
		<xsl:element name="PRODUCTSCHEME">
			<xsl:choose>
				<xsl:when test="@PRODUCTSCHEME">
					<xsl:value-of select="@PRODUCTSCHEME"/>
				</xsl:when>
				<xsl:otherwise>0</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
		<xsl:element name="OMIGADIPREFERENCENUMBER">
			<xsl:value-of select="@APPLICATIONNUMBER"/>
		</xsl:element>
		<xsl:choose>
			<xsl:when test="@RISKASSESSMENTDECISION='10'">
				<xsl:element name="GENERALPARAGRAPHCODE">10</xsl:element>
			</xsl:when>
			<xsl:otherwise>
				<xsl:element name="REFERRALPARAGRAPHCODE">20</xsl:element>
			</xsl:otherwise>
		</xsl:choose>
		<xsl:apply-templates select="CUSTOMERKYC"/>
	</xsl:template>
	<xsl:template match="CUSTOMERKYC">
		<xsl:variable name="eName">APPLICANT<xsl:value-of select="position()"/>KYCSTATEMENTCODE</xsl:variable>
		<xsl:element name="{$eName}">
			<xsl:choose>
				<xsl:when test="@CUSTOMERKYCSTATUS">
					<xsl:value-of select="@CUSTOMERKYCSTATUS"/>
				</xsl:when>
				<xsl:otherwise>30</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<!-- EP2_287 shopping list -->
	<xsl:template match="SHOPPINGLISTPROFILE">
		<xsl:element namespace="http://Response.SubmitAiP.Omiga.vertex.co.uk" name="SHOPPINGLISTPROFILE">
			<xsl:value-of select="text()"/>
		</xsl:element>
	</xsl:template>
	<!-- EP2_202 -->
	<xsl:template match="MORTGAGEPRODUCTLIST">
		<xsl:element name="PRODUCTCASCADING">
			<xsl:element name="MORTGAGEPRODUCTLIST">
				<xsl:for-each select="@*">
					<xsl:copy/>
				</xsl:for-each>
				<xsl:for-each select="*">
					<xsl:call-template name="mpNameSpace"/>
				</xsl:for-each>
			</xsl:element>
		</xsl:element>
	</xsl:template>
	<xsl:template name="mpNameSpace">
		<xsl:element name="{name()}" namespace="http://Response.FindMortgageProducts.Omiga.vertex.co.uk">
			<xsl:value-of select="text()"/>
			<xsl:for-each select="*">
				<xsl:call-template name="mpNameSpace"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<!-- EP2_202 ends -->
</xsl:stylesheet>
