<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control==========================================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/FindIntroducersResponse.xslt $
Workfile:             $Workfile: FindIntroducersResponse.xslt $
Current Version   $Revision: 9 $
Last Modified      $Modtime: 20/04/07 18:18 $
Modified By        $Author: Mheys $

Copyright				Copyright © 2006 Vertex Financial Services
Description			epsom specific FindIntroducer response transformation

History:

Author		Date				Defect		Description
IK				23/10/2006	EP2_21	created
IK				06/11/2006	EP2_32	remove date formatting (now done by omCRUD)
IK				06/11/2006	EP2_32	drop ARFIRM.UNITID, FIRMPERMISSIONS.FIRMPERMISSIONSSEQNO from output
IK				08/11/2006	EP2_35	add FSAONLYINDIVIDUAL search
IK				14/11/2006	EP2_89	fix PACKAGEINDICATOR renaming
IK				22/11/2006	EP2_160	add MORTGAGECLUBNETWORKASSOCIATION to PRINCIPALFIRM
PSC			08/03/2007 	EP2_1190 Amend set up of ARFIRm data
IK				20/04/20070 	EP2_337	add sort to MORTGAGECLUBNETWORKASSOCIATION
================================================================================================-->
<?altova_samplexml C:\omiga4Trace\omResponse.xml?>
<xsl:stylesheet version="2.0" xmlns="http://Response.FindIntroducers.Omiga.vertex.co.uk" xmlns:msgdt="http://msgtypes.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="RESPONSE">
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="MESSAGE"/>
			<xsl:apply-templates select="ERROR"/>
			<xsl:apply-templates select="OPERATION/RESPONSE/MORTGAGECLUBNETWORKASSOCIATION">
				<xsl:sort select="@MORCLUBNETWORKASSOCNAME"/>
			</xsl:apply-templates>
			<xsl:for-each select="OPERATION/RESPONSE/PRINCIPALFIRM">
				<xsl:call-template name="fsaData"/>
			</xsl:for-each>
			<xsl:for-each select="OPERATION/RESPONSE/ARFIRM">
				<xsl:call-template name="fsaData"/>
			</xsl:for-each>
			<xsl:for-each select="OPERATION/RESPONSE/INTRODUCERDETAILS">
				<xsl:element inherit-namespaces="yes" name="{local-name()}">
					<xsl:for-each select="@*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
					<xsl:apply-templates select="*"/>
				</xsl:element>
			</xsl:for-each>
			<xsl:for-each select="OPERATION/RESPONSE/FSAINDIVIDUAL">
				<xsl:element inherit-namespaces="yes" name="{local-name()}">
					<xsl:for-each select="@*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
					<xsl:apply-templates select="*[local-name() = 'INDIVIDUALCONTROLS']"/>
					<xsl:apply-templates select="*[local-name() = 'INDIVIDUALEMPLOYMENT']"/>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MESSAGE">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="MESSAGE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ERROR">
		<xsl:element name="ERROR">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MORTGAGECLUBNETWORKASSOCIATION">
		<xsl:call-template name="fsaData"/>
	</xsl:template>
	<xsl:template match="INDIVIDUALCONTROLS">
		<xsl:call-template name="copyIndControl"/>
	</xsl:template>
	<xsl:template match="INDIVIDUALEMPLOYMENT">
		<xsl:call-template name="copyIndEmploy"/>
	</xsl:template>
	<xsl:template match="INTRODUCERFIRM">
		<xsl:for-each select="PRINCIPALFIRM">
			<xsl:element inherit-namespaces="yes" name="{local-name()}">
				<xsl:for-each select="@*">
					<xsl:copy-of select="."/>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
		<xsl:for-each select="ARFIRM">
			<xsl:element inherit-namespaces="yes" name="{local-name()}">
				<xsl:for-each select="@*">
					<xsl:copy-of select="."/>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template match="ORGANISATIONUSER">
		<xsl:element inherit-namespaces="yes" name="{local-name()}">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
				<!-- switch off for now
				<xsl:choose>
					<xsl:when test="name(.)='USERTITLE'">
						<xsl:choose>
							<xsl:when test="../@USERTITLE_TEXT">
								<xsl:attribute name="USERTITLE"><xsl:value-of select="../@USERTITLE_TEXT"/></xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:copy-of select="."/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="name(.)='USERTITLE_TEXT'"/>
					<xsl:otherwise>
						<xsl:copy-of select="."/>
					</xsl:otherwise>
				</xsl:choose> -->
			</xsl:for-each>
			<xsl:for-each select="ORGANISATIONUSERCONTACTDETAILS/CONTACTDETAILS">
				<xsl:element name="{local-name()}">
					<xsl:for-each select="@*">
						<xsl:copy-of select="."/>
					</xsl:for-each>
					<xsl:for-each select="CONTACTTELEPHONEDETAILS">
						<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="{local-name()}">
							<xsl:for-each select="@*">
								<xsl:copy-of select="."/>
							</xsl:for-each>
						</xsl:element>
					</xsl:for-each>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="fsaData">
		<xsl:element inherit-namespaces="yes" name="{local-name()}">
			<xsl:for-each select="@*">
				<xsl:variable name="ucName">
					<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
				</xsl:variable>
				<xsl:choose>
					<xsl:when test="$ucName='ADDRESSLINE1'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE1'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE2'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE3'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE4'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE5'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='ADDRESSLINE6'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='POSTCODE'">
						<xsl:call-template name="copyAddress"/>
					</xsl:when>
					<xsl:when test="$ucName='TELEPHONECOUNTRYCODE'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='TELEPHONEAREACODE'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='TELEPHONENUMBER'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='FAXCOUNTRYCODE'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='FAXAREACODE'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='FAXNUMBER'">
						<xsl:call-template name="copyPhone"/>
					</xsl:when>
					<xsl:when test="$ucName='BANKACCOUNTNAME'">
						<xsl:call-template name="copyBank"/>
					</xsl:when>
					<xsl:when test="$ucName='BANKACCOUNTNUMBER'">
						<xsl:call-template name="copyBank"/>
					</xsl:when>
					<xsl:when test="$ucName='BANKACCOUNTBRANCHNAME'">
						<xsl:call-template name="copyBank"/>
					</xsl:when>
					<xsl:when test="$ucName='BANKSORTCODE'">
						<xsl:call-template name="copyBank"/>
					</xsl:when>
					<xsl:when test="$ucName='BANKWIZARDINDICATOR'">
						<xsl:call-template name="copyBank"/>
					</xsl:when>
					<xsl:when test="$ucName='AGREEDPROCFEERATE'">
						<xsl:call-template name="copyProcFee"/>
					</xsl:when>
					<xsl:when test="$ucName='VOLPROCFEEADJUSTMENT'">
						<xsl:call-template name="copyProcFee"/>
					</xsl:when>
					<xsl:when test="$ucName='PROCLOADINGIND'">
						<xsl:call-template name="copyProcFee"/>
					</xsl:when>
					<xsl:when test="$ucName='PAYMENTMETHOD'">
						<xsl:call-template name="copyProcFee"/>
					</xsl:when>
					<xsl:when test="$ucName='LISTINGSTATUS'">
						<xsl:call-template name="copyListingStatus"/>
					</xsl:when>
					<xsl:otherwise>
						<xsl:choose>
							<xsl:when test="local-name(..) = 'ARFIRM' and name(.)='UNITID'"/>
							<xsl:when test="local-name(..) = 'FIRMPERMISSIONS' and name(.)='FIRMPERMISSIONSSEQNO'"/>
							<xsl:when test="local-name(..) = 'PRINCIPALFIRM' and $ucName='MORTGAGENETWORKIND'"/>
							<xsl:when test="local-name(..) = 'FIRMPERMISSIONS' and $ucName='FPEFFECTIVEDATE'">
								<xsl:attribute name="EFFECTIVEDATE"><xsl:value-of select="."/></xsl:attribute>
							</xsl:when>
							<xsl:when test="local-name(..) = 'FIRMPERMISSIONS' and $ucName='FPLASTUPDATEDDATE'">
								<xsl:attribute name="LASTUPDATEDDATE"><xsl:value-of select="."/></xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
			<xsl:for-each select="*">
				<xsl:choose>
					<xsl:when test="name(.)='ARFIRM'">
						<xsl:call-template name="copyARFirm"/>
					</xsl:when>
					<!-- EP2_160	 -->
					<xsl:when test="local-name(.) = 'FIRMCLUBNETWORKASSOCIATION'">
						<xsl:for-each select="*">
							<xsl:element inherit-namespaces="yes" name="{local-name()}">
								<xsl:for-each select="@*">
									<xsl:copy-of select="."/>
								</xsl:for-each>
							</xsl:element>
						</xsl:for-each>
					</xsl:when>
					<xsl:when test="local-name()='FIRMPERMISSIONS' and count(@*) = 0"/>
					<xsl:when test="local-name()='PRINCIPALFIRM' and local-name(..) = 'ARFIRM' and not(@MORTGAGENETWORKIND='1') and ancestor::RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[1]/@SELECTEDNETWORK='1' and count(../PRINCIPALFIRM[@MORTGAGENETWORKIND='1']) > 0"/>
					<xsl:otherwise>
						<xsl:call-template name="fsaData"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="msgdtNS">
		<xsl:element namespace="http://msgtypes.Omiga.vertex.co.uk" name="{local-name(.)}">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:value-of select="."/>
			<xsl:for-each select="*">
				<xsl:call-template name="msgdtNS"/>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	<xsl:template name="copyAddress">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:if test="name(..)='MORTGAGECLUBNETWORKASSOCIATION'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@ADDRESS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='PRINCIPALFIRM'">
			<xsl:if test="../@PACKAGERINDICATOR='0'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ADDRESS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="../@PACKAGERINDICATOR='1'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ADDRESS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='ARFIRM'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@ADDRESS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyPhone">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:if test="name(..)='MORTGAGECLUBNETWORKASSOCIATION'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@TELEPHONECONTACTDETAILS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='PRINCIPALFIRM'">
			<xsl:if test="../@PACKAGERINDICATOR='0'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@TELEPHONECONTACTDETAILS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="../@PACKAGERINDICATOR='1'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@TELEPHONECONTACTDETAILS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='ARFIRM'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@TELEPHONECONTACTDETAILS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyBank">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:if test="name(..)='MORTGAGECLUBNETWORKASSOCIATION'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@BANKDETAILS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='PRINCIPALFIRM'">
			<xsl:if test="../@PACKAGERINDICATOR='0'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@BANKDETAILS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="../@PACKAGERINDICATOR='1'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@BANKDETAILS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyProcFee">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:if test="name(..)='MORTGAGECLUBNETWORKASSOCIATION'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/MORTGAGECLUBLENGTH[@PROCFEEDATA='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='PRINCIPALFIRM'">
			<xsl:if test="../@PACKAGERINDICATOR='0'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@PROCFEEDATA='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="../@PACKAGERINDICATOR='1'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@PROCFEEDATA='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyListingStatus">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:if test="name(..)='PRINCIPALFIRM'">
			<xsl:if test="../@PACKAGERINDICATOR='0'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/BROKERLENGTH[@LISTINGSTATUS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
			<xsl:if test="../@PACKAGERINDICATOR='1'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@ALL='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/PACKAGERLENGTH[@LISTINGSTATUS='1']">
					<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
				</xsl:if>
			</xsl:if>
		</xsl:if>
		<xsl:if test="name(..)='ARFIRM'">
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@ALL='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/ARFIRMLENGTH[@LISTINGSTATUS='1']">
				<xsl:attribute name="{string($ucName)}"><xsl:value-of select="."/></xsl:attribute>
			</xsl:if>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyIndControl">
		<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@ALL='1']">
			<xsl:call-template name="fsaData"/>
		</xsl:if>
		<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@INDIVIDUALCONTROL='1']">
			<xsl:call-template name="fsaData"/>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyIndEmploy">
		<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@ALL='1']">
			<xsl:call-template name="fsaData"/>
		</xsl:if>
		<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@INDIVIDUALEMPLOYMENT='1']">
			<xsl:call-template name="fsaData"/>
		</xsl:if>
	</xsl:template>
	<xsl:template name="copyARFirm">
		<xsl:choose>
			<xsl:when test="name(..)='INDIVIDUALEMPLOYMENT'">
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@ALL='1']">
					<xsl:call-template name="fsaData"/>
				</xsl:if>
				<xsl:if test="//RESPONSE/REQUEST/LENGTHSEARCH/FSAONLYINDIVIDUALLENGTH[@APPOINTMENTS='1']">
					<xsl:call-template name="fsaData"/>
				</xsl:if>
			</xsl:when>
			<xsl:otherwise>
				<xsl:call-template name="fsaData"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
</xsl:stylesheet>
