<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Archive               $Archive: /Epsom Phase2/2 INT Code/ExternalXML/FindIntroducersResponseGUI.xslt $
Workfile:             $Workfile: FindIntroducersResponseGUI.xslt $
Current Version   $Revision: 6 $
Last Modified      $Modtime: 30/01/07 9:21 $
Modified By        $Author: Dbarraclough $

Copyright				Copyright © 2006 Vertex Financial Services
Description			epsom specific FindIntroducer response transformation

History:

Author		Date		Defect		Description
INR		09/11/2006  	EP2_12    created from FindIntroducersResponse.xslt
INR		08/01/2007  	EP2_543   response should be INTRODUCERDETAILS
INR		09/01/2007  	EP2_684  May have multiple IDs with values in a single ApplicationIntroducer and Validations have changed
INR		10/01/2007  	EP2_543   match INTRODUCERDETAILS by Forename and/or Surname
SR		29/01/2007	EP2_115 modified template for INTRODUCER - logic to get telephone and fax numbers.
IK			29/01/2007	EP2_863 return INTRODUCERFIRM.TRADINGAS
================================================================================-->
<?altova_samplexml U:\EP2_543\ForenameSurnameRequest.xml?>
<xsl:stylesheet version="2.0" xmlns="http://Response.FindIntroducers.Omiga.vertex.co.uk" xmlns:msgdt="http://msgtypes.Omiga.vertex.co.uk" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="no"/>
	<xsl:template match="RESPONSE">
		<xsl:variable name="PACKAGERCRITERIA" select="//REQUEST/SEARCHCRITERIA/@SEARCHREF"/>
		<xsl:variable name="SCREEN" select="//REQUEST/SEARCHCRITERIA/@SCREENREF"/>
		<xsl:variable name="FORENAME" select="//REQUEST/OPERATION/INTRODUCERDETAILS/@FORENAME"/>
		<xsl:variable name="COMPFORENAME" select="translate($FORENAME, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"/>
		<xsl:variable name="SURNAME" select="//REQUEST/OPERATION/INTRODUCERDETAILS/@SURNAME"/>
		<xsl:variable name="COMPSURNAME" select="translate($SURNAME, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"/>
		<xsl:variable name="SELECTBYFULLNAME" select="OPERATION/RESPONSE/INTRODUCERDETAILS[ORGANISATIONUSER[translate(@USERFORENAME,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = $COMPFORENAME] and ORGANISATIONUSER[translate(@USERSURNAME,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = $COMPSURNAME]]"/>

		<xsl:variable name="SELECTBYFORENAME" select="OPERATION/RESPONSE/INTRODUCERDETAILS[ORGANISATIONUSER[translate(@USERFORENAME,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') =$COMPFORENAME]]"/>
		<xsl:variable name="SELECTBYSURNAME" select="OPERATION/RESPONSE/INTRODUCERDETAILS[ORGANISATIONUSER[translate(@USERSURNAME,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = $COMPSURNAME]]"/>
		<xsl:element name="RESPONSE">
			<xsl:for-each select="@*">
				<xsl:copy-of select="."/>
			</xsl:for-each>
			<xsl:apply-templates select="MESSAGE"/>
			<xsl:apply-templates select="ERROR"/>
			<xsl:choose>
				<xsl:when test="($SCREEN = 'DC015')">
					<xsl:for-each select="OPERATION/RESPONSE/MORTGAGECLUBNETWORKASSOCIATION">
						<xsl:choose>
							<xsl:when test="@PACKAGERINDICATOR=$PACKAGERCRITERIA">
								<xsl:element name="INTERMEDIARY">
									<xsl:call-template name="mortgageClub"/>
								</xsl:element>
							</xsl:when>
						</xsl:choose>
					</xsl:for-each>
					<xsl:for-each select="OPERATION/RESPONSE/PRINCIPALFIRM">
						<xsl:element name="INTERMEDIARY">
							<xsl:call-template name="principalFirm"/>
						</xsl:element>
					</xsl:for-each>
					<xsl:for-each select="OPERATION/RESPONSE/ARFIRM">
						<xsl:element name="INTERMEDIARY">
							<xsl:call-template name="arFirm"/>
						</xsl:element>
					</xsl:for-each>
					<xsl:choose>
						<xsl:when test="$FORENAME">
							<xsl:choose>
								<xsl:when test="$SURNAME">
									<xsl:for-each select="$SELECTBYFULLNAME">
										<xsl:element name="INTERMEDIARY">
											<xsl:call-template name="introducer"/>
										</xsl:element>
									</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="$SELECTBYFORENAME">
										<xsl:element name="INTERMEDIARY">
											<xsl:call-template name="introducer"/>
										</xsl:element>
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:when>
						<xsl:otherwise>
							<xsl:choose>	
								<xsl:when test="$SURNAME">
										<xsl:for-each select="$SELECTBYSURNAME">
											<xsl:element name="INTERMEDIARY">
												<xsl:call-template name="introducer"/>
											</xsl:element>
										</xsl:for-each>
								</xsl:when>
								<xsl:otherwise>
									<xsl:for-each select="OPERATION/RESPONSE/INTRODUCERDETAILS">
										<xsl:element name="INTERMEDIARY">
											<xsl:call-template name="introducer"/>
										</xsl:element>
									</xsl:for-each>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>

				</xsl:when>
				<xsl:when test="($SCREEN = 'DC010')">
					<xsl:for-each select="//APPLICATIONINTRODUCER">
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
				</xsl:when>
			</xsl:choose>
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
	<xsl:template name="formatMsgDate">
		<xsl:variable name="ucName">
			<xsl:value-of select="translate(name(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
		</xsl:variable>
		<xsl:attribute name="{string($ucName)}"><xsl:variable name="Year"><xsl:value-of select="substring(.,1,4)"/></xsl:variable><xsl:variable name="Month"><xsl:value-of select="substring(.,6,2)"/></xsl:variable><xsl:variable name="Day"><xsl:value-of select="substring(.,9,2)"/></xsl:variable><xsl:value-of select="concat($Day,'/',$Month,'/',$Year)"/></xsl:attribute>
	</xsl:template>
</xsl:stylesheet>
