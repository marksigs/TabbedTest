<?xml version="1.0" encoding="UTF-8"?>
<!--==============================XML Document Control=============================
Description: GenerateTargettingXML

History:

Version Author    Date       	   Description
01.01	LDM			13/04/2006 EP352  Epsom. Updated to deal with new layout. Cut from mars v3 
01.02	SAB		04/05/2006 EP482  Targetted addresses now appear in proper case
01.02 	SAB	   	11/05/2006 EP483  Stores the Nature of Occupancy	
01.03	LDM			25/05/2006  EP618 Force the address details by truncating, to be the right size for experian
01.03	LDM			26/05/2006  EP624 Avoid NaN on addresses
================================================================================-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="http://vertex.com/omiga4">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<msxsl:script language="JScript" implements-prefix="user"><![CDATA[
		function toPC(v)
		{
			var s = " " + v.toLowerCase();
			var a;
			while(a = s.match(/ [a-z]|'[a-z]|-[a-z]|mc[a-z]|Mc[a-z]/))
				s = s.substr(0,a.lastIndex -1) + s.substr(a.lastIndex -1,1).toUpperCase() + s.substring(a.lastIndex);
			s = s.replace(/ Bfpo /g," BFPO ");
			s = s.replace(/ Hms /g," HMS ");
			s = s.replace(/ Po /g," PO ");
			s = s.replace(/ Von /g," von ");
			s = s.replace(/ Van /g," van ");
			s = s.replace(/ De /g," de ");
			s = s.replace(/ La /g," la ");
			return(s.substring(1));
		}
	]]></msxsl:script>
	<xsl:variable name="Root" select="/GEODS"/>
	<xsl:variable name="Omiga" select="$Root/TARGETINGDATA"/>
	<xsl:variable name="CCRequest" select="$Root/TARGETINGDATA/EXPERIAN"/>
	<xsl:variable name="CCBUKResponse" select="$Root/REQUEST"/>
	<xsl:variable name="CCAUKResponse" select="$Root/REQUEST"/>
	<xsl:template match="/">
		<xsl:element name="TARGETINGDATA">
			<xsl:element name="ADDRESSTARGETING">YES</xsl:element>
			<xsl:element name="ADDRESSMAP">
				<xsl:for-each select="$Omiga/CUSTOMER[1]/ADDRESS">
					<xsl:call-template name="WriteAddressMapLine">
						<xsl:with-param name="Customer" select="ancestor::*"/>
						<xsl:with-param name="Address" select="."/>
						<xsl:with-param name="CustomerIndex" select="'1'"/>
						<xsl:with-param name="AddressIndex" select="position()"/>
					</xsl:call-template>
					<xsl:call-template name="WriteMatchingAddressMapLines">
						<xsl:with-param name="AddressID" select="@ID"/>
						<xsl:with-param name="CustomerIndex" select="'1'"/>
						<xsl:with-param name="AddressIndex" select="position()"/>
					</xsl:call-template>
				</xsl:for-each>
				<xsl:for-each select="$Omiga/CUSTOMER[2]/ADDRESS">
					<xsl:if test="not(./@DUPLICATE)">
						<xsl:call-template name="WriteAddressMapLine">
							<xsl:with-param name="Customer" select="ancestor::*"/>
							<xsl:with-param name="Address" select="."/>
							<xsl:with-param name="CustomerIndex" select="'2'"/>
							<xsl:with-param name="AddressIndex" select="position()"/>
						</xsl:call-template>
					</xsl:if>
				</xsl:for-each>
			</xsl:element>
			<xsl:element name="CCN1LIST">
				<xsl:for-each select="$CCBUKResponse/ADDR-UKBUREA">
					<xsl:call-template name="CreateCNN1">
						<xsl:with-param name="BlockSeqNo" select="./@seq"/>
					</xsl:call-template>
				</xsl:for-each>
				<xsl:for-each select="$CCRequest/RESY">
					<xsl:call-template name="CreateAUKCNN1">
						<xsl:with-param name="BlockSeqNo" select="./ADDRESSNUMBER"/>
					</xsl:call-template>
				</xsl:for-each>
			</xsl:element>
			<xsl:for-each select="$CCRequest/RESY">
				<xsl:call-template name="CreateTargetDisplay">
					<xsl:with-param name="BlockSeqNo" select="./ADDRESSNUMBER"/>
					<xsl:with-param name="AddressIndicator" select="./ADDRESSINDICATOR"/>
				</xsl:call-template>
			</xsl:for-each>
			<xsl:element name="DECLAREADDRESSLIST">
				<xsl:for-each select="$CCRequest/RESY">
					<xsl:call-template name="CreateDeclareAddress">
						<xsl:with-param name="BlockSeqNo" select="./ADDRESSNUMBER"/>
					</xsl:call-template>
				</xsl:for-each>
			</xsl:element>
			<xsl:apply-templates select="$CCBUKResponse/ADDR-UKBUREA"/>
			<xsl:apply-templates select="$CCAUKResponse/ADDR-UK"/>
		</xsl:element>
	</xsl:template>
	<xsl:template name="FindMatchingAddress">
		<xsl:param name="AddressID"/>
		<xsl:if test="$Omiga/CUSTOMER/ADDRESS[@DUPLICATE=$AddressID]">
			<xsl:attribute name="MATCHINGADDRESSGUID"><xsl:value-of select="$Omiga/CUSTOMER/ADDRESS[@DUPLICATE=$AddressID]/@ADDRESSGUID"/></xsl:attribute>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateCNN1">
		<xsl:param name="BlockSeqNo"/>
		<xsl:element name="CCN1">
			<xsl:attribute name="BLOCKSEQNUMBER"><xsl:value-of select="$BlockSeqNo"/></xsl:attribute>
			<xsl:attribute name="NAME"><xsl:choose><xsl:when test="$CCRequest/RESY[ADDRESSNUMBER=$BlockSeqNo and NAMESEQUENCE = '01'] and $CCRequest/RESY[ADDRESSNUMBER=$BlockSeqNo and NAMESEQUENCE = '02'] "><xsl:value-of select="concat($CCRequest/NAME[1]/FORENAME, ' ', $CCRequest/NAME[1]/SURNAME, ' and ', $CCRequest/NAME[2]/FORENAME, ' ', $CCRequest/NAME[2]/SURNAME)"/></xsl:when><xsl:when test="$CCRequest/RESY[ADDRESSNUMBER=$BlockSeqNo and NAMESEQUENCE = '01']"><xsl:value-of select="concat($CCRequest/NAME[1]/FORENAME, ' ', $CCRequest/NAME[1]/SURNAME)"/></xsl:when><xsl:when test="$CCRequest/RESY[ADDRESSNUMBER=$BlockSeqNo and NAMESEQUENCE = '02']"><xsl:value-of select="concat($CCRequest/NAME[2]/FORENAME, ' ', $CCRequest/NAME[2]/SURNAME)"/></xsl:when></xsl:choose></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<xsl:template name="CreateAUKCNN1">
		<xsl:param name="BlockSeqNo"/>
		<xsl:if test="$CCAUKResponse/ADDR-UK[@seq=$BlockSeqNo]">
			<xsl:call-template name="CreateCNN1">
				<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
			</xsl:call-template>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateTargetDisplay">
		<xsl:param name="BlockSeqNo"/>
		<xsl:param name="AddressIndicator"/>
		<xsl:if test="($CCAUKResponse/ADDR-UK[@seq=$BlockSeqNo] or $CCBUKResponse/ADDR-UKBUREA[@seq=$BlockSeqNo])">
			<xsl:element name="TARGETDISPLAY">
				<xsl:attribute name="ADDRESSINDICATOR"><xsl:value-of select="$AddressIndicator"/></xsl:attribute>
				<xsl:attribute name="SEQNUMBER"><xsl:value-of select="$BlockSeqNo"/></xsl:attribute>
				<xsl:attribute name="ADDRESSTYPE"><xsl:choose><xsl:when test="$CCAUKResponse/ADDR-UK[@seq=$BlockSeqNo]"><xsl:value-of select="'Ambiguous'"/></xsl:when><xsl:when test="$CCBUKResponse/ADDR-UKBUREA[@seq=$BlockSeqNo]"><xsl:value-of select="'Invalid'"/></xsl:when></xsl:choose></xsl:attribute>
				<xsl:choose>
					<xsl:when test="$CCBUKResponse/ADDR-UKBUREA[@seq=$BlockSeqNo]">
						<xsl:variable name="Chk1respC" select="$CCBUKResponse/CHK1[@seq=$BlockSeqNo]/child::node()='C'"/>
						<xsl:variable name="Chk1respR" select="$CCBUKResponse/CHK1[@seq=$BlockSeqNo]/child::node()='R'"/>
						<xsl:choose>
							<xsl:when test="$Chk1respC or $Chk1respR">
								<xsl:attribute name="ADDRESSRESOLVED">N</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="ADDRESSRESOLVED">Y</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="ADDRESSRESOLVED">N</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateDeclareAddress">
		<xsl:param name="BlockSeqNo"/>
		<xsl:if test="$CCAUKResponse/ADDR-UK[@seq=$BlockSeqNo] or $CCBUKResponse/ADDR-UKBUREA[@seq=$BlockSeqNo]">
			<xsl:choose>
				<xsl:when test="$BlockSeqNo = '01'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[1]"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="$BlockSeqNo = '11'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[2]"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="$BlockSeqNo = '21'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[3]"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="$BlockSeqNo = '02'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[4]"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="$BlockSeqNo = '12'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[5]"/>
					</xsl:call-template>
				</xsl:when>
				<xsl:when test="$BlockSeqNo = '22'">
					<xsl:call-template name="CreateDeclareAddressElement">
						<xsl:with-param name="BlockSeqNo" select="$BlockSeqNo"/>
						<xsl:with-param name="Address" select="$CCRequest/ADDR[6]"/>
					</xsl:call-template>
				</xsl:when>
			</xsl:choose>
		</xsl:if>
	</xsl:template>
	<xsl:template name="CreateDeclareAddressElement">
		<xsl:param name="BlockSeqNo"/>
		<xsl:param name="Address"/>
		<xsl:if test="$Address">
			<xsl:element name="DECLAREADDRESS">
				<xsl:attribute name="BLOCKSEQNUMBER"><xsl:value-of select="$BlockSeqNo"/></xsl:attribute>
				<xsl:attribute name="FLAT"><xsl:value-of select="substring($Address/FLAT, 1, 16)"/></xsl:attribute>
				<xsl:attribute name="HOUSENAME"><xsl:value-of select="substring($Address/HOUSENAME, 1, 26)"/></xsl:attribute>
				<xsl:attribute name="HOUSENUMBER"><xsl:value-of select="substring($Address/HOUSENUMBER, 1, 10)"/></xsl:attribute>
				<xsl:attribute name="STREET"><xsl:value-of select="substring($Address/STREET, 1, 40)"/></xsl:attribute>
				<xsl:attribute name="DISTRICT"><xsl:value-of select="substring($Address/DISTRICT, 1, 30)"/></xsl:attribute>
				<xsl:attribute name="TOWN"><xsl:value-of select="substring($Address/TOWN, 1, 20)"/></xsl:attribute>
				<xsl:attribute name="COUNTY"><xsl:value-of select="substring($Address/COUNTY, 1, 20)"/></xsl:attribute>
				<xsl:attribute name="POSTCODE"><xsl:value-of select="substring($Address/POSTCODE, 1, 8)"/></xsl:attribute>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<xsl:template name="WriteMatchingAddressMapLines">
		<xsl:param name="AddressID"/>
		<xsl:param name="CustomerIndex"/>
		<xsl:param name="AddressIndex"/>
		<xsl:for-each select="$Omiga/CUSTOMER/ADDRESS[@DUPLICATE=$AddressID]">
			<xsl:call-template name="WriteAddressMapLine">
				<xsl:with-param name="Customer" select="ancestor::*"/>
				<xsl:with-param name="Address" select="."/>
				<xsl:with-param name="CustomerIndex" select="$CustomerIndex"/>
				<xsl:with-param name="AddressIndex" select="$AddressIndex"/>
			</xsl:call-template>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="WriteAddressMapLine">
		<xsl:param name="Customer"/>
		<xsl:param name="Address"/>
		<xsl:param name="CustomerIndex"/>
		<xsl:param name="AddressIndex"/>
		<xsl:element name="ADDRESS">
			<xsl:attribute name="BLOCKSEQNUMBER"><xsl:choose><xsl:when test="$CustomerIndex = '1'"><xsl:choose><xsl:when test="$AddressIndex = 1"><xsl:value-of select="'01'"/></xsl:when><xsl:when test="$AddressIndex = 2"><xsl:value-of select="'11'"/></xsl:when><xsl:when test="$AddressIndex = 3"><xsl:value-of select="'21'"/></xsl:when></xsl:choose></xsl:when><xsl:when test="$CustomerIndex = '2'"><xsl:choose><xsl:when test="$AddressIndex = 1"><xsl:value-of select="'02'"/></xsl:when><xsl:when test="$AddressIndex = 2"><xsl:value-of select="'12'"/></xsl:when><xsl:when test="$AddressIndex = 3"><xsl:value-of select="'22'"/></xsl:when></xsl:choose></xsl:when></xsl:choose></xsl:attribute>
			<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="$Customer/@CUSTOMERNUMBER"/></xsl:attribute>
			<xsl:attribute name="CUSTOMERVERSIONNUMBER"><xsl:value-of select="$Customer/CUSTOMERVERSION/@CUSTOMERVERSIONNUMBER"/></xsl:attribute>
			<xsl:attribute name="TARGETADDRESSTYPE"><xsl:choose><xsl:when test="$AddressIndex = 1"><xsl:value-of select="'TH'"/></xsl:when><xsl:when test="$AddressIndex = 2"><xsl:value-of select="'TP'"/></xsl:when><xsl:when test="$AddressIndex = 3"><xsl:value-of select="'TP1'"/></xsl:when></xsl:choose></xsl:attribute>
			<xsl:attribute name="DATEMOVEDIN"><xsl:value-of select="translate($Address/@DATEMOVEDIN, '/', '')"/></xsl:attribute>
			<xsl:attribute name="DATEMOVEDOUT"><xsl:value-of select="translate($Address/@DATEMOVEDOUT, '/', '')"/></xsl:attribute>
			<xsl:attribute name="NATUREOFOCCUPANCY"><xsl:value-of select="$Address/@NATUREOFOCCUPANCY"/></xsl:attribute>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ADDR-UKBUREA">
		<xsl:variable name="Chk1respC" select="$CCBUKResponse/CHK1[@seq=./@seq]/child::node()='C'"/>
		<xsl:variable name="Chk1respR" select="$CCBUKResponse/CHK1[@seq=./@seq]/child::node()='R'"/>
		<xsl:variable name="Chk1node" select="$CCBUKResponse/CHK1[@seq=./@seq]"/>
		<xsl:element name="ADDRESSTARGET">
			<xsl:attribute name="ID"><xsl:value-of select="position()"/></xsl:attribute>
			<xsl:attribute name="BLOCKTYPE">BUK1</xsl:attribute>
			<xsl:attribute name="BLOCKSEQNUMBER"><xsl:value-of select="./@seq"/></xsl:attribute>
			<xsl:attribute name="RMCCODE"><xsl:value-of select="./RMC"/></xsl:attribute>
			<xsl:attribute name="DETAILCODE">8</xsl:attribute>
			<xsl:attribute name="REGIONCODE"><xsl:value-of select="./REGIONNUMBER"/></xsl:attribute>
			<xsl:attribute name="ADDRESSKEY"><xsl:value-of select="./ADDRESSKEY"/></xsl:attribute>
			<xsl:attribute name="FLATENHANCED"><xsl:value-of select="$Chk1node/FLAT"/></xsl:attribute>
			<xsl:attribute name="HOUSENAMEENHANCED"><xsl:value-of select="$Chk1node/HOUSENAME"/></xsl:attribute>
			<xsl:attribute name="HOUSENUMBERENHANCED"><xsl:value-of select="$Chk1node/HOUSENUMBER"/></xsl:attribute>
			<xsl:attribute name="STREETENHANCED"><xsl:value-of select="$Chk1node/STREET"/></xsl:attribute>
			<xsl:attribute name="DISTRICTENHANCED"><xsl:value-of select="$Chk1node/DISTRICT"/></xsl:attribute>
			<xsl:attribute name="TOWNENHANCED"><xsl:value-of select="$Chk1node/TOWN"/></xsl:attribute>
			<xsl:attribute name="COUNTYENHANCED"><xsl:value-of select="$Chk1node/COUNTY"/></xsl:attribute>
			<xsl:attribute name="POSTCODEENHANCED"><xsl:value-of select="$Chk1node/POSTCODE"/></xsl:attribute>
			<xsl:attribute name="FLAT"><xsl:value-of select="user:toPC(substring(normalize-space(./FLAT), 1, 16))"/></xsl:attribute>
			<xsl:attribute name="HOUSENAME"><xsl:value-of select="user:toPC(substring(normalize-space(./HOUSENAME), 1, 26))"/></xsl:attribute>
			<xsl:attribute name="HOUSENUMBER"><xsl:value-of select="substring(./HOUSENUMBER, 1, 10)"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="user:toPC(substring(normalize-space(./STREET), 1, 40))"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="user:toPC(substring(normalize-space(./DISTRICT), 1, 30))"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="user:toPC(substring(normalize-space(./TOWN), 1, 20))"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="user:toPC(substring(normalize-space(./COUNTY), 1, 20))"/></xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="substring(./POSTCODE, 1, 8)"/></xsl:attribute>
			<xsl:choose>
				<xsl:when test="$Chk1respC or $Chk1respR">
					<xsl:attribute name="ADDRESSRESOLVED">N</xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="ADDRESSRESOLVED">Y</xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:element>
	</xsl:template>
	<xsl:template match="ADDR-UK">
		<xsl:element name="ADDRESSTARGET">
			<xsl:attribute name="ID"><xsl:value-of select="position()"/></xsl:attribute>
			<xsl:attribute name="BLOCKTYPE">AUK1</xsl:attribute>
			<xsl:attribute name="BLOCKSEQNUMBER"><xsl:value-of select="./@seq"/></xsl:attribute>
			<xsl:attribute name="DETAILCODE">1</xsl:attribute>
			<xsl:attribute name="RMCCODE"><xsl:value-of select="./RMC"/></xsl:attribute>
			<xsl:attribute name="REGIONCODE"><xsl:value-of select="./REGIONNUMBER"/></xsl:attribute>
			<xsl:attribute name="FLAT"><xsl:value-of select="user:toPC(substring(normalize-space(./FLAT), 1, 16))"/></xsl:attribute>
			<xsl:attribute name="HOUSENAME"><xsl:value-of select="user:toPC(substring(normalize-space(./HOUSENAME), 1, 26))"/></xsl:attribute>
			<xsl:attribute name="HOUSENUMBER"><xsl:value-of select="substring(./HOUSENUMBER, 1, 10)"/></xsl:attribute>
			<xsl:attribute name="STREET"><xsl:value-of select="user:toPC(substring(normalize-space(./STREET), 1, 40))"/></xsl:attribute>
			<xsl:attribute name="DISTRICT"><xsl:value-of select="user:toPC(substring(normalize-space(./DISTRICT), 1, 30))"/></xsl:attribute>
			<xsl:attribute name="TOWN"><xsl:value-of select="user:toPC(substring(normalize-space(./TOWN), 1, 20))"/></xsl:attribute>
			<xsl:attribute name="COUNTY"><xsl:value-of select="user:toPC(substring(normalize-space(./COUNTY), 1, 20))"/></xsl:attribute>
			<xsl:attribute name="POSTCODE"><xsl:value-of select="substring(./POSTCODE, 1, 8)"/></xsl:attribute>
			<xsl:attribute name="ADDRESSRESOLVED">N</xsl:attribute>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
