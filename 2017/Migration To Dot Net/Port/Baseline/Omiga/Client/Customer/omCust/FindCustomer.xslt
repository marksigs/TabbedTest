<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/">
		<xsl:element name="CUSTOMERLIST">
			<xsl:for-each select="CUSTOMERLIST/CUSTOMER">
			<xsl:sort order="ascending" select=".//SURNAME[1]"/>
			<xsl:sort order="ascending" select=".//FIRSTFORENAME[1]"/>
			<xsl:sort order="ascending" select=".//SECONDFORENAME[1]"/>
			<xsl:sort order="ascending" select=".//OTHERFORENAMES[1]"/>	
			<xsl:sort order="ascending" select="CUSTOMERNUMBER[1]"/>	
				<xsl:element name="CUSTOMER">
					<xsl:element name="CUSTOMERNUMBER">
						<xsl:choose>
							<xsl:when test=".//CUSTOMERNUMBER[1]">
								<xsl:value-of select=".//CUSTOMERNUMBER[1]"/>
							</xsl:when>
							<xsl:when test=".//OMIGACUSTOMERNUMBER[1]">
								<xsl:value-of select=".//OMIGACUSTOMERNUMBER[1]"/>
							</xsl:when>
							<xsl:otherwise/>
						</xsl:choose>
					</xsl:element>	
					<xsl:element name="CUSTOMERVERSIONNUMBER">
						<xsl:value-of select=".//CUSTOMERVERSIONNUMBER[1]"/>
					</xsl:element>			
					<xsl:element name="OTHERSYSTEMCUSTOMERNUMBER">
						<xsl:value-of select=".//OTHERSYSTEMCUSTOMERNUMBER[1]"/>
					</xsl:element>			
					<xsl:element name="SURNAME">
						<xsl:value-of select=".//SURNAME[1]"/>
					</xsl:element>			
					<xsl:element name="FIRSTFORENAME">
						<xsl:value-of select=".//FIRSTFORENAME[1]"/>
					</xsl:element>			
					<xsl:element name="SECONDFORENAME">
						<xsl:value-of select=".//SECONDFORENAME[1]"/>
					</xsl:element>			
					<xsl:element name="OTHERFORENAMES">
						<xsl:value-of select=".//OTHERFORENAMES[1]"/>
					</xsl:element>			
					<xsl:element name="CORRESPONDENCESALUTATION">
						<xsl:value-of select=".//CORRESPONDENCESALUTATION[1]"/>
					</xsl:element>			
					<xsl:element name="DATEOFBIRTH">
						<xsl:value-of select=".//DATEOFBIRTH[1]"/>
					</xsl:element>			
					<xsl:element name="BUILDINGORHOUSENAME">
						<xsl:value-of select=".//BUILDINGORHOUSENAME[1]"/>
					</xsl:element>			
					<xsl:element name="BUILDINGORHOUSENUMBER">
						<xsl:value-of select=".//BUILDINGORHOUSENUMBER[1]"/>
					</xsl:element>			
					<xsl:element name="FLATNUMBER">
						<xsl:value-of select=".//FLATNUMBER[1]"/>
					</xsl:element>			
					<xsl:element name="STREET">
						<xsl:value-of select=".//STREET[1]"/>
					</xsl:element>			
					<xsl:element name="DISTRICT">
						<xsl:value-of select=".//DISTRICT[1]"/>
					</xsl:element>			
					<xsl:element name="TOWN">
						<xsl:value-of select=".//TOWN[1]"/>
					</xsl:element>			
					<xsl:element name="COUNTY">
						<xsl:value-of select=".//COUNTY[1]"/>
					</xsl:element>			
					<xsl:element name="COUNTRY">
						<xsl:value-of select=".//COUNTRY[1]"/>
					</xsl:element>			
					<xsl:element name="POSTCODE">
						<xsl:value-of select=".//POSTCODE[1]"/>
					</xsl:element>			
					<xsl:element name="ADDRESSTYPE">
						<xsl:value-of select=".//ADDRESSTYPE[1]"/>
					</xsl:element>			
					<xsl:element name="CHANNELID">
						<xsl:value-of select=".//CHANNELID[1]"/>
					</xsl:element>			
					<xsl:element name="OTHERSYSTEMTYPE">
						<xsl:value-of select=".//OTHERSYSTEMTYPE[1]"/>
					</xsl:element>			
					<xsl:element name="MOTHERSMAIDENNAME">
						<xsl:value-of select=".//MOTHERSMAIDENNAME[1]"/>
					</xsl:element>			
					<xsl:element name="CUSTOMERSTATUS">
						<xsl:value-of select=".//CUSTOMERSTATUS[1]"/>
					</xsl:element>			
					<xsl:element name="SOURCE">
						<xsl:value-of select=".//SOURCE[1]"/>
					</xsl:element>			
			</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
