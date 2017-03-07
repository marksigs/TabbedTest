<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
		<xsl:template match='/'>
			<xsl:element name="EXISTINGBUSINESSLIST">
				<xsl:for-each select="EXISTINGBUSINESS">
				<xsl:sort order="descending" select="DATECREATED"/>
				<xsl:sort order="ascending" select="APPLICATIONNUMBERORACCOUNTNUMBER"/>
					<xsl:element name="EXISTINGBUSINESS">
						<xsl:element name="BUSINESSTYPE">
							<xsl:value-of select="BUSINESSTYPE"/>
						</xsl:element>
						<xsl:element name="BUSINESSTYPEINDICATOR">
							<xsl:value-of select="BUSINESSTYPEINDICATOR"/>
						</xsl:element>
						<xsl:element name="PACKAGENUMBER">
							<xsl:value-of select="PACKAGENUMBER"/>
						</xsl:element>
						<xsl:element name="APPLICATIONNUMBERORACCOUNTNUMBER">
							<xsl:choose>
								<xsl:when test="APPLICATIONNUMBER">
									<xsl:value-of select="APPLICATIONNUMBER"/>
								</xsl:when><xsl:when test="ACCOUNTNUMBER">
									<xsl:value-of select="ACCOUNTNUMBER"/>
								</xsl:when>
							</xsl:choose>
						</xsl:element>
						<xsl:element name="APPLICATIONNUMBER">
							<xsl:value-of select="APPLICATIONNUMBER"/>
						</xsl:element>
						<xsl:element name="ACCOUNTNUMBER">
							<xsl:value-of select="ACCOUNTNUMBER"/>
						</xsl:element>
						<xsl:element name="APPLICATIONFACTFINDNUMBER">
							<xsl:value-of select="APPLICATIONFACTFINDNUMBER"/>
						</xsl:element>
						<xsl:element name="DATECREATED">
							<xsl:value-of select="DATECREATED"/>
						</xsl:element>
						<xsl:element name="AMOUNT">
							<xsl:choose>
								<xsl:when test="AMOUNT">
									<xsl:value-of select="AMOUNT"/>
								</xsl:when>
								<xsl:when test="OUTSTANDINGBALANCE">
									<xsl:value-of select="OUTSTANDINGBALANCE"/>
								</xsl:when>
							</xsl:choose>
						</xsl:element>
						<xsl:element name="STAGENUMBER">
							<xsl:value-of select="STAGENUMBER"/>
						</xsl:element>
						<xsl:element name="STATUS">
							<xsl:value-of select="STATUS"/>
						</xsl:element>
						<xsl:element name="CORRESPONDENCESALUTATION">
							<xsl:value-of select="CORRESPONDENCESALUTATION"/>
						</xsl:element>
						<xsl:element name="TYPEOFAPPLICATION">
							<xsl:attribute name="TEXT">
								<xsl:value-of select="TYPEOFAPPLICATION/@TEXT"/>
							</xsl:attribute>
							<xsl:value-of select="TYPEOFAPPLICATION"/>
						</xsl:element>
						<xsl:element name="NUMBEROFAPPLICANTS">
							<xsl:value-of select="NUMBEROFAPPLICANTS"/>
						</xsl:element>
						<xsl:element name="CUSTOMERROLETYPE">
							<xsl:attribute name="TEXT">
								<xsl:value-of select="CUSTOMERROLETYPE/@TEXT"/>
							</xsl:attribute>
							<xsl:value-of select="CUSTOMERROLETYPE"/>
						</xsl:element>
						<xsl:element name="DRAWDOWN">
							<xsl:value-of select="DRAWDOWN"/>
						</xsl:element>
						<xsl:element name="OVERPAYMENTS">
							<xsl:value-of select="OVERPAYMENTS"/>
						</xsl:element>
						<xsl:element name="LOANCLASSTYPE">
							<xsl:value-of select="LOANCLASSTYPE"/>
						</xsl:element>
						<xsl:element name="TYPE">
							<xsl:value-of select="TYPE"/>
						</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
		</xsl:template>
</xsl:stylesheet>
