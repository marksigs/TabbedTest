<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/">
		<xsl:element name="LCPSORTED">
			<xsl:for-each select="LOANCOMPONENTPAYMENT">
			<xsl:sort order="descending" select="@LOANCOMPONENTSEQUENCENUMBER"/>
				<xsl:element name="LOANCOMPONENTPAYMENT">
					<xsl:attribute name="APPLICATIONNUMBER">
						<xsl:value-of select ="@APPLICATIONNUMBER"/>
					</xsl:attribute>
					<xsl:attribute name="LOANCOMPONENTSEQUENCENUMBER">
						<xsl:value-of select ="@LOANCOMPONENTSEQUENCENUMBER"/>
					</xsl:attribute>
					<xsl:attribute name="MORTGAGESUBQUOTENUMBER">
						<xsl:value-of select ="@MORTGAGESUBQUOTENUMBER"/>
					</xsl:attribute>
					<xsl:attribute name="AMOUNT">
						<xsl:value-of select ="@AMOUNT"/>
					</xsl:attribute>
					<xsl:attribute name="INTERESTONLYELEMENT">
						<xsl:value-of select ="@INTERESTONLYELEMENT"/>
					</xsl:attribute>
					<xsl:attribute name="CAPITALANDINTERESTELEMENT">
						<xsl:value-of select ="@CAPITALANDINTERESTELEMENT"/>
					</xsl:attribute>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>

