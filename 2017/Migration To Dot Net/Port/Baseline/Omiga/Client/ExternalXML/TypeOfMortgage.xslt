<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<!--
		This xml is used to create nodes depending on the type of mortgage, for example, if transfer of equity (TOE) the nodes ABOTOEP and MAINADVORTOE will be created.
		This is done to allow mortgage-type specific text in letters.
		All which apply should be shown.
		
		MAINADVORTOE		Main advance OR transfer of equity
		ABOTOEP					Additional borrowing OR transfer of equity OR port
		ADDBRWGORPSW1	Additional borrowing OR product switch #1
		ADDBRWGORPSW2	As above , #2
		CREDLIMINC				Credit limit increase
		
	-->
	<!--OS	09/03/2007	EP2_734	Fixed TOE, ABTOE and CLI cases -->
	
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
		<xsl:template match="COMBOVALIDATION">
		<!-- Check if different mortgagetype or if additional text required -->

			<xsl:choose>
				<xsl:when test="@VALIDATIONTYPE = 'FT'">
					<xsl:element name="MAINADVORTOE"/>
					<xsl:element name="ABOTOEP"/>
				</xsl:when>
				<xsl:when test="@VALIDATIONTYPE = 'HM'">
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>
				<xsl:when test="@VALIDATIONTYPE = 'R'">
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>
				<!--RTB condition to be removed as per EP2_1512-->
				<!--<xsl:when test="@VALIDATIONTYPE = 'RTB'">
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>-->
				<!--Replicating logic from document-functions.xsl, for ABTOE cases-->
				<xsl:when test="@VALIDATIONTYPE = 'ABO'">
					<xsl:choose>
						<xsl:when test="//COMBOVALIDATION[@VALIDATIONTYPE='ABTOE']">
							<xsl:attribute name="MORTGAGETYPE">Additional Borrowing with Transfer of Equity</xsl:attribute>
							<xsl:element name="ABOTOEP"/>
							<xsl:element name="MAINADVORTOE"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="MORTGAGETYPE">Additional Borrowing</xsl:attribute>
							<xsl:element name="ABOTOEP"/>
							<xsl:element name="ADDBRWGORPSW1"/>
							<xsl:element name="ADDBRWGORPSW2"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="@VALIDATIONTYPE = 'TOE'">
					<xsl:attribute name="MORTGAGETYPE">Transfer of Equity</xsl:attribute>
					<xsl:element name="ABOTOEP"/>
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>
				<!--<xsl:when test="@VALIDATIONTYPE = 'ABTOE'">
					<xsl:attribute name="MORTGAGETYPE">Additional Borrowing with Transfer of Equity</xsl:attribute>
					<xsl:element name="ABOTOEP"/>
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>-->
				<xsl:when test="@VALIDATIONTYPE = 'NP'">
					<xsl:attribute name="MORTGAGETYPE">Porting</xsl:attribute>
					<xsl:element name="ABOTOEP"/>
					<xsl:element name="MAINADVORTOE"/>
				</xsl:when>
				<xsl:when test="@VALIDATIONTYPE = 'PSW'">
					<xsl:attribute name="MORTGAGETYPE">Product Switch</xsl:attribute>
					<xsl:element name="ADDBRWGORPSW1"/>
					<xsl:element name="ADDBRWGORPSW2"/>
				</xsl:when>
				<xsl:when test="@VALIDATIONTYPE = 'CLI'">
					<xsl:attribute name="MORTGAGETYPE">Credit Limit Increase</xsl:attribute>
					<xsl:element name="CREDLIMINC">
						<xsl:attribute name="INCREASEAMT"><xsl:value-of select="//MORTGAGESUBQUOTE/@AMOUNTREQUESTED"/></xsl:attribute>
					</xsl:element>
				</xsl:when>
			</xsl:choose>
			
	</xsl:template>
</xsl:stylesheet>
