<?xml version="1.0" encoding="UTF-8"?>
<!-- MAR1166 GHun -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="*">
		<xsl:copy>
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>

	<!-- Elements to remove from the response -->
	<xsl:template match="NETCONFIRMEDALLOWABLEINCOME"/>
	<xsl:template match="MAXIMUMBORROWING/CONFIRMEDINCOME1"/>
	<xsl:template match="MAXIMUMBORROWING/CONFIRMEDINCOME2"/>
	<xsl:template match="MAXIMUMBORROWING/CONFIRMEDINCOMEMULTIPLIERTYPE"/>
	<xsl:template match="MAXIMUMBORROWING/CONFIRMEDINCOMEMULTIPLE"/>
	<xsl:template match="MAXIMUMBORROWING/CONFIRMEDMAXIMUMBORROWINGAMOUNT"/>

</xsl:stylesheet>
