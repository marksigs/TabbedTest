<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\omiga4Projects\projectMars\xml\ValidateQuoteData\ValidateQuoteDataResponse.xml?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="http://mycompany.com/mynamespace">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<msxsl:script language="JScript" implements-prefix="user"><![CDATA[
			function getExpiryDate(sPerMos)
			{
				if(isNaN(parseInt(sPerMos))) return("cannot compute");
				var dt = new Date();
				dt.setMonth(dt.getMonth() + parseInt(sPerMos));
				var day = (dt.getDate() +100).toString().substr(1);
				var month = (dt.getMonth() +101).toString().substr(1);
				return(day + "/" + month + "/" + dt.getFullYear());
			}
		]]></msxsl:script>
	<xsl:template match="RESPONSE[@TYPE='SUCCESS']">
		<xsl:element name="RESPONSE">
			<xsl:attribute name="TYPE">SUCCESS</xsl:attribute>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="APPLICATION">
		<xsl:element name="APPLICATION">
			<xsl:for-each select="@APPLICATIONNUMBER">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@ACCEPTEDQUOTENUMBER">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:for-each select="@ACTIVEQUOTENUMBER">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:attribute name="RATEDIFFIND">false</xsl:attribute>
			<xsl:if test="@RATEDIFFIND='1'">
				<xsl:attribute name="RATEDIFFIND">true</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="QUOTATION">
		<xsl:element name="QUOTATION">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="MORTGAGESUBQUOTE">
		<xsl:element name="MORTGAGESUBQUOTE">
			<xsl:for-each select="@*">
				<xsl:copy/>
			</xsl:for-each>
			<xsl:attribute name="VALFEEDIFFIND">false</xsl:attribute>
			<xsl:attribute name="PRODUCTRESERVEDIND">false</xsl:attribute>
			<xsl:attribute name="PRODUCTWITHDRAWNIND">false</xsl:attribute>
			<xsl:if test="VALUATIONFEE[@AMOUNT]">
				<xsl:if test="MORTGAGEONEOFFCOST[@AMOUNT]">
					<xsl:variable name="FeeAmnt" select="VALUATIONFEE/@AMOUNT"/>
					<xsl:variable name="CostAmnt" select="MORTGAGEONEOFFCOST/@AMOUNT"/>
					<xsl:if test="$FeeAmnt != $CostAmnt">
						<xsl:attribute name="VALFEEDIFFIND">true</xsl:attribute>
					</xsl:if>
				</xsl:if>
			</xsl:if>
			<xsl:if test="LOANCOMPONENT[@RESERVEDPRODUCTSTARTDATE]">
				<xsl:attribute name="PRODUCTRESERVEDIND">true</xsl:attribute>
			</xsl:if>
			<xsl:if test="LOANCOMPONENT[@WITHDRAWNDATE]">
				<xsl:attribute name="PRODUCTWITHDRAWNIND">true</xsl:attribute>
			</xsl:if>
			<xsl:if test="not(LOANCOMPONENT[@WITHDRAWNDATE])">
				<xsl:for-each select="LOANCOMPONENT[not(@WITHDRAWNDATE)][@INTERESTRATEENDDATE]">
					<xsl:element name="LOANCOMPONENT">
						<xsl:for-each select="@LOANCOMPONENTSEQUENCENUMBER"><xsl:copy/></xsl:for-each>
						<xsl:for-each select="@MORTGAGEPRODUCTCODE">
							<xsl:copy/>
						</xsl:for-each>
						<xsl:attribute name="CURRENTRATEEXPIRYDATE">
							<xsl:value-of select="substring(@INTERESTRATEENDDATE,9,2)"/>/<xsl:value-of select="substring(@INTERESTRATEENDDATE,6,2)"/>/<xsl:value-of select="substring(@INTERESTRATEENDDATE,1,4)"/>
						</xsl:attribute>
					</xsl:element>
				</xsl:for-each>
				<xsl:for-each select="LOANCOMPONENT[not(@WITHDRAWNDATE)][@INTERESTRATEPERIOD]">
					<xsl:variable name="periodMonths">
						<xsl:value-of select="@INTERESTRATEPERIOD"/>
					</xsl:variable>
					<xsl:element name="LOANCOMPONENT">
						<xsl:for-each select="@LOANCOMPONENTSEQUENCENUMBER">
							<xsl:copy/>
						</xsl:for-each>
						<xsl:for-each select="@MORTGAGEPRODUCTCODE">
							<xsl:copy/>
						</xsl:for-each>
						<xsl:attribute name="CURRENTRATEEXPIRYDATE"><xsl:value-of select="user:getExpiryDate(string($periodMonths))"/></xsl:attribute>
					</xsl:element>
				</xsl:for-each>
			</xsl:if>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
