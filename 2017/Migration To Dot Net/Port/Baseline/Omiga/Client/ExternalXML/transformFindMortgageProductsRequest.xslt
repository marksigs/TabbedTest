<?xml version="1.0" encoding="UTF-8"?>
<!--SR 19/09/2006 EP2_1 modified namespaces for Epsom -->
<!--SR 19/09/2006 EP2_283 add default values if called from Web -->
<!--IK 15/01/2007 EP2_835 default value for FLEXIBLENONFLEXIBLEPRODUCTS -->
<!--IK 31/01/2007 EP2_1152 default USERAUTHORITYLEVEL on REQUEST -->
<!--IK 04/02/2007 EP2_882 default GUARANTORIND on REQUEST -->
<!--IK 13/02/2007 EP2_1086 default SEARCHCONTEXT= 'COST MODELLING' with a (space) -->
<!--IK 16/02/2007 EP2_1086 default ALLPRODUCTSWITHCHECKS -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" 
	   xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:msg="http://Request.FindMortgageProducts.Omiga.vertex.co.uk">
	<xsl:template match="/">
		<xsl:for-each select="msg:REQUEST">
			<xsl:element name="REQUEST">
				<xsl:for-each select="@*">
					<xsl:copy-of select="."/>
				</xsl:for-each>
				<!-- EP2_1152 -->
				<xsl:attribute name="USERAUTHORITYLEVEL">90</xsl:attribute>
				<xsl:for-each select="msg:MORTGAGEPRODUCTREQUEST">
					<xsl:element name="MORTGAGEPRODUCTREQUEST">										
						<xsl:if test="not(msg:DISCOUNTEDPRODUCTS)">
							<xsl:element name="DISCOUNTEDPRODUCTS">0</xsl:element>
						</xsl:if> 
		
						<xsl:if test="not(msg:FIXEDRATEPRODUCTS)">
							<xsl:element name="FIXEDRATEPRODUCTS">0</xsl:element>
						</xsl:if> 
	
						<xsl:if test="not(msg:STANDARDVARIABLEPRODUCTS)">
							<xsl:element name="STANDARDVARIABLEPRODUCTS">0</xsl:element>
						</xsl:if> 
	
						<xsl:if test="not(msg:CAPPEDFLOOREDPRODUCTS)">
							<xsl:element name="CAPPEDFLOOREDPRODUCTS">0</xsl:element>
						</xsl:if> 
	
						<xsl:if test="not(msg:CASHBACKPRODUCTS)">
							<xsl:element name="CASHBACKPRODUCTS">0</xsl:element>
						</xsl:if> 
						<!-- EP2_835 -->
						<xsl:element name="FLEXIBLENONFLEXIBLEPRODUCTS">1</xsl:element>

						<xsl:if test="not(msg:IMPAREDCREDITRATING)">
							<xsl:element name="IMPAREDCREDITRATING">0</xsl:element>
						</xsl:if>

						<xsl:if test="not(msg:PRODUCTSBYGROUP)">
							<xsl:element name="PRODUCTSBYGROUP">0</xsl:element>
						</xsl:if>
	
						<xsl:if test="not(msg:ALLPRODUCTSWITHCHECKS)">
							<xsl:element name="ALLPRODUCTSWITHCHECKS">1</xsl:element>
						</xsl:if>
						
						<xsl:if test="not(msg:SEARCHCONTEXT)">
							<xsl:element name="SEARCHCONTEXT">COST MODELLING</xsl:element> <!-- EP2_1086 -->
						</xsl:if>
						
						<xsl:if test="not(msg:MEMBEROFSTAFF)">
							<xsl:element name="MEMBEROFSTAFF">0</xsl:element>
						</xsl:if>

						<xsl:if test="not(msg:EXISTINGCUSTOMER)">
							<xsl:element name="EXISTINGCUSTOMER">0</xsl:element>
						</xsl:if>
						
						<!-- EP2_882 -->
						<xsl:if test="not(msg:GUARANTORIND)">
							<xsl:element name="GUARANTORIND">0</xsl:element>
						</xsl:if>
							
						<xsl:apply-templates/>							
					</xsl:element>
			</xsl:for-each>							
			</xsl:element>
		</xsl:for-each>
	</xsl:template>

	<xsl:template match="*">	
		<xsl:element name="{name()}">
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	
</xsl:stylesheet>
