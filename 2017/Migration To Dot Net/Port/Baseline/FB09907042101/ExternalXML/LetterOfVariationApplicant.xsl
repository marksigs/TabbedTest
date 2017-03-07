<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">

<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	DS 		11/01/2007 	EP2_739			Created 
	DS 		19/01/2007 	EP2_739			Added OFFERISSUEDATE attribute logic
	DS		09/02/2007	EP2_739   	   Removed OFFERISSUEDDATE logic, and put that in document-functions-applicant.xls
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:import href="document-functions-applicant.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
   
	<!--============================================================================================================-->
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:call-template name="APPLICANTINFO">
				<xsl:with-param name="RESPONSE" select="/RESPONSE/LETTEROFVARIATIONAPPLICANT"/>
			</xsl:call-template>
		<!--<xsl:apply-templates select="/RESPONSE/LETTEROFVARIATIONAPPLICANT/DOCUMENTAUDITDETAILS[position()=1]" /> -->
		</xsl:element>
	</xsl:template>
		
	<!--<xsl:template match="DOCUMENTAUDITDETAILS">
		<xsl:element name="APPOFFER">
		 <xsl:attribute name="OFFERISSUEDDATE">
			  <xsl:value-of select="concat('{',msg:GetDate(string(@CREATIONDATE)),'}')" /> 
		  </xsl:attribute>
	  </xsl:element>
    </xsl:template> -->

	
	<!--============================================================================================================-->    	
	
</xsl:stylesheet>