<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PB		12/07/2006	EP543			Modified to use 'Other' title if selected
	PE		16/08/2006	EP103			Reworked. Abstracted javascript functions
	PB		19/01/2007	EP2_804		Modified as per new spec
	DS        31/01/2007   EP2_804       Modified to fix the defect to get Landlord Address,changed the template structure as per new structure
	PB		23/02/2007	EP2_804		Modified to only print a letter for the customer id passed in
	OS		19/03/2007	EP2_2007	Modified to print a letter only if Tenancy record is present
	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<!--============================================================================================================-->
		
		<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
		<xsl:for-each select="/RESPONSE/CURRENTLANDLORDREFERENCE/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMERVERSION[@CUSTOMERNUMBER=/RESPONSE/CURRENTLANDLORDREFERENCE/@CUSTOMERNUMBER and TENANCY]">
				<xsl:element name="INFORMATION">
					<xsl:apply-templates select="current()"/>
				
					<xsl:if test="position() > 1">
						<xsl:element name="PAGEBREAK"/>
					</xsl:if>
				</xsl:element>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
	
	<!--============================================================================================================-->

</xsl:stylesheet>
