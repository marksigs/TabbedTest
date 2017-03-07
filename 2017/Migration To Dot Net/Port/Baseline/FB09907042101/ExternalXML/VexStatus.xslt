<?xml version="1.0" encoding="UTF-8"?>
<!--
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Author		Date			Description
PE			10/11/2006	EP2_24 - Created / Xit2 Web Service
PE			07/12/2006	EP2_345 - Xit2 Interface - InstructionSystemID
====================================================================================
--> 
<!--=================================================================================================-->
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<!--=================================================================================================-->
	<xsl:template match="/">
		<xsl:apply-templates select="instruction"/>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template match="instruction">
		<Instruction xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<InstructionRef>
				<xsl:value-of select="/instruction/@applicationnumber"/>
				<xsl:value-of select="/instruction/@instructionref"/>
			</InstructionRef>
			<InstructionRef1/>
			<InstructionRef2/>
			<InstructionRef3/>
			<InstructionDX/>
			<SourceGUID>
				<xsl:value-of select="/instruction/@sourceguid"/>
			</SourceGUID>
			<WhenCreated>
				<xsl:value-of select="/instruction/@instructiondate"/>
			</WhenCreated>
			<ReadyToSend>1</ReadyToSend>
			<IsSent>0</IsSent>
			<WhenSent>
				<xsl:value-of select="/instruction/@instructiondate"/>
			</WhenSent>
			<xsl:call-template name="Status"/>
		</Instruction>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="Status">
		<Status>
			<InstructionSystemID>
				<xsl:value-of select="/instruction/@vexinstructionsystemid"/>
			</InstructionSystemID>
			<StatusCode>CANCEL</StatusCode>
			<StatusDateActioned>
				<xsl:value-of select="/instruction/@instructiondate"/>
			</StatusDateActioned>
			<xsl:call-template name="SourceFirm"/>
			<xsl:call-template name="SourceUser"/>
			<Reason/>
		</Status>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="SourceFirm">
		<SourceFirm>
			<SourceFirmName>
				<xsl:value-of select="/instruction/sourcefirm/@sourcefirmname"/>
			</SourceFirmName>
			<SourceFirmID>
				<xsl:value-of select="/instruction/sourcefirm/@sourcefirmid"/>
			</SourceFirmID>
			<SourceFirmAddress>
				<HouseName/>
				<HouseNumber/>
				<Street/>
				<District/>
				<Area/>
				<City/>
				<County/>
				<Country/>
				<PostCode/>
			</SourceFirmAddress>
		</SourceFirm>
	</xsl:template>
	<!--=================================================================================================-->
	<xsl:template name="SourceUser">
		<SourceUser>
			<SourceUserID></SourceUserID>
			<SourceUserName>
				<FullName></FullName>
				<Title/>
				<ForeName/>
				<MiddleNames/>
				<LastName/>
			</SourceUserName>
		</SourceUser>
	</xsl:template>
	<!--=================================================================================================-->
</xsl:stylesheet>
<!--=================================================================================================-->
