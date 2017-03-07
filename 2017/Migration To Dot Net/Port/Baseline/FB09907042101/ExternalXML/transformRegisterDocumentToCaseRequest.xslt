<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msg="http://Request.RegisterDocumentToCase.IDUK.Omiga.marlboroughstirling.com">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="msg:REQUEST">
		<xsl:element name="REQUEST">
			<xsl:if test="msg:DOCUMENTDETAILS[@WORKGROUPID][not(@DOCUMENTGROUP)]">
				<xsl:attribute name="preProcRef">true</xsl:attribute>
				<xsl:attribute name="preProcProgId">omRegisterDocumentToCase.PreProcBO</xsl:attribute>
			</xsl:if>
			<xsl:for-each select="@*">
				<xsl:if test="name() = 'USERID'">
					<xsl:copy/>
				</xsl:if>
				<xsl:if test="name() = 'UNITID'">
					<xsl:copy/>
				</xsl:if>
				<xsl:if test="name() = 'CHANNELID'">
					<xsl:copy/>
				</xsl:if>
				<xsl:if test="name() = 'PASSWORD'">
					<xsl:copy/>
				</xsl:if>
				<xsl:if test="name() = 'USERAUTHORITYLEVEL'">
					<xsl:copy/>
				</xsl:if>
				<xsl:attribute name="postProcRef">true</xsl:attribute>
				<xsl:attribute name="postProcProgId">omRegisterDocumentToCase.PostProcBO</xsl:attribute>
			</xsl:for-each>
			<xsl:apply-templates/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="msg:DMSREGDOC"/>
	<xsl:template match="msg:DOCUMENTDETAILS">
		<xsl:element name="OPERATION">
			<xsl:attribute name="CRUD_OP">CREATE</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">DOCUMENTAUDITDETAILS</xsl:attribute>
			<xsl:element name="DOCUMENTAUDITDETAILS">
				<xsl:for-each select="@*">
					<xsl:choose>
						<xsl:when test="name() = 'UNITNAME'"/>
						<xsl:when test="name() = 'CREATETIMESTAMP'">
							<xsl:attribute name="CREATIONDATE"><xsl:value-of select="."/></xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:copy/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
				<xsl:attribute name="INBOUNDDOCUMENT">1</xsl:attribute>
			</xsl:element>
		</xsl:element>
		<xsl:element name="OPERATION">
			<xsl:attribute name="CRUD_OP">CREATE</xsl:attribute>
			<xsl:attribute name="ENTITY_REF">EVENTAUDITDETAIL</xsl:attribute>
			<xsl:element name="EVENTAUDITDETAIL">
				<xsl:for-each select="@*">
					<xsl:choose>
						<xsl:when test="name() = 'APPLICATIONNUMBER'">
							<xsl:copy/>
						</xsl:when>
						<xsl:when test="name() = 'DOCUMENTGUID'">
							<xsl:copy/>
							<xsl:attribute name="FILENETIMAGEREF"><xsl:value-of select="."/></xsl:attribute>
							<xsl:attribute name="FILEGUID">00000000000000000000000000000000</xsl:attribute>
						</xsl:when>
						<xsl:when test="name() = 'CREATETIMESTAMP'">
							<xsl:attribute name="EVENTDATE"><xsl:value-of select="."/></xsl:attribute>
						</xsl:when>
					</xsl:choose>
				</xsl:for-each>
				<xsl:for-each select="../@USERID">
					<xsl:copy/>
				</xsl:for-each>
				<xsl:for-each select="../@UNITID">
					<xsl:copy/>
				</xsl:for-each>
				<xsl:attribute name="FILEVERSION">1</xsl:attribute>
				<xsl:attribute name="FILEVERSIONNUM">1</xsl:attribute>
				<xsl:attribute name="EVENTKEY">0</xsl:attribute>
				<xsl:attribute name="NUMBEROFCOPIES">0</xsl:attribute>
				<xsl:attribute name="PRINTLOCATION">Scanned Image</xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
