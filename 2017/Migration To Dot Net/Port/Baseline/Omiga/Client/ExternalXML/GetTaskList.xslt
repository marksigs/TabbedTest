<?xml version="1.0" encoding="utf-8"?>
<!-- IK EP2_2227 add TASK.REMOTEOWNERTASKIND -->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="xml" indent="yes"/>
	<xsl:template match="/">
		<xsl:element name="TASK">
			<xsl:for-each select="./TASK">
				<xsl:attribute name="TASKID"><xsl:value-of select="@TASKID"/></xsl:attribute>
				<xsl:attribute name="APPLICANT"><xsl:value-of select="@APPLICANT"/></xsl:attribute>
				<xsl:attribute name="AUTOMATICTASKIND"><xsl:value-of select="@AUTOMATICTASKIND"/></xsl:attribute>
				<xsl:attribute name="CHASINGPERIODDAYS"><xsl:value-of select="@CHASINGPERIODDAYS"/></xsl:attribute>
				<xsl:attribute name="CHASINGPERIODHOURS"><xsl:value-of select="@CHASINGPERIODHOURS"/></xsl:attribute>
				<xsl:attribute name="CHASINGPERIODMINUTES"><xsl:value-of select="@CHASINGPERIODMINUTES"/></xsl:attribute>
				<xsl:attribute name="CONTACTCUSTOMERIND"><xsl:value-of select="@CONTACTCUSTOMERIND"/></xsl:attribute>
				<xsl:attribute name="CUSTOMERTASK"><xsl:value-of select="@CUSTOMERTASK"/></xsl:attribute>
				<xsl:attribute name="NOTAPPLICABLEFLAG"><xsl:value-of select="@NOTAPPLICABLEFLAG"/></xsl:attribute>
				<xsl:attribute name="TASKNAME"><xsl:value-of select="@TASKNAME"/></xsl:attribute>
				<xsl:attribute name="TASKOWNERTYPE"><xsl:value-of select="@TASKOWNERTYPE"/></xsl:attribute>
				<xsl:attribute name="TASKTYPE"><xsl:value-of select="@TASKTYPE"/></xsl:attribute>
				<xsl:attribute name="REMOTEOWNERTASKIND"><xsl:value-of select="@REMOTEOWNERTASKIND"/></xsl:attribute>
				<xsl:attribute name="ALWAYSAUTOMATICONCREATION"><xsl:value-of select="@ALWAYSAUTOMATICONCREATION"/></xsl:attribute>
				<xsl:choose>
					<xsl:when test="TASKP">
						<xsl:for-each select="TASKP">
							<xsl:attribute name="CASEPRIORITY"><xsl:value-of select="@CASEPRIORITY"/></xsl:attribute>
							<xsl:attribute name="CASEPRIORITYTEXT"><xsl:value-of select="@CASEPRIORITYTEXT"/></xsl:attribute>
							<xsl:attribute name="ADJUSTMENTDAYS"><xsl:value-of select="@ADJUSTMENTDAYS"/></xsl:attribute>
							<xsl:attribute name="ADJUSTMENTHOURS"><xsl:value-of select="@ADJUSTMENTHOURS"/></xsl:attribute>
							<xsl:attribute name="ADJUSTMENTMINUTES"><xsl:value-of select="@ADJUSTMENTMINUTES"/></xsl:attribute>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="CASEPRIORITY">0</xsl:attribute>
						<xsl:attribute name="CASEPRIORITYTEXT"/>
						<xsl:attribute name="ADJUSTMENTDAYS">0</xsl:attribute>
						<xsl:attribute name="ADJUSTMENTHOURS">0</xsl:attribute>
						<xsl:attribute name="ADJUSTMENTMINUTES">0</xsl:attribute>
					</xsl:otherwise>
				</xsl:choose>
				<xsl:choose>
					<xsl:when test="STAGETASK">
						<xsl:for-each select="STAGETASK">
							<xsl:attribute name="TASKUSERID"><xsl:value-of select="@TASKUSERID"/></xsl:attribute>
							<xsl:attribute name="TASKUNITID"><xsl:value-of select="@TASKUNITID"/></xsl:attribute>
						</xsl:for-each>
					</xsl:when>
					<xsl:otherwise>
						<xsl:attribute name="TASKUSERID"/>
						<xsl:attribute name="TASKUNITID"/>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
