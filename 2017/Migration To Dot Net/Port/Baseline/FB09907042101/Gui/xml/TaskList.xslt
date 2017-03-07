<?xml version="1.0" encoding="UTF-8"?>
<!-- PB	13/09/2006	EP602	Created (after I. Kemp) for EP602 / CC47 -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="html" version="4" encoding="UTF-8" indent="yes"/>
	<!-- Define Styles-->
	<xsl:template match="/">
		<html>
			<head>
				<title/>
				<style>
				body {margin-left:20px;font-size: 10pt; font-family: 'trebuchet ms'}
				.head1{font-size:12pt;font-weight:bold}
				.prompt {width:10%;font-weight:bold}
				.boxHead {background-color: gainsboro;border-bottom: thin solid;font-weight:bold;padding-left:4px}
			</style>
			</head>
			<body>
				<xsl:apply-templates/>
			</body>
		</html>
	</xsl:template>
	
<!-- Main Routine-->
	<xsl:template match="REQUEST/TEMPLATEDATA">
		<xsl:for-each select="PRINTDOCSEARCH">
			<div class="head1" align="center">Task List Criteria</div>
			<br></br>
			<xsl:call-template name="criteraDets"/>
		</xsl:for-each>
		
		<hr></hr><br></br>
		<div class="head1" align="center">Due Tasks</div>
		
		<div class="boxHead">
			<span style="width:80px;">
				Application number
			</span>
			<span style="width:250px;">
				Task
			</span>
			<span style="width:80px;">
				Status
			</span>
			<span style="width:150px;">
				Date / Time
			</span>			
		</div>
		<xsl:for-each select="PRINTDOCRESULT">
			<div style="width:95%;margin-top:16px;">
				<xsl:call-template name="resultDets"/>
			</div>
		</xsl:for-each>
		<br></br>
	</xsl:template>
	<!-- Standard Page Header -->
	<xsl:template name="criteraDets">
		<div>
			<span class="prompt" style="width:150px;">Unit name.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHUNITNAME"/>
			</span>
			<span class="prompt" style="width:150px;">User name.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHUSERNAME"/>
			</span>
		</div>
		<div>
			<span class="prompt" style="width:150px;">Task type.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHTASKTYPE"/>
			</span>
			<span class="prompt" style="width:150px;">Application number.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHAPPNUMBER"/>
			</span>
		</div>
		<div>
			<span class="prompt" style="width:150px;">Due date start.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHDUEDATESTART"/>
			</span>
			<span class="prompt" style="width:150px;">Due date end.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHDUEDATEEND"/>
			</span>
		</div>
		<div>
			<span class="prompt" style="width:150px;">Sales lead status.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHSALESLEADSTATUS"/>
			</span>
			<span class="prompt" style="width:150px;">SLA Expiry within.:</span>
			<span style="width:150px;">
			    <xsl:value-of select="//REQUEST/TEMPLATEDATA/PRINTDOCSEARCH/@SEARCHSLAEXPIRYWITHIN"/>
			</span>
		</div>
	</xsl:template>
	<!-- Task Details-->
	<xsl:template name="resultDets">
		<div style="margin-top:4px">
			<span style="width:80px"><xsl:value-of select="@RESULTAPPNUMBER"/></span>
			<span style="width:250px;"><xsl:value-of select="@RESULTTASK"/></span>
			<span style="width:80px;"><xsl:value-of select="@RESULTSTATUS"/></span>
			<span style="width:150px;"><xsl:value-of select="@RESULTDATETIME"/></span>
		</div>
	</xsl:template>
</xsl:stylesheet>
