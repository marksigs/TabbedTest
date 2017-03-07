<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ?????.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   ?????
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
BG		22/01/01	SYS1868 - Created screen
MC		21/04/2004	BMIDS517	White space padded to the title text

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/


%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">

<title>Flexible Mortgage Calculator  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<%/* Div containing all the serious stuff to show flexicalc*/%>
<div id="divBackground" style="TOP: 5px; LEFT: 10px; HEIGHT: 660px; WIDTH: 440px; POSITION: ABSOLUTE" bgColor="white">

	<object classid="clsid:166B1BCA-3F9C-11CF-8075-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=8,0,0,0" ID="FlexibleMortgageCalculator" width="440" height="620" VIEWASTEXT>
	<param name="src" value="FlexibleMortgageCalculator.dcr">
	<param name="swRemote" value="swSaveEnabled='true' swVolume='true' swRestart='true' swPausePlay='true' swFastForward='true' swContextMenu='true' ">
	<param name="swStretchStyle" value="none">
	<param NAME="bgColor" VALUE="#FFFFFF"> 
	<embed src="FlexibleMortgageCalculator.dcr" bgColor="#FFFFFF" swRemote="swSaveEnabled='true' swVolume='true' swRestart='true' swPausePlay='true' swFastForward='true' swContextMenu='true' " swStretchStyle="none" type="application/x-director" pluginspage="http://www.macromedia.com/shockwave/download/">
	</object>

	
</div>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 25px; POSITION: absolute; TOP: 625px; WIDTH: 400px"> 
	<!-- #include FILE="msgButtons.asp" -->
</div> 


<% /* Specify Code Here */ %>
<script language="JScript">
<!--

function window.onload()
{	
	var sArguments = window.dialogArguments;
		window.dialogTop	= sArguments[0];
		window.dialogLeft	= sArguments[1];
		window.dialogWidth	= sArguments[2];
		window.dialogHeight = sArguments[3];
			
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
}

function btnSubmit.onclick()
{	
	window.close();
}
-->
</script>
</body>
</html>

