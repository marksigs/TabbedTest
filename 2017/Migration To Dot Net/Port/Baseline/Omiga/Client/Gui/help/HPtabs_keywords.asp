<%@ Language=JavaScript%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="Help_Style.css">
<title></title>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--


function btnContents_onclick()
{
	//load the contents page into the main frame
	top.frames("main").document.location.href="HP010.asp"
	//change the way the tabs look
	top.frames("tabs").document.location.href="HPtabs_contents.asp"
}

function btnKeywords_onclick()
{
	//load the contents page into the main frame
	top.frames("main").document.location.href="HP020.asp"
	//change the way the tabs look
	top.frames("tabs").document.location.href="HPtabs_keywords.asp"
}

function btnHelp_onclick()
{
	//load the help page into the main frame
	var Help =  top.document.frames("context").txtPageName.value	
	top.frames("main").document.location.href = "Help_" + Help + ".asp";
	//change the way the tabs look
	top.frames("tabs").document.location.href="HPtabs_Help.asp"
}

//-->
</script>
</head>
<body LANGUAGE="javascript">
<P align=center>
<IMG height=30 id=btnContents onclick="return btnContents_onclick()" src="../images/btnContents.gif" width=150>
<IMG height=30 id=btnKeywords onclick="return btnKeywords_onclick()" src="../images/btnKeywords_dis.gif" width=150>
<IMG height=30 id=btnHelp onclick="return btnHelp_onclick()" src="../images/btnHelp.gif" width=150>
</P>

</body>
</html>
