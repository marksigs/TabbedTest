<%@ Language=JavaScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="Help_Style.css">
</HEAD>
<BODY>
<DIV> 
	<SPAN style="LEFT: 10px; POSITION: absolute; TOP: 30px; WIDTH: 280px">
		<p>Type in the keyword to find </p>
	</SPAN>			
	<INPUT class=msgTxt id=txtFind maxLength=10 style="LEFT: 160px; POSITION: absolute; TOP: 30px; WIDTH: 250px"> 
	<DIV style="LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 520px" ><!-- <div id="HP040" style="HEIGHT: 19px; LEFT:0px; POSITION: absolute; TOP: 40px; WIDTH:625px"> --><!--#include FILE="HP040.htm" --> 
	</DIV>	
	<DIV style = "LEFT: 430px; POSITION: absolute; TOP: 330px">
		<INPUT class=msgButton id=btnDisplay onclick="DisplayRecords()" style="WIDTH: 80px" type=button value=" DISPLAY">
	</DIV>
</DIV>
<script language="JScript">
<!--

//function ShowHelpPage(this.options[this.selectedIndex].value)
function ShowHelpPage(sPage)
{
	top.frames("main").document.location.href=sPage;
}

function DisplayRecords(sPage)
{
	ShowHelpPage(cboHelpType.options[cboHelpType.selectedIndex].value)
}

function txtFind.onkeyup()
{
	var re = new RegExp("^" + window.txtFind.value,"i");
	var icount, strValue, match; 

	for (icount = 0; icount < cboHelpType.options.length ; icount++)
	{
		strValue = 	cboHelpType.options.item(icount).text;
		match = strValue.match(re);
		if (match!= null)
		{
		cboHelpType.selectedIndex=icount;
		return;
		}
	}	
}

-->
</script>
</BODY>
</html>