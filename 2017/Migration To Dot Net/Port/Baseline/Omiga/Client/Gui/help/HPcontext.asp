<%@ Language=JavaScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function window_onload() 
{
var strFromOpener, strtitle, objOpener, strScreen;
objOpener=null;
objOpener=top.window.parent.opener;
	if (objOpener==null)
	{
		alert("No opener exists");
	}
	else
	{
		//LOAD CONTEXT FROM OPENER
		strScreen=objOpener.lblFW030ID.innerHTML;   // fw030.asp
		strSelected=objOpener.txtSelected.value;   //txtSelected
		txtChoice.value = strSelected;
		txtPageName.value = strScreen; 
	}
	switch(txtChoice.value)
		{
		case "1":
			//load the contents page into the main frame
			top.frames("main").document.location.href="HP010.asp"
			//change the way the tabs look
			top.frames("tabs").document.location.href="HPtabs_contents.asp";
			break;
		case "2":
			//load the contents page into the main frame
			top.frames("main").document.location.href="HP020.asp";
			//change the way the tabs look
			top.frames("tabs").document.location.href="HPtabs_keywords.asp";
			break;
		case "3":
			//load the contents page into the main frame
			try
			{
			top.frames("main").document.location.href= "Help_"+ txtPageName.value + ".asp";	
			}
			catch(e)
			{
			alert(e)
			}
			
			//change the way the tabs look
			top.frames("tabs").document.location.href="HPtabs_Help.asp";
			break;
		default:	
			//load the contents page into the main frame
			top.frames("main").document.location.href="HP010.asp"
			//change the way the tabs look
			top.frames("tabs").document.location.href="HPtabs_contents.asp"
			break;
		}
}	

function window_onfocus()
{
}

//-->
</script>

<LINK rel="stylesheet" type="text/css" href="Help_Style.css">

</head>
<body LANGUAGE="javascript" onload="return window_onload()" onfocus="return window_onfocus()">

<p>HPcontext
Choice<input id="txtChoice" name="txtChoice">
PageName<input id="txtPageName" name="txtPageName">
</p>

</body>
</html>
