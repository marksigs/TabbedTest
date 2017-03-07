<%@ language=jscript %>
<%
	/*
	
	Version history
	---------------------------------------------------------------------------------------------------
	Prg		Date			AQR			Description
	PB		13/09/2006		EP602		New page for CC47. Added 'Print without archive' facility for task list.
	
	*/
%>
<html>
	<head>
		<title>
			Due Tasks
		</title>
		<script language="javascript">
			function startUp()
			{
				var strHTML=self.parent.opener.getHTMLToPrint();
				parent.fraPrintDocument.document.write(strHTML);
			}
			function printNow()
			{
				parent.fraPrintDocument.focus();
				window.print();
			}
		</script>
	</head>
	<body stype='background-color:#f0f0f0;' onload='startUp();'>
		<div id='menuBar' style='WIDTH:100%;HEIGHT:50px;BACKGROUND-COLOR:#f0f0f0'>
			<input type="button" style='LEFT:10px;POSITION:relative;TOP:10px;' onclick='printNow()' value='Print' ID="Button1" NAME="Button1"/>
			<input type="button" style='LEFT:10px;POSITION:relative;top:10px;' onclick='window.parent.close()' value='Close' ID="Button2" NAME="Button2"/>
		</div>
	</body>
</html>