<% /* Background
------------------------- */ %>
<table style="width:627px; background-color:#00006a; position:absolute; left:-6px; top:0px; z-index:-2; height: 50px" id="blueback">
<tr><td style="height:44px"></td></tr>
</table>

<% /* ------------------------- Menu Buttons ------------------------- */ %>
<img src="images/memopadr.gif" width=27 height=28 alt="MemoPad" border="0" style="position: absolute; left:5; top:1; z-index:-1" id="MemoUp" onMouseDown="imageSwap(memoDown, memoPadMenu)" onMouseOver="SetCursor(memoUp)" onMouseOut="RestoreCursor(memoUp)">
<img src="images/memopaddown.gif" width=27 height=28 alt="MemoPad" border="0" style="position: absolute; left:5; top:1; z-index:-2" id="MemoDown" onClick="imageRestore(memoDown, memoPadMenu)" onMouseOver="SetCursor(memoDown)" onMouseOut="RestoreCursor(memoDown)">
<img src="images/printmenur.gif" width=27 height=28 alt="Print Menu" border="0" style="position: absolute; left:35; top:1; z-index:-1" id="PrintUp" onMouseDown="noMenuImageSwap(PrintDown)" onMouseOver="SetCursor(PrintUp)" onMouseOut="RestoreCursor(PrintUp)">
<img src="images/printmenudown.gif" width=27 height=28 alt="Print Menu" border="0" style="position: absolute; left:35; top:1; z-index:-2" id="PrintDown" onMouseUp="noMenuImageRestore(PrintDown)" onMouseOver="SetCursor(printDown)" onMouseOut="RestoreCursor(printDown)">
<img src="images/underwritingr.gif" width=27 height=28 alt="Application Underwriting" border="0" style="position: absolute; left:65; top:1; z-index:-1" id="AppRevUp" onMouseDown="imageSwap(AppRevDown, AppRevMenu)" onMouseOver="SetCursor(AppRevUp)" onMouseOut="RestoreCursor(AppRevUp)">
<img src="images/underwritingdown.gif" width=27 height=28 alt="Application Underwriting" border="0" style="position: absolute; left:65; top:1; z-index:-2" id="AppRevDown" onClick="imageRestore(AppRevDown, AppRevMenu)" onMouseOver="SetCursor(AppRevDown)" onMouseOut="RestoreCursor(AppRevDown)">
<img src="images/storedqr.gif" width=27 height=28 alt="Stored Quotes" border="0" style="position: absolute; left:95; top:1; z-index:-1" id="QuoteUp" onMouseDown="noMenuImageSwap(QuoteDown)" onMouseOver="SetCursor(quoteUp)" onMouseOut="RestoreCursor(quoteUp)">
<img src="images/storedqdown.gif" width=27 height=28 alt="Stored Quotes" border="0" style="position: absolute; left:95; top:1; z-index:-2" id="QuoteDown" onMouseUp="noMenuImageRestore(QuoteDown)" onMouseOver="SetCursor(quotedown)" onMouseOut="RestoreCursor(quotedown)">
<img src="images/presbuttonr.gif" width=27 height=28 alt="Presentations" border="0" style="position: absolute; left:125; top:1; z-index:-1" id="PresUp" onMouseDown="imageSwap(PresDown, PresMenu)" onMouseOver="SetCursor(presUp)" onMouseOut="RestoreCursor(presUp)">
<img src="images/presbuttondown.gif" width=27 height=28 alt="Presentations" border="0" style="position: absolute; left:125; top:1; z-index:-2" id="PresDown" onClick="imageRestore(PresDown, PresMenu)" onMouseOver="SetCursor(presDown)" onMouseOut="RestoreCursor(presDown)">
<img src="images/helpr.gif" width=27 height=28 alt="Help Menu" border="0" style="position: absolute; left:155; top:1; z-index:-1" id="HelpUp" onMouseDown="imageSwap(HelpDown, HelpMenu)" onMouseOver="SetCursor(helpUp)" onMouseOut="RestoreCursor(helpUp)">
<img src="images/helpdown.gif" width=27 height=28 alt="Help Menu" border="0" style="position: absolute; left:155; top:1; z-index:-2" id="HelpDown" onClick="imageRestore(HelpDown, HelpMenu)" onMouseOver="SetCursor(helpDown)" onMouseOut="RestoreCursor(helpDown)">

<% /* ------------------------- Fields ------------------------- */ %>
<div class="msgMenu" style="position:absolute; z-index:-1; left:430; top:3; width: 191px; height: 13px"> 
<b>Advisor:<span id=mbAdvisor></span></b>
</div>

<div class="msgMenu" style="position:absolute; z-index:-1; left:535px; top:32px; width: 86px; height: 13px" id="screenId"> 
<b><span id=mbScreenId></span></b>
</div>

<div class="msgMenu" style="position:absolute; z-index:-1; left:9px; top:32px; width: 471px; height: 13px"> 
<b><span id=mbScreenTitle></span></b>
</div>

<div class="msgMenu" style="position:absolute; z-index:-1; left:390px; top:18px; width: 231px; height: 13px" id="caseNo"> 
<b><span id=mbCaseNoPrompt>Case No:</span><span id=mbCaseNo></span></b>
</div>

<% /* Memo Pad Menu */ %>
<div id="MemoPadMenu" style="position:absolute; top:30; left:5; visibility:hidden; z-index: 100"> 
	<table id="MemoPadMen" width="90" class="msgMenu2">
	<tr><td id="mnu1opt1" onmouseOver="menuHighlight(mnu1opt1)" onmouseOut="menuNormal(mnu1opt1)" onClick="imageRestore(memoDown, memoPadMenu)">Screen Notes</td></tr>
	<tr><td id="mnu1opt2" onmouseOver="menuHighlight(mnu1opt2)" onmouseOut="menuNormal(mnu1opt2)" onClick="imageRestore(memoDown, memoPadMenu)">Summary</td></tr>
	</table>
</div>
<% /* Application Review Menu */ %>
<div id="AppRevMenu" style="position:absolute; top:30; left:65; visibility:hidden; z-index: 100"> 
	<table id="AppRevMen" width="190" class="msgMenu2">
	<tr><td id="mnu2opt1" onmouseOver="menuHighlight(mnu2opt1)" onmouseOut="menuNormal(mnu2opt1)" onClick="imageRestore(AppRevDown, AppRevMenu)">Application Review</td></tr>
	<tr><td id="mnu2opt2" onmouseOver="menuHighlight(mnu2opt2)" onmouseOut="menuNormal(mnu2opt2)" onClick="imageRestore(AppRevDown, AppRevMenu)">Existing Mortgage Account Summary</td></tr>
	<tr><td id="mnu2opt3" onmouseOver="menuHighlight(mnu2opt3)" onmouseOut="menuNormal(mnu2opt3)" onClick="imageRestore(AppRevDown, AppRevMenu)">Completeness Check</td></tr>
	<tr><td id="mnu2opt4" onmouseOver="menuHighlight(mnu2opt4)" onmouseOut="menuNormal(mnu2opt4)" onClick="imageRestore(AppRevDown, AppRevMenu)">Risk Assessment</td></tr>
	<tr><td id="mnu2opt5" onmouseOver="menuHighlight(mnu2opt5)" onmouseOut="menuNormal(mnu2opt5)" onClick="imageRestore(AppRevDown, AppRevMenu)">Bureau Data</td></tr>
	</table>
</div>
<% /* Presentations Menu */ %>
<div id="PresMenu" style="position:absolute; top:30; left:125; visibility:hidden; z-index: 100"> 
	<table id="PresMen" width="160" class="msgMenu2">
	<tr><td id="mnu3opt1" onmouseOver="menuHighlight(mnu3opt1)" onmouseOut="menuNormal(mnu3opt1)" onClick="imageRestore(PresDown, PresMenu)">The Mortgage Lifecycle</td></tr>
	<tr><td id="mnu3opt2" valign="top">----------------------------------</td></tr>
	<tr><td id="mnu3opt4" onmouseOver="menuHighlight(mnu3opt4)" onmouseOut="menuNormal(mnu3opt4)" onClick="imageRestore(PresDown, PresMenu)">Mortgage Products</td></tr>
	<tr><td id="mnu3opt5" onmouseOver="menuHighlight(mnu3opt5)" onmouseOut="menuNormal(mnu3opt5)" onClick="imageRestore(PresDown, PresMenu)">Buildings & Contents Products</td></tr>
	</table>
</div>
<% /* Help Pad Menu */ %>
<div id="HelpMenu" style="position:absolute; top:30; left:155; visibility:hidden; z-index: 100"> 
	<table id="HelpMen" width="110" class="msgMenu2">
	<tr><td id="mnu4opt1" onmouseOver="menuHighlight(mnu4opt1)" onmouseOut="menuNormal(mnu4opt1)" onClick="imageRestore(HelpDown, HelpMenu)">Contents</td></tr>
	<tr><td id="mnu4opt2" onmouseOver="menuHighlight(mnu4opt2)" onmouseOut="menuNormal(mnu4opt2)" onClick="imageRestore(HelpDown, HelpMenu)">About OMIGA</td></tr>
	<tr><td>--------------------------</td></tr>
	<tr><td id="mnu4opt3" onmouseOver="menuHighlight(mnu4opt3)" onmouseOut="menuNormal(mnu4opt3)" onClick="imageRestore(HelpDown, HelpMenu)">Glossary of Terms</td></tr>
	</table>
</div>

<% /* Disabled Icons */ %>
<img src="images/presbutton.gif" width=27 height=28 alt="Presentations" border="0" style="position: absolute; left:125; top:1; z-index:-1" id="PresDisabled">
<img src="images/help.gif" width=27 height=28 alt="Help Menu" border="0" style="position: absolute; left:155; top:1; z-index:-1" id="HelpDisabled">
<img src="images/storedq.gif" width=27 height=28 alt="Stored Quotes" border="0" style="position: absolute; left:95; top:1; z-index:-1" id="QuoteDisabled">
<img src="images/underwriting.gif" width=27 height=28 alt="Application Underwriting" border="0" style="position: absolute; left:65; top:1; z-index:-1" id="AppRevDisabled">
<img src="images/printmenu.gif" width=27 height=28 alt="Print Menu" border="0" style="position: absolute; left:35; top:1; z-index:-1" id="PrintDisabled">
<img src="images/memopad.gif" width=27 height=28 alt="MemoPad" border="0" style="position: absolute; left:5; top:1; z-index:-1" id="MemoDisabled">

<script language="VBscript">
<!--
Sub imageSwap(img, menu)
	memoPadMenu.style.visibility = "hidden"
	AppRevMenu.style.visibility = "hidden"
	presMenu.style.visibility = "hidden"
	helpMenu.style.visibility = "hidden"
	memoDown.style.zindex="-2"
	AppRevDown.style.zindex="-2"
	PresDown.style.zindex="-2"
	HelpDown.style.zindex="-2"
	img.style.zindex="0"
	menu.style.visibility = "visible"
End Sub

Sub imageRestore(img, menu)
	img.style.zindex="-2"
	menu.style.visibility = "hidden"
End Sub

Sub noMenuImageSwap(img)
	memoPadMenu.style.visibility = "hidden"
	AppRevMenu.style.visibility = "hidden"
	presMenu.style.visibility = "hidden"
	helpMenu.style.visibility = "hidden"
	memoDown.style.zindex="-2"
	AppRevDown.style.zindex="-2"
	PresDown.style.zindex="-2"
	HelpDown.style.zindex="-2"
	img.style.zindex="0"
End Sub

Sub noMenuImageRestore(img)
	img.style.zindex="-2"
End Sub

Sub menuHighlight(menuID)
	menuID.style.cursor="hand"
	menuID.style.color="red"
End Sub

Sub menuNormal(menuID)
	menuID.style.cursor="default"
	menuID.style.color="white"
End Sub

Sub setCursor(menuID)
	menuID.style.cursor="hand"
End Sub

Sub restoreCursor(menuID)
	menuID.style.cursor="default"
End Sub
-->
</script>
