<% /*  
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      fw030.asp
Copyright:     Copyright © 2006 Marlborough Stirling

Description:   Top Menu/Title Bar
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
LH		22/09/06	Initial Revision
AW		27/10/06	CC56   EP1240  Make new DMS screens available without an application
AW		07/12/06	CC56   EP1240  Use new global parameter
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

This file is incorporated as an include file.  It must be included at the end of the screen design,
otherwise the popup menus will not appear.
This functionality assumes that all combos in a screen exist in the frmScreen form

To initialise the top bar correctly the following function must be called:

function FW030SetTitles(sScreenTitle,sScreenID,objScreenFunctions)
	sScreenTitle - the title of the screen
	sScreenID - the ID of the screen (XX999)
	objScreenFunctions - the screen functions object as declared in the calling screen
*/
%>

<div id="divFW030Menu" onmouseout="FW030HideMenus(false)" class="pageTitleBackground" style="HEIGHT: 50px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 625px; MENU: true">
	<span id="spnFW030First" tabindex="0"></span>

	<table id="tblFW030Menu" width="625" border="0" cellspacing="4" cellpadding="0" class="menuBackground" style="MENU: true">
	
		<tr><td class="menuText" width="100" align="center" style="BORDER-RIGHT: #6495ed 1pt solid; MENU: true">
			<A href="#" id=aFW030Notes onclick=FW030ShowMenus(aFW030Notes,divFW030Notes) onmouseover=FW030ShowMenus(aFW030Notes,divFW030Notes) style="MENU: true" tabIndex=-1 >Notes</a>
			<label id="lblFW030Notes" class="menuDisabledText" style="DISPLAY: none; MENU: true">Notes</label>
		</td>
		
		<!-- DR SYS???? New Print Options -->	
		<td class="menuText" width="80" align="center" style="BORDER-RIGHT: #6495ed 1pt solid; MENU: true">
			<A href="#" id=aFW030Print onclick="FW030ShowMenus(aFW030Print,divFW030Print)" onmouseover=FW030ShowMenus(aFW030Print,divFW030Print) style="MENU: true" tabIndex=-1 >Print</a>
			<label id="lblFW030Print" class="menuDisabledText" style="DISPLAY: none; MENU: true">Print</label> 
		</td>
		
		<td class="menuText" width="100" align="center" style="BORDER-RIGHT: #6495ed 1pt solid; MENU: true">
			<A href="#" id=aFW030Application onclick=FW030ShowMenus(aFW030Application,divFW030App) onmouseover=FW030ShowMenus(aFW030Application,divFW030App) style="MENU: true" tabIndex=-1 >Application</a>
			<label id="lblFW030Application" class="menuDisabledText" style="DISPLAY: none; MENU: true">Application</label> 
		</td>
		
		<td class="menuText" width="125" align="center" style="BORDER-RIGHT: #6495ed 1pt solid; MENU: true">
			<A href="#" id=aFW030StoredQuotes onclick="FW030ClickOption('CM065P')" onmouseover=FW030ShowMenus(aFW030StoredQuotes,null) style="MENU: true" tabIndex=-1 >Stored Quotes</a>
			<label id="lblFW030StoredQuotes" class="menuDisabledText" style="DISPLAY: none; MENU: true">Stored Quotes</label>
		</td>
		
		<td class="menuText" width="100" align="center" style="BORDER-RIGHT: #6495ed 1pt solid; MENU: true">
			<A href="#" id=aFW030Tools onclick=FW030ShowMenus(aFW030Tools,divFW030Tools) onmouseover=FW030ShowMenus(aFW030Tools,divFW030Tools) style="MENU: true" tabIndex=-1 >Tools</a>
		</td>
		
		<td class="menuText" width="155" style="PADDING-LEFT: 15px; MENU: true">
			<A href="#" id=aFW030Help onclick=FW030ShowMenus(aFW030Help,divFW030Help) onmouseover=FW030ShowMenus(aFW030Help,divFW030Help) style="MENU: true" tabIndex=-1 >Help</a>
		</td>
		
		<td class="menuText" width="155" style="PADDING-LEFT: 15px; MENU: true">
			<A href="#" id=aFW030TestingTools onclick=FW030ShowMenus(aFW030TestingTools,divFW030TestingTools) onmouseover=FW030ShowMenus(aFW030TestingTools,divFW030TestingTools) style="MENU: true" tabIndex=-1 >Testing Tools</a>
		</td></tr>
	
	</table>
	<input class="msgTxt" id="txtSelected" maxLength="1" style="LEFT: 600px; POSITION: absolute; TOP: -3px; VISIBILITY: hidden; WIDTH: 10px">
	<label id="lblFW030Title" style="LEFT: 10px; POSITION: absolute; TOP: 30px; WIDTH: 350px; MENU: true" class="msgPageTitle"></label>
	<label id="lblFW030ID" style="LEFT: 280px; POSITION: absolute; TOP: 33px; WIDTH: 50px; MENU: true" class="msgPageInfo"></label>	
	<label id="lblFW030Status" style="LEFT: 340px; POSITION: absolute; TOP: 33px; VISIBILITY: hidden; WIDTH: 70px; MENU: true" class="msgPageInfo"><font color="cyan">R/O</font></label>
	<label id="lblFW030ReviewStatus" style="LEFT: 385px; POSITION: absolute; TOP: 33px; VISIBILITY: hidden; WIDTH: 100px; MENU: true" class="msgPageInfo"><font color="cyan">U/R</font></label>
	<label id="lblFW030Stage" style="LEFT: 435px; POSITION: absolute; TOP: 33px; WIDTH: 190px; MENU: true" class="msgPageInfo"></label>
	
	<label id="lblFW030UserName" style="LEFT: 280px; POSITION: absolute; TOP: 23px; WIDTH: 200px; MENU: true" class="msgPageInfo"></label>

	<% /* Change width label from 200px to 190px to get rid of the unnecessary horizontal scroll bar at the bottom of most screens if you leave the window at its default width. */ %>
	<label id="lblFW030UnitName" style="LEFT: 435px; POSITION: absolute; TOP: 23px; WIDTH: 190px; MENU: true" class="msgPageInfo"></label>


<span id="spnFW030Last" tabindex="0"></span>
</div>
	
<div id="divFW030MenuMask" style="HEIGHT: 50px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 625px">
</div>

<div id="divFW030Collections" style="LEFT: 0px; POSITION: absolute; TOP: 17px">
	<div id="divFW030Notes" style="LEFT: 0px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden">
		<table onmouseout="FW030HideMenus(false)" width="130" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">
			<tr><td class="menuText" style="MENU: true">
				<A href="#" onclick="FW030ClickOption('GN100')" style="MENU: true" tabIndex=-1 >Memo Pad</a><br>
				<A href="#" onclick="FW030ClickOption('GN400')" style="MENU: true" tabIndex=-1 >Contact History</a>
			</td></tr>
		</table>
	</div>
	<!-- DR SYS???? New Print Options for DMS2 -->
	<div id="divFW030Print" style="LEFT: 75px; POSITION: absolute; TOP: 3px; VISIBILITY: hidden">
		<table onmouseout="FW030HideMenus(false)" width="250" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">
			<tr><td class="menuText" style="MENU: true">
				<!-- BS BM0271 20/02/02
				<A href="#" onclick="FW030ClickOption('PM010')" style="MENU: true" tabIndex=-1 >Printing</a><br>
				<A href="#" onclick="FW030ClickOption('DMS105')" style="MENU: true" tabIndex=-1 >Document History & Events</a>-->
				<div id="divFW030PrintingOptionEnabled" style="MENU: true">
					<A href="#" id="aFW030PrintingOption" onclick="FW030ClickOption('PM010')" style="MENU: true" tabIndex=-1 >Printing</a><br>
					<A href="#" onclick="FW030ClickOption('DMS105')" style="MENU: true" tabIndex=-1 >Document History &amp; Events</a>
				</div>
				<div id="divFW030PrintingOptionDisabled">
					<label id="lblFW030PrintingOption" class="menuDisabledText" style="DISPLAY: block; MENU: true">Printing</label>
					<A href="#" onclick="FW030ClickOption('DMS105')" style="MENU: true" tabIndex=-1 >Document History &amp; Events</a>
				</div>
				<!-- BS BM0271 End 20/02/02 -->
			</td></tr>
			<tr><td class="menuText" style="BORDER-TOP: #6495ed 1pt solid; MENU: true">
			
			</td></tr>
		</table>
	</div>
	<div id="divFW030App" style="LEFT: 149px; POSITION: absolute; TOP: 3px; VISIBILITY: hidden">
		<table onmouseout="FW030HideMenus(false)" width="160" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">
			<tr><td class="menuText" style="MENU: true">
				<A href="#" onclick="FW030ClickOption('AP011P')" style="MENU: true" tabIndex=-1 >Application Review</a>
				<A href="#" onclick="FW030ClickOption('AP040P')" style="MENU: true" tabIndex=-1 >Application Summary</a>
			</td></tr>
			<tr><td class="menuText" style="BORDER-TOP: #6495ed 1pt solid; MENU: true">
				<A href="#" onclick="FW030ClickOption('GN300P')" style="MENU: true" tabIndex=-1 >Completeness Check</a><br>
				
				<% /* BMIDS678 Restict access to Experian data */ %>
				<div id="divFW030CreditCheckOptionEnabled" style="MENU: true">
					<% /* MAR13 new credit check summary popup */ %>
					<A href="#" id="CreditCheckOptionEnabled" onclick="FW030ClickOption('RA021')" style="MENU: true" tabIndex=-1 >Credit Check</a><br>
					<% /* MAR13 End */ %>	
				</div>
				<div id="divFW030CreditCheckOptionDisabled">
					<label id="lblFW030CreditCheckOption" class="menuDisabledText" style="DISPLAY: block; MENU: true">Credit Check</label>
				</div>						
				<% /* BMIDS678 End */ %>	
				
				<A href="#" onclick="FW030ClickOption('RA012')" style="MENU: true" tabIndex=-1 >Case Assessment</a><br>
				<% /* BMIDS00336 MDC 05/09/2002
				<A href="#" onclick="FW030ClickOption('RA039')" style="MENU: true" tabIndex=-1 >Bureau Data</a><br> */ %>
				<A href="#" onclick="FW030ClickOption('AP300P')" style="MENU: true" tabIndex=-1 >Conditions Applied</a><br> 
				<% /* MAR28 wrap up */ %>
				<% /* <A href="#" onclick="FW030ClickOption('DC360')" style="MENU: true" tabIndex=-1 >Wrap Up</a><br> LDM 22/8/2006 EP1096 remove wrap up menu option */ %>
				<% /* EP1103  */ %>
				<A href="#" onclick="FW030ClickOption('TM040')" style="MENU: true" tabIndex=-1 >Progress Chart</a><br>
			</td></tr>
			<tr><td class="menuText" style="BORDER-TOP: #6495ed 1pt solid; MENU: true">
			
			</td></tr>
		</table>
	</div>	
	<div id="divFW030Tools" style="LEFT: 330px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden">
		<table onmouseout="FW030HideMenus(false)" width="150" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">
			<tr><td class="menuText" style="MENU: true">
				<A href="#" onclick="FW030ClickOption('GN500')" style="MENU: true" tabIndex=-1 >Currency Calculator</a>
				<A href="#" onclick="FW030ClickOption('mc040')" style="MENU: true" tabIndex=-1 >Flexible Mortgage Calculator</a><br>
			</td></tr>		
		</table>
	</div><!--original source safe stuff	<div id="divFW030Help" style="LEFT: 462px; TOP: 0px; POSITION: absolute; VISIBILITY: hidden">		<table onmouseout="FW030HideMenus(false)" width="130px" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">			<tr><td class="menuText" style="MENU: true"><%//				<a href="#" onclick="FW030ClickOption(null)" tabindex="-1" style="MENU: true">Contents</a>%>				<label class="menuDisabledText" style="MENU: true">Help Contents</label>			</td></tr>			<tr><td class="menuText" style="MENU: true">				<label class="menuDisabledText" style="MENU: true">Keyword Search</label>			</td></tr>			<tr><td class="menuText" style="MENU: true">				<label class="menuDisabledText" style="MENU: true">Help on this Screen</label>			</td></tr>			<tr><td class="menuText" style="BORDER-TOP: #6495ED 1pt solid; MENU: true"><%//				<a href="#" onclick="FW030ClickOption(null)" tabindex="-1" style="MENU: true">About OMIGA4</a><br>%>				<label class="menuDisabledText" style="MENU: true">About OMIGA4</label>			</td></tr>		</table>	</div>	-->
	<div id="divFW030Help" style="LEFT: 390px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden">
		<table onmouseout="FW030HideHelpMenus(false)" width="130" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true" id=TABLE1>
			<tr><td class="menuText" style="MENU: true">
				<A href="#" onclick="HandleHelp('Contents')" style="MENU: true" tabIndex=-1 >Contents</a><br>
				<A href="#" onclick="HandleHelp('Keywords')" style="MENU: true" tabIndex=-1 >Keywords</a><br>
				<A href="#" onclick="HandleHelp('Help')" style="MENU: true" tabIndex=-1 >Help on this screen</a><br>
			</td></tr>
			<tr><td class="menuText" style="BORDER-TOP: #6495ed 1pt solid; MENU: true"><!--			<a href="#" onclick="FW030ClickOption(null)" tabindex="-1" style="MENU: true">About OMIGA4</a><br>-->
				<A href="#" onclick="FW030ClickOption('SC035')" style="MENU: true" tabIndex=-1 ><!--#include file="customise/fw030customise.asp"--></a><br>
			</td></tr>			
		</table>
	</div>
	<div id="divFW030TestingTools" style="LEFT: 462px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden">
		<table onmouseout="FW030HideHelpMenus(false)" width="150" border="0" cellspacing="4" cellpadding="0" class="menuDropDown" style="MENU: true">
			<tr><td class="menuText" style="MENU: true">
				<A href="#" onclick="FW030ClickOption('GuiContext')" style="MENU: true" tabIndex=-1 >DEBUG: GUI Context</a><br>
				<!-- GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2 SG 11/06/02 SYS4856 Added Completions -->
				<a href="#" onclick="FW030ClickOption('cu010')" tabindex="-1" style="MENU: true">Completions</a><br>
				<A href="#" onclick="FW030ClickOption('DownAIP')" style="MENU: true" tabIndex=-1 >Download AIP</a><br>
				<A href="#" onclick="FW030ClickOption('DownloadFormFill')" style="MENU: true" tabIndex=-1 >Download Mortgage Application</a><br>

			</td></tr>
		</table>
	</div>
</div>

<script language="JScript" type="text/javascript">
<!--
var m_objFW030PrevFocus = null;
var m_bFW030IsOnMenu = false;
var m_FW030ScreenFunctions;
var m_sFW030ScreenTitle;
var m_sFW030ScreenID;
var m_FW030RequestArray;
var m_UserName;
var m_UnitName;

<% /* TW 04/11/2005 MAR425 */ %>
var iPrintMenuAuthorityLevel = 0;
var m_AccessType;
var GlobalXML = null;
<% /* TW 04/11/2005 MAR425 End */ %>


<%	// spnFW030First and spnFW030Last are there to ensure that any pop up menus
	// are closed down correctly if the user tabs off the menu bar
%>

function FW030ShowHelpMenus(AnchorID,DivID)
{
	// Get the object that the current focus is on.  If the object isn't null and if
	// it isn't one of the menu objects then remember it

	// Show the pop up menu if one exists and it is not already visible and hide any others
	// Else if there is no pop up menu hide any which may be visible
	// If a pop up is shown, hide any combos on the screen

	// Set the focus to the menu selection, setting its tabindex to 0
	// Flag that we are on a menu
	var objCurrentFocus = document.activeElement;
	//alert("FWO30 Show Help Menus");
	
	if(objCurrentFocus != null)
		//if(objCurrentFocus.style.getAttribute("MENU") == null)
		//	m_objFW030PrevFocus = objCurrentFocus;
			m_objFW030PrevFocus=null
	if(DivID != null)
	{
		if(DivID.style.visibility != "visible")
		{
			FW030DoHide();
			FW030DoComboProcess(false);
			DivID.style.visibility = "visible";
		}
	}
	else if(FW030DoHide()) FW030DoComboProcess(true);

	AnchorID.tabIndex = "0";
	AnchorID.focus();
	m_bFW030IsOnMenu = true;
}

function FW030HideHelpMenus(bAuto)
{
	// Only do if we are on a menu option
	// Check whether the mouse has moved over a non-menu element i.e. off the menu section.
	// If it has get rid of the pop up menu.
	// If bAuto is true then do this without checking
	var bOutOfMenu = false;
		
	if(!m_bFW030IsOnMenu) return;
	if(!bAuto)
	{
		if(window.event.toElement == null) bOutOfMenu = true;
		else if(window.event.toElement.style == null) bOutOfMenu = true;
		else if(window.event.toElement.style.getAttribute("MENU") == null) bOutOfMenu = true;
	}
	else bOutOfMenu = true;

	if(bOutOfMenu)
	{	
		if(FW030DoHide()) FW030DoComboProcess(true);
	//	if(m_objFW030PrevFocus != null) m_objFW030PrevFocus.focus();
		m_bFW030IsOnMenu = false;
	}
}

function HandleHelp(sScreen)
{
switch(sScreen)
	{
		case "Contents":
			txtSelected.value = "1";
			window.open("HELP/HP030.asp","Paul","toolbar=no width=550 height=525 top=10 left=120");
		//menubar=no resizeable=no height=500 width=500
			return;
		break;	
		case "Keywords":
			txtSelected.value = "2";
			window.open("HELP/HP030.asp","Paul","toolbar=no width=550 height=525 top=10 left=120");
			return;
		break;	
		case "Help":
			txtSelected.value = "3";
			window.open("HELP/HP030.asp","Paul","toolbar=no width=550 height=525 top=10 left=120");
			return;
		break;	
	}
}


function spnFW030First.onfocus()
{
	FW030HideMenus(true);
}

function spnFW030Last.onfocus()
{
	FW030HideMenus(true);
}

function FW030DoHide()
{
<%	// Set the tabIndex on all top level anchors to -1
	// Hide the visible pop-up menu, if there is one
	// Return if there was one visible or not
%>	var bIsVisible = false;

	for(var nALoop = 0;nALoop < divFW030Menu.all.length;nALoop++)
	{
		var sId = divFW030Menu.all(nALoop).id;
		if(sId.indexOf("aFW030") != -1) divFW030Menu.all(nALoop).tabIndex = "-1";
	}
		
	for(var nDivLoop = 0;nDivLoop < divFW030Collections.all.length;nDivLoop++)
	{
		var sId = divFW030Collections.all(nDivLoop).id;
		if(sId.indexOf("divFW030") != -1)
			if(divFW030Collections.all(nDivLoop).style.visibility == "visible")
			{
				bIsVisible = true;
				divFW030Collections.all(nDivLoop).style.visibility = "hidden";
			}
	}

	return bIsVisible;
}

function FW030DoComboProcess(bDisplay)
{
<%	// Display or hide any combos which exist in the calling screen's form
	// THIS MUST BE CALLED frmScreen
	// This is required because combos always appear as the topmost object
	// It is assumed that normal screen functionality does not set a combo's display
	// attribute to "none"
%>

	<% /* MAR7 GHun
	var sDisplay = "block";
	if(!bDisplay) sDisplay = "none";
	
	for(var nLoop = 0;nLoop < frmScreen.elements.length;nLoop++)
		if(frmScreen.elements(nLoop).tagName == "SELECT")
			frmScreen.elements(nLoop).style.display = sDisplay;
	*/ %>
	
	if(!(document.forms("frmScreen"))) return;
	for(var nLoop = 0;nLoop < frmScreen.all.tags("SELECT").length;nLoop++)
	{
		var o = frmScreen.all.tags("SELECT").item(nLoop);
		if(!(o.menusafe && o.menusafe == "true"))
		{
			if(o.menuhide && o.menuhide == "visibility")
				o.style.visibility = bDisplay ? "visible" : "hidden";
			else
				o.style.display = bDisplay ? "block" : "none";
		}
	}
	<% /* MAR7 End */ %>
}

function FW030HideMenus(bAuto)
{
<%	// Only do if we are on a menu option
	// Check whether the mouse has moved over a non-menu element i.e. off the menu section.
	// If it has get rid of the pop up menu.
	// If bAuto is true then do this without checking
%>	var bOutOfMenu = false;

	if(!m_bFW030IsOnMenu) return;
	if(!bAuto)
	{
		if(window.event.toElement == null) bOutOfMenu = true;
		else if(window.event.toElement.style == null) bOutOfMenu = true;
		else if(window.event.toElement.style.getAttribute("MENU") == null) bOutOfMenu = true;
	}
	else bOutOfMenu = true;

	if(bOutOfMenu)
	{	
		if(FW030DoHide()) FW030DoComboProcess(true);
		if(m_objFW030PrevFocus != null) m_objFW030PrevFocus.focus();
		m_bFW030IsOnMenu = false;
	}
}

function FW030ShowMenus(AnchorID,DivID)
{
<%	// Get the object that the current focus is on.  If the object isn't null and if
	// it isn't one of the menu objects then remember it

	// Show the pop up menu if one exists and it is not already visible and hide any others
	// Else if there is no pop up menu hide any which may be visible
	// If a pop up is shown, hide any combos on the screen

	// Set the focus to the menu selection, setting its tabindex to 0
	// Flag that we are on a menu
%>	var objCurrentFocus = document.activeElement;
	if(objCurrentFocus != null)
		if(objCurrentFocus.style.getAttribute("MENU") == null)
			m_objFW030PrevFocus = objCurrentFocus;

	if(DivID != null)
	{
		if(DivID.style.visibility != "visible")
		{
			FW030DoHide();
			FW030DoComboProcess(false);
			DivID.style.visibility = "visible";
		}
	}
	else if(FW030DoHide()) FW030DoComboProcess(true);

	AnchorID.tabIndex = "0";
	AnchorID.focus();
	m_bFW030IsOnMenu = true;
}

function FW030ClickOption(sScreen)
{
<%	// If an option is clicked do its process
%>	if(FW030DoHide()) FW030DoComboProcess(true);
	if(sScreen == null)
	{
		alert("No action specified for this option");
		return;
	}

	var ArrayArguments = new Array();
	var nWidth;
	var nHeight;
 
	switch(sScreen)
	{
		case "DownloadAIP":   //DOWNLOAD AIP
		case "DownloadFormFill": // DOWNLOAD MORTGAGE APPLICATION
			FW030PerformDownload(sScreen);
			return;
		break;
		case "GN100":  //NOTES
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[3] = m_sFW030ScreenID;
			ArrayArguments[4] = m_sFW030ScreenTitle;
			ArrayArguments[5] = m_FW030RequestArray;
			<% /* BS BM0271 24/04/2003 Remove IsMainScreenReadOnly indicator, add ReadOnly and ProcessingInd  */ %>
			<% /* BS BM0271 20/02/2003 Add ReadOnly indicator */ %>
			//ArrayArguments[6] = scScreenFunctions.IsMainScreenReadOnly(window, "");
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idReadOnly",null);
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
			nWidth = 630;
			nHeight = 485;
		break;
		case "GN300P":  //COMPONENTS CHECK
			if(m_sFW030ScreenID == "CM010") ArrayArguments[0] = true;
			else ArrayArguments[0] = false;
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			//var sValue = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationStage",null);
			var sValue = m_FW030ScreenFunctions.GetContextParameter(window,"idStageId",null);
			ArrayArguments[3] = sValue;
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[5] = m_FW030RequestArray;
			nWidth = 630;
			nHeight = 540;
		break;
		case "GN400": //CONTACT HISTORY
			var aCustomerNameArray = new Array();
			var aCustomerNumberArray = new Array();
			var aCustomerVersionNumberArray = new Array();	<% /* BMIDS00442 */ %>
			var iCount = 0;
			var sCustomerName = "";
			var sCustomerNumber = "";
			var sCustomerVersionNumber = "";   <% /* BMIDS00442 */ %>
			for (iCount = 1; iCount <= 5; iCount++)
			{								
				sCustomerNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
				if (sCustomerNumber != "")
				{				
					sCustomerName = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName" + iCount,null);
					sCustomerVersionNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null); <% /* BMIDS00442 */ %>
					aCustomerNameArray[iCount-1] = sCustomerName;
					aCustomerNumberArray[iCount-1] = sCustomerNumber;
					aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber; <% /* BMIDS00442 */ %>
				}
			}			
			ArrayArguments[0] = aCustomerNameArray;
			ArrayArguments[1] = aCustomerNumberArray;
			<% /* BS BM0271 20/02/2003 */ %>
			//ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idReadOnly",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.IsMainScreenReadOnly(window, "GN400"); <% /* MAR1675 GHun */ %>
			<% /* BS BM0271 End 20/02/2003 */ %>
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[5] = m_FW030RequestArray;
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserName",null);
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitName",null);
			ArrayArguments[8] = m_FW030ScreenFunctions.GetContextParameter(window,"idRole",null);
			ArrayArguments[9] = aCustomerVersionNumberArray;
			nWidth  = 630; 
			nHeight = 355; 		
		break;	
		case "RA012": //Case Assessment
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			sValue = m_FW030ScreenFunctions.GetContextParameter(window,"idStageId",null);
			ArrayArguments[4] = sValue;
			ArrayArguments[5] = m_FW030RequestArray;
			ArrayArguments[6] = m_FW030ScreenFunctions.IsMainScreenReadOnly(window, "RA012");
			<% /*
			Peter Edney - 05/04/06
			MAR1563
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null);
			*/ %>
			<% /* MAR1447 GHun */ %>

			var frozen = "0";
			var sDataFreezeIndicator = m_FW030ScreenFunctions.GetContextParameter(window, "idFreezeDataIndicator", "0");
			if (sDataFreezeIndicator == "1") 
			{
				if (m_FW030ScreenFunctions.IsDataFreezeScreen(window, "RA012") == true)	<% /* MAR1618 GHun */ %>
				{
					frozen = "1";
				}
			}
			ArrayArguments[7] = frozen;
			<% /* MAR1447 End */ %>
			ArrayArguments[8] = false;   // Called from Accept Quote
			ArrayArguments[9] = m_FW030ScreenFunctions.GetContextParameter(window,"idRole",null); //MAR253 - UserRole

			nWidth = 650;
			nHeight = 620;        // MAR667
		break;	
		case "RA021": //CREDIT CHECK MAR13 "ra20" is changed to "ra21"
			<% /* BMIDS00336 MDC 05/09/2002 */ %>
			/*
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
			ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerRoleType2",null);
			ArrayArguments[6] = m_FW030RequestArray;
			*/
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
			if(m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerRoleType2",null)=="1")
			{
				ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
				ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
				ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
			}
			else
			{
				ArrayArguments[4] = "";
				ArrayArguments[5] = "";
				ArrayArguments[7] = "";
			}
			ArrayArguments[8] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[9] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			
			ArrayArguments[10] = m_FW030RequestArray ;
			ArrayArguments[11] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerRoleType2",null);
			<% /* BMIDS00336 MDC 05/09/2002 - End */ %>
			
			<% /* Old Screen RA020 */ %>
			/* nWidth = 623;
			nHeight = 570; */
			<% /* Old Screen RA020 - End */ %>
			// Peter Edney - MAR309 - 13/12/2005
			// nWidth = 626;
			// nHeight = 580; 
			nWidth = 627;
			nHeight = 610; 			
		break;
		
		case "DC360": //WRAP UP MAR29 
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030RequestArray;
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
			ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
			
			<% /* Old Screen DC360 */ %>
			/*nWidth = 626;
			nHeight = 580;  */
			<% /* Old Screen DC360 - End */ %>
			<% /* MAR126 Maha T - Adjust Screen Size */ %>
			nWidth = 626;
			nHeight = 590; 
		break;

		case "RA039": //BUREAU DATA
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
			if(m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerRoleType2",null)=="1")
			{
				ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
				ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
				ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
			}
			else
			{
				ArrayArguments[4] = "";
				ArrayArguments[5] = "";
				ArrayArguments[7] = "";
			}
			ArrayArguments[8] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[9] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			
			ArrayArguments[10] = m_FW030RequestArray ;
			nWidth  = 450;
			nHeight = 280;
		break;
		case "AP300P": // Conditions Applied
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030RequestArray ;
			nWidth = 630;
			nHeight = 500;
		break;
		case "GuiContext": //GUI CONTEXT
			ArrayArguments[0] = window.parent.frames("omigamenu").document.forms("frmContext");
			nWidth = 430;
			nHeight = 455;
		break;
		//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2
		case "cu010": //Completions		SG 11/06/02 SYS4856
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			nWidth = 500;
			nHeight = 710;
		break;
		case "mc040": //FLEXIBLE MORTGAGE CALCULATOR 
			ArrayArguments[0] = window.parent.frames("omigamenu").document.forms("frmContext");
			nWidth = 465;
			nHeight = 725;
		break;
		case "GN500": //CURRENCY CALCULATOR
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idMachineId",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			nWidth = 390; 
			nHeight = 310;
			nTop = 300;
			nLeft = 320;
		break;
		case "AP011P": //APPLICATION REVIEW (POPUP) 
			//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var ArrayArguments = new Array();
			ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
			ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
			nWidth = 630; 
			nHeight = 380; 
		break;
		case "CM065P": //STORED QUOTES (POPUP) 
			//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var ArrayArguments = new Array();
			ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");//	m_sApplicationMode
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idReadOnly",0);			// m_sReadOnly
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");	//	m_sApplicationNumber
			ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",""); // m_sApplicationFactFindNumber
			ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idQuotationNumber",""); //m_sActiveQuotationNumber
			ArrayArguments[6] = scScreenFunctions.GetContextParameter(window,"idMetaAction",""); //m_sMetaAction
			ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idXML",""); //m_sidXML
			nWidth = 630; 
			nHeight = 350; 
		break;
		case "SC035": //ABOUT OMIGA 
			nWidth = 470; 
			nHeight = 425; 
			break;
		case "PM010": //PRINT
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); <% /* MAR7 GHun */ %>
			var aCustomerNameArray = new Array();
			var aCustomerNumberArray = new Array();
			var aCustomerVersionNumberArray = new Array();			
			var iCount = 0;
			var sCustomerName = "";
			var sCustomerNumber = "";
			var sCustomerVersionNumber = "";
			for (iCount = 1; iCount <= 5; iCount++)
			{								
				sCustomerNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
				if (sCustomerNumber != "") 
				{				
					sCustomerName = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName" + iCount,null);
					sCustomerVersionNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);
					aCustomerNameArray[iCount-1] = sCustomerName;
					aCustomerNumberArray[iCount-1] = sCustomerNumber;
					aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber;			
				}
			}			
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idStageid",null);
			ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idRole",null);
			ArrayArguments[8] = m_FW030RequestArray;
			ArrayArguments[9] = aCustomerNameArray;
			ArrayArguments[10] = aCustomerNumberArray;
			ArrayArguments[11] = aCustomerVersionNumberArray;
			<% /* MAR7 GHun */ %>
			ArrayArguments[16] = XML.GetGlobalParameterString(document,"EmailAdministrator");
			ArrayArguments[17] = m_FW030ScreenFunctions.GetContextParameter(window,"idMachineId",null);
			ArrayArguments[18] = m_FW030ScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			ArrayArguments[19] = m_FW030ScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
			<% /* MAR7 End */ %>
			nWidth = 535;
			nHeight = 570;
		break;

		case "DMS105":  //NOTES
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); <% /* MAR7 GHun */ %>
			
			//NEW STYLEE
			var aCustomerNameArray = new Array();
			var aCustomerNumberArray = new Array();
			var aCustomerVersionNumberArray = new Array();			
			var iCount = 0;
			var sCustomerName = "";
			var sCustomerNumber = "";
			var sCustomerVersionNumber = "";
			for (iCount = 1; iCount <= 5; iCount++)
			{								
				sCustomerNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
				if (sCustomerNumber != "") 
				{				
					sCustomerName = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName" + iCount,null);
					sCustomerVersionNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);
					aCustomerNameArray[iCount-1] = sCustomerName;
					aCustomerNumberArray[iCount-1] = sCustomerNumber;
					aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber;			
				}
			}			
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idStageid",null);
			ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idStageName",null);
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
			ArrayArguments[8] = m_FW030ScreenFunctions.GetContextParameter(window,"idRole",null);						
			
			ArrayArguments[9] = m_FW030ScreenFunctions.GetContextParameter(window,"idMachineId",null);
			ArrayArguments[10] = m_FW030ScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			
			ArrayArguments[11] = m_FW030RequestArray;
			ArrayArguments[12] = aCustomerNameArray
			ArrayArguments[13] = aCustomerNumberArray
			ArrayArguments[14] = aCustomerVersionNumberArray
			<% /* BS BM0271 19/02/03 Add parameter*/ %>
			ArrayArguments[15] = m_FW030ScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
			<% /* BS BM0271 16/04/03 Add ReadOnly parameter*/ %>
			ArrayArguments[16] = scScreenFunctions.GetContextParameter(window,"idReadOnly",0);	
			ArrayArguments[17] = XML.GetGlobalParameterString(document,"EmailAdministrator");	<% /* MAR7 GHun */ %>
			<% /* AW EP1240	27/10/06  - Route to new DMS screen if configured*/ %>
			ArrayArguments[18] = scScreenFunctions.GetContextParameter(window,"idIsUserQualityChecker",0);
			ArrayArguments[19] = scScreenFunctions.GetContextParameter(window,"idIsUserFulfillApproved",0);
			ArrayArguments[20] = scScreenFunctions.GetContextParameter(window,"idCorrespondenceSalutation",0);
			
			var bGeminiPrintEnabled = XML.GetGlobalParameterBoolean(document,"GeminiPrintEnabled", 0);
			
			if(bGeminiPrintEnabled)
			{
				sScreen = "DMS110";
				nWidth = 640;
				nHeight = 600;
			}
			else
			{
				nWidth = 630;
				nHeight = 460;
			}
		break;		
		<% /*AW	EP1103 */ %>
		case "TM040":  //PROGRESS CHART
			
			var aCustomerNameArray = new Array();
			var aCustomerNumberArray = new Array();
			var aCustomerVersionNumberArray = new Array();			
			var iCount = 0;
			var sCustomerName = "";
			var sCustomerNumber = "";
			var sCustomerVersionNumber = "";
			for (iCount = 1; iCount <= 5; iCount++)
			{								
				sCustomerNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
				if (sCustomerNumber != "") 
				{				
					sCustomerName = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName" + iCount,null);
					sCustomerVersionNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);
					aCustomerNameArray[iCount-1] = sCustomerName;
					aCustomerNumberArray[iCount-1] = sCustomerNumber;
					aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber;			
				}
			}
						
			ArrayArguments[0] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[1] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idReadOnly",1);
			ArrayArguments[3] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_FW030ScreenFunctions.GetContextParameter(window,"idStageid",null);
			ArrayArguments[5] = m_FW030ScreenFunctions.GetContextParameter(window,"idStageName",null);
			ArrayArguments[6] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
			ArrayArguments[7] = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
			ArrayArguments[8] = m_FW030ScreenFunctions.GetContextParameter(window,"idRole",null);						
			
			ArrayArguments[9] = m_FW030ScreenFunctions.GetContextParameter(window,"idMachineId",null);
			ArrayArguments[10] = m_FW030ScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			
			ArrayArguments[11] = m_FW030RequestArray;
			ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idReadOnly",0);	
			ArrayArguments[13] = aCustomerNameArray;
			ArrayArguments[14] = aCustomerNumberArray;
			ArrayArguments[15] = aCustomerVersionNumberArray;
			ArrayArguments[16] = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationPriority","0");
			
			nWidth = 650;
			nHeight = 700;
		break;
		
		case "AP040P":  //Full application summary screen
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var ArrayArguments = new Array();
			
			ArrayArguments[0] = XML.CreateRequestAttributeArray(window);
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
			ArrayArguments[2] = m_FW030ScreenFunctions.IsMainScreenReadOnly(window, "idReadOnly");
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idAppUnderReview",null);
			ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idStageName",null);
			ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idStageId",null);
			ArrayArguments[6] = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);	
			ArrayArguments[8] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
			ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
			ArrayArguments[10] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
			ArrayArguments[11] = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
			ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber3",null);
			ArrayArguments[13] = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber3",null);
			ArrayArguments[14] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber4",null);
			ArrayArguments[15] = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber4",null);
			ArrayArguments[16] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber5",null);
			ArrayArguments[17] = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber5",null);
			
			ArrayArguments[18] = scScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
			ArrayArguments[19] = scScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
			ArrayArguments[20] = scScreenFunctions.GetContextParameter(window,"idCustomerName3",null);
			ArrayArguments[21] = scScreenFunctions.GetContextParameter(window,"idCustomerName4",null);
			ArrayArguments[22] = scScreenFunctions.GetContextParameter(window,"idCustomerName5",null);
			ArrayArguments[23] = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitName",null);			
			/*var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var ArrayArguments = new Array();
			ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
			ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);*/
				
			nWidth = 650;
			nHeight = 600;
		break;
		
		default:
			alert(sScreen + ".ASP will be displayed here");
			return;
		break;
	}

	var rtnValue = m_FW030ScreenFunctions.DisplayPopup(window, document, sScreen + ".asp", ArrayArguments, nWidth, nHeight);
	
	<% /* Start: MAR57 - Maha T  */ %>
	<% /* Close Omiga Application (Right now this return value is paased from dc360.asp */ %>
	if (rtnValue == "exitOmiga")
	{
		// close Omiga 
		window.top.close();
	}
	<% /* End: MAR57 */ %>
}

function SetUserAndUnitNames()
{
	//Set the Unit and User names on the screen 
	//Extract UserID and UnitID from context
	m_UserName = m_FW030ScreenFunctions.GetContextParameter(window,"idUserName",null);
	m_UnitName = m_FW030ScreenFunctions.GetContextParameter(window,"idUnitName",null);
	m_FW030ScreenFunctions.SizeTextToField(lblFW030UnitName,m_UnitName);
	m_FW030ScreenFunctions.SizeTextToField(lblFW030UserName,m_UserName);
	
}

function FW030SetTitles(sScreenTitle,sScreenID,objScreenFunctions)
{

<%	// Display the top menu table (its hidden on entry so that its functionality isn't accidentally kicked off)
	// Show the screen title/ID
	// Show the read only status if applicable
	// Disable application specific options if no application
	// Show the application stage if applicable

	// NOTE - this function expects there to be an scXMLFunctions in the parent screen
%>	m_FW030ScreenFunctions = objScreenFunctions;
	m_sFW030ScreenTitle = sScreenTitle;
	m_sFW030ScreenID = sScreenID;
	
	SetUserAndUnitNames();
	
	//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_FW030RequestArray = XML.CreateRequestAttributeArray(window);
	XML = null;

	divFW030MenuMask.style.display = "none";
	m_FW030ScreenFunctions.SizeTextToField(lblFW030Title,sScreenTitle);
	m_FW030ScreenFunctions.SizeTextToField(lblFW030ID,sScreenID);
	//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2 
	<% /* BS BM0271 19/02/03 */ %>
	//if(m_FW030ScreenFunctions.GetContextParameter(window,"idReadOnly","0") == "1")
	if (m_FW030ScreenFunctions.IsMainScreenReadOnly(window, ""))
	{

		var sReadOnly = m_FW030ScreenFunctions.GetContextParameter(window,"idReadOnly",null);
		var sProcessingIndicator = m_FW030ScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
		var sDataFreezeFlag =  m_FW030ScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null);
		var sCDDataFreezeFlag = m_FW030ScreenFunctions.GetContextParameter(window,"idCancelDeclineFreezeDataIndicator",null);

		if ((sReadOnly == "1") || (sProcessingIndicator =="0"))
		{

			lblFW030Status.style.visibility = "visible";
			//CL SYS4456 Disables print menu if we are in read only mode
			<% /* BS BM0271 19/02/03 
			Always enable print menu, just disable printing option*/ %>		
			//aFW030Print.style.display = "none";
			//lblFW030Print.style.display = "block";
			divFW030PrintingOptionEnabled.style.display = "none";
			//divFW030PrintingOptionDisabled.style.display = "block";<% //BM0567-ensure printing disabled%>
			<% /* BS BM0271 End 19/02/03 */ %>	
		}
		else
		{
			//Not 'true' readonly, just data freeze

			<%/* EP989 - Check if either data freeze flag has been set */%>

			if (((sReadOnly != "1") && (sProcessingIndicator == "1")) && ((sDataFreezeFlag == "1") || (sCDDataFreezeFlag == "1")))
			{
				lblFW030Status.style.visibility = "hidden"; 
				divFW030PrintingOptionDisabled.style.display = "none";
				divFW030PrintingOptionEnabled.style.display = "block";
			}		
		
		}
		
	}
	else 
	{
		divFW030PrintingOptionDisabled.style.display = "none";
		//divFW030PrintingOptionEnabled.style.display = "block";<% //BM0567-ensure printing enabled%>
		lblFW030Status.style.visibility = "hidden"; <%//BM0567-Remove R/O (if displayed)%>
	}


	
	// APS SYS1920
	if(m_FW030ScreenFunctions.GetContextParameter(window,"idAppUnderReview","0") == "1") lblFW030ReviewStatus.style.visibility = "visible";

	GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if(m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber","") == "")
	{		
		aFW030Notes.style.display = "none";
		lblFW030Notes.style.display = "block";
		aFW030Application.style.display = "none";
		lblFW030Application.style.display = "block";
		aFW030StoredQuotes.style.display = "none";
		lblFW030StoredQuotes.style.display = "block";
		<% /* AW EP1240	27/10/06  - Allow acces to DMS without an application new if configured*/ %>
		var bGeminiPrintEnabled = GlobalXML.GetGlobalParameterBoolean(document,"GeminiPrintEnabled", 0);
		
		if(bGeminiPrintEnabled)
		{
			divFW030PrintingOptionDisabled.style.display = "none";
			aFW030PrintingOption.style.display = "none";
		}
		else
		{
			aFW030Print.style.display = "none";
			lblFW030Print.style.display = "block";
		}
	}
	else
	{
		//var sDefault = new Array(null,null);
		//var sStage = m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationStage",sDefault);
		var sStageName = m_FW030ScreenFunctions.GetContextParameter(window,"idStageName",null);
		m_FW030ScreenFunctions.SizeTextToField(lblFW030Stage,sStageName);
	}
	
	// HMA  BMIDS678
	if(m_FW030ScreenFunctions.GetContextParameter(window,"idCreditCheckAccess","0") == "1") 
	{
		divFW030CreditCheckOptionDisabled.style.display = "none";
	}
	else
	{
		divFW030CreditCheckOptionEnabled.style.display = "none";
	}
	<% /* TW 04/11/2005 MAR425 */ %>
	m_AccessType = parseInt(scScreenFunctions.GetContextParameter(window, "idRole", "0"));

	iPrintMenuAuthorityLevel = GlobalXML.GetGlobalParameterAmount(document,'PrintMenuAuthorityLevel');

	// Disable Print menu if User Access Type in context is less than GlobalParameter PrintMenuAuthority
	if (m_AccessType < iPrintMenuAuthorityLevel)
	{
		aFW030Print.style.display = "none";
		lblFW030Print.style.display = "block";
	}
	<% /* TW 04/11/2005 MAR425 End */ %>

	// MAR 1310 - Peter Edney - 27/02/2006
	// Disable testing menu if the user role is < than then global parameter: TestingRole
	if(m_FW030ScreenFunctions.GetContextParameter(window,"idAllowTesting","") == "")
	{
		var iTestingRole = GlobalXML.GetGlobalParameterAmount(document,'TestingRole');
		if (m_AccessType < iTestingRole)
		{
			m_FW030ScreenFunctions.SetContextParameter(window,"idAllowTesting","false");
		}
		else
		{
			m_FW030ScreenFunctions.SetContextParameter(window,"idAllowTesting","true");
		}		
	}		
	if(m_FW030ScreenFunctions.GetContextParameter(window,"idAllowTesting","false") == "false")
	{
		aFW030TestingTools.style.display = "none";
	}
	
}

function FW030PerformDownload(sOption)
{
	switch(sOption)
	{
		case "DownloadAIP":
			var sType = "AIP";
		break;
		case "DownloadFormFill":
			var sType = "FormFill";
		break;
		default:
			alert("Invalid Option for FW030PerformDownload");
			return;
		break;
	}

	if(confirm("Are you sure you wish to download to Omiga3?"))
	{
		//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,"DOWNLOAD");
		XML.CreateTag("TYPE",sType);
		XML.CreateTag("APPLICATIONNUMBER",m_FW030ScreenFunctions.GetContextParameter(window,"idApplicationNumber",""));
		XML.RunASP(document,"DownloadToOmiga3.asp");
		if(XML.IsResponseOK())
			alert("Application successfully downloaded. Omiga3 Application Number is " + XML.GetTagText("Application_Number"));
	}
}
function FW030GetStageInfo(sStageId)
{
<%	/* Returns an array holding the stagename in [0] and the stagesequencenumber in [1]
		from the STAGE table for the given stageid */
%>	//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var saReturnArray = new Array();
	
	// ik_bm0396 - do not retrieve task details
	// XML.CreateRequestTag(window, "GetStageDetail");
	XML.CreateRequestTag(window, "GetStageList");
	// ik_bm0396 - ends

	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("STAGEID", sStageId);
	XML.RunASP(document, "MsgTMBO.asp");
	
	<% /* MAR7 GHun
	   if(XML.IsResponseOK())
	*/ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0] == true)
	<% /* MAR7 End */ %>
	{
		XML.SelectTag(null, "STAGE");
		saReturnArray[0] = XML.GetAttribute("STAGENAME");
		saReturnArray[1] = XML.GetAttribute("STAGESEQUENCENO");
	}
	XML = null;
	return(saReturnArray);
}
-->
</script>
