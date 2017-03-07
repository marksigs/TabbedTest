<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile	:   DC035attribs.asp
Copyright	:   Copyright © 1999 Marlborough Stirling

Description	:   Alias/Association Details Screen attributes
Author		:	Mallesh Cheekoti
Date		:	18/05/2004
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks()
{
	
}

<% /* Get data required for client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ClientPopulateScreen() 
{

}

function getComboXML(comboName)
{
	var comboXML = null;
	var XML =null;
	var sGroupList = null;
	var blnSuccess = true;
	
	if(comboName==null || comboName=='')
	{
		return;
	}
	
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sGroupList = new Array(comboName);

	if(XML!=null)
	{
		if(XML.GetComboLists(document, sGroupList))
		{
			comboXML = XML.GetComboListXML(comboName);
				
		}
	}

	XML = null;		
	return comboXML;
}


</SCRIPT>