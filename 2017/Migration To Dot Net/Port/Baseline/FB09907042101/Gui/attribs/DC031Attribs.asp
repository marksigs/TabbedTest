<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	frmScreen.idVerificationDate.setAttribute("filter","[0-9/]");
	frmScreen.idVerificationDate.setAttribute("date","DD/MM/YYYY");
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.idIssuer.style.textTransform = "capitalize";
	//END: MAR36
}

<% /* Get data required for client validation here */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
<% /* BMIDS692 - DRC 03/02/04  */ %>

  /* START: CODE COMMENTED BY Joyce Joseph AQR-MAR59  
  if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Rev_and_Decis")=="1") 
  {
    
  //temporarily overide the main context freezedata to disable read only status for the whole of this screen
    
    m_blnReadOnly = false;
    scScreenFunctions.SetCollectionState(divCustomerName, "W");
    //scScreenFunctions.SetCollectionState(divPersonalId, "W");
    //frmScreen.cboPersonalType1.onchange();
    //frmScreen.cboPersonalType2.onchange();
    //scScreenFunctions.SetCollectionState(divResidencyId, "W");
    //frmScreen.cboResidencyType1.onchange();
    //frmScreen.cboResidencyType2.onchange();
    
  } 
	END: CODE COMMENTED BY Joyce Joseph AQR-MAR59 */
}

function PopulateComboEx(comboName,objScreenCombo,clearCombo)
{
	var comboXML = null;
	var XML =null;
	var sGroupList = null;
	var blnSuccess = true;
	
	if(comboName==null || comboName=='')
	{
		return;
	}
	
	if(clearCombo==1)
	{
		while(objScreenCombo.options.length > 0) 
			objScreenCombo.remove(0);
	}
	
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sGroupList = new Array(comboName);

	if(XML!=null)
	{
		if(XML.GetComboLists(document, sGroupList))
		{
			comboXML = XML.GetComboListXML(comboName);
			
			try
			{
				blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, objScreenCombo,comboXML,true);
			}
			catch(ex)
			{
				//Error Occured set to false
				blnSuccess = false;
			}
			
			if(!blnSuccess)
				scScreenFunctions.SetScreenToReadOnly(frmScreen);
		}
	}

	XML = null;		
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
