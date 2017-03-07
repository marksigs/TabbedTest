<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
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
<% /* BMIDS692 - DRC - allow access to dc270 screen and allow adding a new third party  */ %>
  /* START: CODE COMMENTED BY Joyce Joseph AQR-MAR59 		
  if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Rev_and_Decis")=="1") 
  {
    m_blnReadOnly = false;  
    frmScreen.btnAdd.disabled = false;
    //Ensure that only Bank & Building societies can be added
    
	// Mark those third party types in the list for deletion which are not bank/building societies
	
	*/<% /* remove all but <Select> & bank option */ %>
	/*for (iCount = frmScreen.cboThirdPartyType.length - 1; iCount > 0; iCount--)
	{
	  // bank/building societies are hard coded 
		if (!(frmScreen.cboThirdPartyType.item(iCount).value=="3"))
			 frmScreen.cboThirdPartyType.remove(iCount);
	}
	frmScreen.cboThirdPartyType.disabled = false;
	frmScreen.cboThirdPartyType.style.backgroundColor = "white";
	frmScreen.btnContinue.disabled = false;
  }
     END: CODE COMMENTED BY Joyce Joseph AQR-MAR59 */
  
}

</SCRIPT>
