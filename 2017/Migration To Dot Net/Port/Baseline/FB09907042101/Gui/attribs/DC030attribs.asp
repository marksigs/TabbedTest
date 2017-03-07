<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC030attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Personal Details Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		18/11/99	Created
JLD		10/12/1999	DC/026 - let electoral roll date be optional
SA		23/05/01	SYS1053 Added mask for Year Added to Electoral Roll
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		31/05/2002	BMIDS00013 - Amended Filter for NumberOfDependants
MV		20/08/2002	BMIDS00204	Restricted the NoOfDependents to two Digits
MV		23/08/2002	BMIDS00204	Made NumberOfDependents Mandatory
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
MF		27/07/2005	MAR19	WP01 personal fields apart from DOB are now readonly
							DOB readonly status set in DC030
JJ		27/09/2005  MAR84   Disable Other Forenames on Screens					
JJ		07/10/2005  MAR119  NumberOfDependants only between 0-7 
HMA     08/10/2005  MAR135  Correct use of CustomerXML
MahaT	10/10/2005	MAR95	Month (Time at bank) should not accept values > 11.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific history

SW		15/06/2006  EP705	Added Capitalisation for txtOtherTitle
MAH		08/12/2006	EP2_346	Text only Filter added to names
AShaw	04/01/2007	EP2_610 Add mandatory attribs to fields (reversing some MAR19 changes).
DS		18/01/2007  EP2_610 Removed SpecialNeeds combo as mandatory field as per new requirement
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~--------------------------------~~*/
function SetMasks()
{
	// EP2_610 re-instate Mandatory attribs + add Special needs.
	frmScreen.txtSurname.setAttribute("required", "true");
	frmScreen.txtFirstForename.setAttribute("required", "true");
	frmScreen.cboTitle.setAttribute("required", "true");
	frmScreen.txtDateOfBirth.setAttribute("required", "true");
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
	frmScreen.cboGender.setAttribute("required", "true");
	<% /* DS : EP2_610
	frmScreen.cboSpecialNeeds.setAttribute("required", "true");
	*/ %>
	
	<% /* MAR19
	frmScreen.cboNationality.setAttribute("required", "true");
	*/ %>
	frmScreen.txtElectoralRoll.setAttribute("filter", "[0-9]");
	//START: MAR119 Code changed by Joyce Joseph on 07-Oct-05
	//NumberOfDependants only between 0-7
	frmScreen.txtNumberOfDependants.setAttribute("integer","true");
	frmScreen.txtNumberOfDependants.setAttribute("min","0");
	frmScreen.txtNumberOfDependants.setAttribute("max","7");
	//END: MAR119
	<% /* MAR19
	frmScreen.txtNumberOfDependants.setAttribute("required", "true");	
	*/ %>

	<% /* MAR20 Add filter for Time at Bank Years & Time at Months */ %>
	frmScreen.txtBankYears.setAttribute("filter", "[0-9]");
	frmScreen.txtBankMonths.setAttribute("filter", "[0-9]");
	<% /* Start: MAR95 - Added by Mahasen T on 10/10/2005 
				Month (Time at Bank) should not accept values > 11 */  %>
	frmScreen.txtBankMonths.setAttribute("integer", "true");
	frmScreen.txtBankMonths.setAttribute("min", "0");
	frmScreen.txtBankMonths.setAttribute("max", "11");
	<% /* End:   MAR95 */ %>
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtSurname.style.textTransform = "capitalize";
	frmScreen.txtFirstForename.style.textTransform = "capitalize";
	frmScreen.txtSecondForename.style.textTransform = "capitalize";
	frmScreen.txtOtherForenames.style.textTransform = "capitalize";
	// SW 15/06/2006 EP705 - Captialise other title
	frmScreen.txtOtherTitle.style.textTransform = "capitalize";
	
	//END: MAR36

	//START: MAR84 - New code added by Joyce Joseph on 27-Sep-2005
	// Disable Other Forenames on Screens
	frmScreen.txtOtherForenames.style.visibility="hidden";
	//END: MAR84
	
	<%/*MAH Ep2_346*/%>
	frmScreen.txtSurname.setAttribute("filter", "[a-zA-Z]");
	frmScreen.txtFirstForename.setAttribute("filter", "[a-zA-Z]");
	frmScreen.txtSecondForename.setAttribute("filter", "[a-zA-Z]");
	
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
	//START: MAR72 - New code added by Joyce Joseph on 27-Sep-2005
	if (CustomerXML.GetTagText("CUSTOMERSTATUS") == "2")
	{
		frmScreen.txtSecondForename.disabled=true;
	}
	//END: MAR72
}
</SCRIPT>