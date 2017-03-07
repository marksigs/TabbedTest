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
<% /* On Ok - output a message to input full 3 year address history if the following conditions are met:

CustomerAddress.AddressType validation type of "H" (Home/Current in combogroup 'customeraddresstype') does not exist OR CustomerAddress.AddressType Home/Current does exist with CustomerAddress.DateMovedIn less than 3 years ago and Address with CustomerAddress.Address Type = "P" (Previous in combogroup 'customeraddresstype') does not exist  OR CustomerAddress.Address Type = "P" does exist and the earliest occurence (i.e. the one with the oldest DateMovedIn) has CustomerAddress.Datemovedin < 3 years ago */ %>

<% /* TW 2/1/2003 Requirement changed - Check Home/Current for EACH applicant, not for ANY. In addition, where we are validating for 3 years addresses, this should be done for each applicant too, rather than looping through all recrds as a whole. (Otherwise App1 has 1 current < 3 years, App2 has 1 previous > 3 years, and the screen passes) */ %>

var aValueMDY = null;
var addressNodesXML = m_XMLAddressSummary; var bHomeExists = false; var bPreviousExists = false; var dHomeResidentFrom = new Date(); var dPreviousResidentFrom = new Date(); var dThisDate = null; var dCutoffDate = new Date(); var vCustomerName = "";

	if(document.activeElement.id != "btnSubmit")
		{
		return 0;
		}

	dCutoffDate = new Date(dCutoffDate.getFullYear() - 3, dCutoffDate.getMonth(), dCutoffDate.getDate());
	addressNodesXML.ActiveTag = null;
	addressNodesXML.CreateTagList("CUSTOMERADDRESS");

// TW 2/1/2003
	var vCustomerNumber = new Array();
	var x = 0;
	var y = 0;
	for (var cLoop = 0; cLoop < addressNodesXML.ActiveTagList.length; cLoop++)
	{
		addressNodesXML.SelectTagListItem(cLoop);
		for (x = 0; x < vCustomerNumber.length; x++)
		{
			if (addressNodesXML.GetTagText("CUSTOMERNUMBER") == vCustomerNumber[x])
				{
				break;
				}
		}
		if (x == vCustomerNumber.length)
			{
			vCustomerNumber[x] = addressNodesXML.GetTagText("CUSTOMERNUMBER");
			}
	}
	
// TW 2/1/2003 End

// TW 2/1/2003 Checks now made for EACH customer rather than ANY

	for (x = 0; x < vCustomerNumber.length; x++)
	{
		bHomeExists = false;
		bPreviousExists = false;
		dPreviousResidentFrom = new Date();
		for (var iLoop = 0; iLoop < addressNodesXML.ActiveTagList.length; iLoop++)
		{				
			addressNodesXML.SelectTagListItem(iLoop);
			if (addressNodesXML.GetTagText("CUSTOMERNUMBER") == vCustomerNumber[x])
			{
				addressNodesXML.SelectTagListItem(iLoop);
				aValueMDY = addressNodesXML.GetTagText("DATEMOVEDIN").split("/",3);
				dThisDate = new Date(aValueMDY[2], aValueMDY[1] - 1, aValueMDY[0]);
				switch (addressNodesXML.GetTagAttribute("ADDRESSTYPE","TEXT"))
				{
					//alert(addressNodesXML.GetTagAttribute("ADDRESSTYPE","TEXT"));
					case "Home/Current":
						bHomeExists = true;
						dHomeResidentFrom = dThisDate;
						break;
					case "Previous":
						bPreviousExists = true;
						if (dThisDate < dPreviousResidentFrom)
							{
							dPreviousResidentFrom = dThisDate;
							}
						break;
				}
			}		
		}
	
//		if ((!bHomeExists || (bHomeExists && dHomeResidentFrom > dCutoffDate)) && (!bPreviousExists || (bPreviousExists && dPreviousResidentFrom > dCutoffDate)))
		if (!bHomeExists || isNaN(dHomeResidentFrom) || (bHomeExists && dHomeResidentFrom > dCutoffDate && (!bPreviousExists || isNaN(dPreviousResidentFrom) || (bPreviousExists && dPreviousResidentFrom > dCutoffDate))))
			{
			for (y = 1; y <= vCustomerNumber.length; y++)
				{
				if(vCustomerNumber[x] == scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + y, null))
					{
					vCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + y, null);
					break;
					}
				}

			return (scScreenFunctions.DisplayClientError("Please enter full 3 year address history for " + vCustomerName, "images/MSGBOX01.ICO"));
			}
	}
        return 0;
}


<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
}

</SCRIPT>
