<%
/*
Workfile:      PAFSearch.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Contains shared functions for PAFSearch
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
LD		23/05/02	SYS4727 Use cached versions of frame functions
DB		29/05/02	SYS4767 MSMS to Core Integration
SG		05/06/02	SYS4818 Error in SYS4767
INR		14/06/02	SYS4803 PAF search not working after Core Integration
GD		26/06/02	BMIDS0077 Merge in SG SYS4930
DRC     23/09/02    BMIDS00436 Fixed the Flat No retrieval
INR		08/11/02	BMIDS00278 Removed debug message
SAB		23/03/2006	EP272 Rolled back from MARS version to allow local postcode validation
SAB		23/03/2006	EP287 Fixed bug preventing detection of the selected country from working
						  correctly.
SAB		24/03/2006	EP288 Only performs the country validation when a country is selected
PB		16/05/2006	EP529 - MAR1556 Limit address returns to Globally set parameter length
PB		26/05/2006	EP622 - PAF broken when multiple addresses returned, since EP529 (Merge with MAR1556)
PB		28/06/2006	EP707 - Set country to UK
*/
%>

<script language="JScript">
function PAFSearch(ctrPostcode,ctrHouseName,ctrHouseNumber,ctrFlatNumber,ctrStreet,
				ctrDistrict,ctrTown,ctrCounty,ctrCountry)
{		
		<% /* PB EP529 / MAR1556 Declare Global Parameters for max address character string length */ %>
		var nWSMaxFlatName = 0;
		var nWSMaxHouseNumber = 0;
		var nWSMaxHouseName = 0;
		var nWSMaxStreet = 0;
		var nWSMaxDistrict = 0;
		var nWSMaxTown = 0;
		var nWSMaxCounty = 0;
		var globalXML;
		
	function IsStringEmpty(strString)
	// Ascertains whether the passed string is empty or contains only a wildcard '*' (for the purposes of a PAF search)
	{
		var bStringIsEmpty = false;
		var ssArray = strString.split(" ");
		var nLength = ssArray.length;
		var nIndex = 0;

		while(ssArray[nIndex] == "" && nIndex < nLength)
			nIndex++;

		if ((ssArray[nIndex] == null) || (ssArray[nIndex] == "*"))
			bStringIsEmpty = true;

		return(bStringIsEmpty);
	}

	var selIndex = ctrCountry.selectedIndex;

	<% /* EP287 & EP288 - <SELECT> = 0 */ %>
	if(selIndex > 0 && scScreenFunctions.IsOptionValidationType(ctrCountry,selIndex,"N"))
		alert("Postal Address File search is not allowed for the selected country");
	else
	{
		//GD BMIDS0077
		<%/*SG 25/06/02 SYS4930 START */%>
		<%/* SG 05/06/02 SYS4818 */%>	
		<%/* DB SYS4767 - MSMS Integration */%>
		//var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<%/* var XML = new scXMLFunctions.XMLObject(); */%>
		<%/* DB End */%>	
		<%/* SG 05/06/02 SYS4818 */%>
		if (m_bIsPopup)
		{
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			//globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); //MAR1556
			globalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject(); //MAR1556
		}
		else
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); //MAR1556
		}
		<%/*SG 25/06/02 SYS4930 END */%>
		var bHasData = false;
		var bPickList = true;
		
		<% /* PB EP529 / MAR1556 Populate Global Parameters for max address character string length */ %>
		nWSMaxFlatName = globalXML.GetGlobalParameterAmount(document, "WSMaxFlatName");
		nWSMaxHouseNumber = globalXML.GetGlobalParameterAmount(document, "WSMaxHouseNumber");
		nWSMaxHouseName = globalXML.GetGlobalParameterAmount(document, "WSMaxHouseName");
		nWSMaxStreet = globalXML.GetGlobalParameterAmount(document, "WSMaxStreet");
		nWSMaxDistrict = globalXML.GetGlobalParameterAmount(document, "WSMaxDistrict");
		nWSMaxTown = globalXML.GetGlobalParameterAmount(document, "WSMaxTown");
		nWSMaxCounty = globalXML.GetGlobalParameterAmount(document, "WSMaxCounty");	
		<% /* EP529 / MAR1556 End */ %>

		XML.CreateActiveTag("REQUEST");
		//DB SYS4767 - MSMS Integration
		XML.CreateActiveTag("ADDRESS");
		
		// Record the search criteria
		if (ctrPostcode != null) if (bHasData = bHasData | !IsStringEmpty(ctrPostcode.value)) XML.SetAttribute("POSTCODE",ctrPostcode.value);
		if (ctrHouseName != null) if (bHasData = bHasData | !IsStringEmpty(ctrHouseName.value)) XML.SetAttribute("NAME",ctrHouseName.value);
		if (ctrHouseNumber != null)  XML.SetAttribute("HOUSENUMBER",ctrHouseNumber.value);
		if (ctrFlatNumber != null)  XML.SetAttribute("FLATNUMBER",ctrFlatNumber.value);
		if (ctrStreet != null) if (bHasData = bHasData | !IsStringEmpty(ctrStreet.value)) XML.SetAttribute("ADDRESSLINE1",ctrStreet.value);
		if (ctrDistrict != null)  XML.SetAttribute("ADDRESSLINE2",ctrDistrict.value);
		if (ctrTown != null)  XML.SetAttribute("ADDRESSLINE3",ctrTown.value);
		if (ctrCounty != null)  XML.SetAttribute("ADDRESSLINE4",ctrCounty.value);
		//if (ctrCountry != null) if (!IsStringEmpty(ctrCountry.value)) XML.CreateTag("COUNTRY",ctrCountry.value);
		//the following attribute needs to be at the end, otherwise
		//Validation of omPAF.PAFBO.validation will need to be changed
		XML.SetAttribute("PICKLIST","Y");
		
		if(!bHasData)
		{
			alert("No valid PAF search criteria entered");
			return(false);
		}

			while(bPickList) 
		{
		//DB End
		
		XML.RunASP(document,"PAFSearch.asp");
		
		if(!XML.IsResponseOK()) 
			return(false);

			XML.ActiveTag = null;
		
			var XMLActiveTag = XML.SelectSingleNode("RESPONSE/ADDRESSDATA");
		
		//DB SYS4767 - MSMS Integration
		// Analyse the PAF search results
			switch(XML.GetAttribute("QASRETURN"))
			{
				case "NOMATCH":
					bPickList = false;
					alert("No address found for the details provided");
					break;
				case "PICKLIST":
					var strKey;
					
					// More than one match, so display the popup
					var sAddress = scScreenFunctions.DisplayPopup(window, document, "za010.asp", XML.XMLDocument.xml , 630, 280);
					if(sAddress == null)
					{
						bPickList = false;
						break;
					}
					else
					{	
						//GD BMIDS0077
						//var addressXML = new scXMLFunctions.XMLObject();
						<%/*SG 25/06/02 SYS4930 START */%>		
						<%/* INR 20/06/02 SYS4803 changed as per SYS4727 */%>
						//var addressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						if (m_bIsPopup)
							var addressXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
						else
							var addressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						<%/*SG 25/06/02 SYS4930 END */%>
						var xmlReturnedItem = new ActiveXObject("microsoft.XMLDOM");
						xmlReturnedItem.loadXML(sAddress);
						
						addressXML.AddXMLBlock(xmlReturnedItem);
						<% /* PB 16/05/2006 EP529 / MAR1556 */ %>
						//AddressNodeList = addressXML.XMLDocument.selectNodes("ValidateAddress/Address");
						//XML = AddressNodeList.item(0);
						//var houseNumberNode = XML.selectSingleNode("HouseOrBuildingNumber");
						//var houseNumber = houseNumberNode.text ;
						<% /* EP529 / MAR1556 End */ %>
						addressXML.SelectTag(null,"ITEM");
						strKey = addressXML.GetAttribute("KEY");
					}

					//Recreate the request for FindAddressWithKey
					XML.SelectTag(null,"RESPONSE");
					XML.RemoveActiveTag();
					
					XML.CreateActiveTag("REQUEST");
					XML.CreateActiveTag("ADDRESSKEY");

					XML.SetAttribute("KEY",strKey);
					XML.SetAttribute("PICKLIST","Y");

					break;
					
				case "ADDRESS":
					
					//Populate Screen
					bPickList = false;
					if (ctrPostcode != null) ctrPostcode.value = XML.GetAttribute("POSTCODE");
					
					<% /* EP529 / MAR1556
					if (ctrHouseName != null) ctrHouseName.value = XML.GetAttribute("BUILDINGORHOUSENAME"); */ %>
					if (ctrHouseName != null)
					{
						var tempHouseName = XML.GetAttribute("BUILDINGORHOUSENAME");
						ctrHouseName.value = tempHouseName.substring(0,nWSMaxHouseName);
					}
					
					//var HouseName = XML.selectSingleNode("HouseOrBuildingName").text;
					//if (ctrHouseName != null && HouseName != null)
					//{
					//	HouseName = HouseName.substring(0,nWSMaxHouseName) ;
					//	ctrHouseName.value = HouseName ;
					//}
					<% /* EP529 / MAR1556 End */ %>
					var FullHouseNumber = "";
					var SubBuildingName = XML.GetAttribute("SUBBUILDINGNAME");
					var BuildingNumber = XML.GetAttribute("BUILDINGORHOUSENUMBER");
					// BMIDS00436 - DRC SubBuildingName gives the flat number
					// if (ctrHouseNumber != null) ctrHouseNumber.value = BuildingNumber + SubBuildingName;
					
					<% /* EP529 / MAR1556
					if (ctrHouseNumber != null) ctrHouseNumber.value = BuildingNumber */ %>
					if (ctrHouseNumber != null) ctrHouseNumber.value = BuildingNumber.substring(0,nWSMaxHouseNumber);
					//if (ctrHouseNumber != null && BuildingNumber != null) 
					//{
					//	BuildingNumber = BuildingNumber.substring(0,nWSMaxHouseNumber) ;
					//	ctrHouseNumber.value = BuildingNumber;
					//}
					<% /* EP529 / MAR1556 End */ %>
					
					//if (ctrFlatNumber != null) 
					//  ctrFlatNumber.value = XML.GetAttribute("FLATNUMBER");
					var FlatNoPos = 0;
					<% /* EP529 / MAR1556
					var FlatNoLength = SubBuildingName.length;
					*/ %>
					
                    FlatNoPos = SubBuildingName.indexOf("Flat");
                    // Chop off the "Flat" at the start
                    if (FlatNoPos == 0)
					{
						<% /* EP529 / MAR1556
						ctrFlatNumber.value = SubBuildingName.substring(FlatNoPos + 5, FlatNoLength); */ %>
						ctrFlatNumber.value = SubBuildingName.substring(FlatNoPos + 5, nWSMaxFlatName + 5);
						<% /* EP529 / MAR1556 End */ %>
					}
                    else
					{
						<% /* EP529 / MAR1556
						ctrFlatNumber.value = SubBuildingName; */ %>
						ctrFlatNumber.value = SubBuildingName.substring(0,nWSMaxFlatName);
						<% /* EP529 / MAR1556 End */ %>
					}	
					// BMIDS00436 - End  
					
					var FullFareName = "";
					var DepFareName = XML.GetAttribute("DEPENDENTTHOROUGHFARENAME");
					var FareName = XML.GetAttribute("THOROUGHFARENAME");
					
					if(DepFareName!="" && FareName!="")
					{
						FullFareName = DepFareName + ", " + FareName;
					}
					else
					{
						FullFareName = DepFareName + FareName;
					}

					<% /* EP529 / MAR1556
					if (ctrStreet != null) ctrStreet.value = FullFareName;
					if (ctrDistrict != null) ctrDistrict.value = XML.GetAttribute("DEPENDENTLOCALITY");
					if (ctrTown != null) ctrTown.value = XML.GetAttribute("POSTTOWN");
					if (ctrCounty != null) ctrCounty.value = XML.GetAttribute("COUNTY"); */ %>
					if (ctrStreet != null) ctrStreet.value = FullFareName.substring(0,nWSMaxStreet);
					if (ctrDistrict != null) ctrDistrict.value = XML.GetAttribute("DEPENDENTLOCALITY").substring(0,nWSMaxDistrict);
					if (ctrTown != null) ctrTown.value = XML.GetAttribute("POSTTOWN").substring(0,nWSMaxTown);
					if (ctrCounty != null) ctrCounty.value = XML.GetAttribute("COUNTY").substring(0, nWSMaxCounty);
					<% /* PB 28/06/2006 EP707 Begin */ %>
					for(var intLoop=0;intLoop<ctrCountry.childNodes.length-1;intLoop++)
					{
						if(ctrCountry.childNodes(intLoop).innerText=='United Kingdom'){
							ctrCountry.selectedIndex=intLoop;
							break;
						}
					}
					<% /* EP707 End */ %>
					//var Street = XML.selectSingleNode("Street").text;
					//if (ctrStreet != null && Street != null)
					//	ctrStreet.value = Street.substring(0,nWSMaxStreet) ;
					//var District = XML.selectSingleNode("District").text;
					//if (ctrDistrict != null && District != null)
					//	ctrDistrict.value = District.substring(0,nWSMaxDistrict) ;
					//var Town = XML.selectSingleNode("TownOrCity").text;
					//if (ctrTown != null && Town != null)
					//	ctrTown.value = Town.substring(0,nWSMaxTown) ;
					//var County = XML.selectSingleNode("County").text;	
					//if (ctrCounty != null && County != null)
					//	ctrCounty.value = County.substring(0, nWSMaxCounty) ;
					<% /* EP529 / MAR1556 End */ %>

					break;
			}
		}
		//DB End
		return(true);
	}
}
</script>
