<script language="JScript">
function DirectorySearch(XML, AttributeArray, sPanelID, sName, sTown, sFilter, sType, sSortCode,
			sPostcode, sBuildingName, sBuildingNumber, sFlatNumber, sStreet)
<%/* Conducts a ThirdPartyBO.FindDirectoryList given the passed parameters
	 Returns the number of records found.
*/%>

{
	function IsStringEmpty(strString)
	// Ascertains whether the passed string is empty or contains only a wildcard '*' (for the purposes of a PAF search)
	{
		var bStringIsEmpty = false;
		
	//BMIDS00150 - DPF 3/7/02 - the following if statement has been inserted to cope with null values being passed
	//line taken from Core version.
		if (strString != null)
		{		
			var ssArray = strString.split(" ");
			var nLength = ssArray.length;
			var nIndex = 0;


			while(ssArray[nIndex] == "" && nIndex < nLength)
				nIndex++;

			if ((ssArray[nIndex] == null) || (ssArray[nIndex] == "*"))
				bStringIsEmpty = true;
		}

		return(bStringIsEmpty);
	}

	var nNumFound = 0;

	<% /* SYS4438 - PanelID can be used on its own */ %>
	<% /* if (IsStringEmpty(sName) && IsStringEmpty(sSortCode)) */ %>
	<% /* ASu BMIDS00601 - Added Town for wildcard searches */ %>
	if (IsStringEmpty(sPanelID) && IsStringEmpty(sName) && IsStringEmpty(sSortCode)&& IsStringEmpty(sTown))
		return(-1);
	<% /* SYS4438 - End */ %>

	with (XML)
	{			
		CreateRequestTagFromArray(AttributeArray,null)
		CreateActiveTag("NAMEANDADDRESSDIRECTORY");
		CreateTag("PANELID", sPanelID);
		CreateTag("COMPANYNAME", sName);
		CreateTag("TOWN", sTown);
		CreateTag("STREET", sStreet);
		CreateTag("POSTCODE", sPostcode);
		CreateTag("BUILDINGORHOUSENAME", sBuildingName);
		CreateTag("BUILDINGORHOUSENUMBER", sBuildingNumber);
		CreateTag("FLATNUMBER", sFlatNumber);
		CreateTag("FILTER", sFilter);
		CreateTag("NAMEANDADDRESSTYPE", sType);
		CreateTag("SORTCODE", sSortCode);
		//BMIDS01039
		//XML.RunASP(document,"FindDirectoryList.asp");
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document,"FindDirectoryList.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
		
	}

	// Check the response from the ThirdPartyBO
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
			
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0])) 
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			XML.ActiveTag = null;
			var tagList = XML.CreateTagList("NAMEANDADDRESSDIRECTORY");
			var nNumFound = tagList.length;
		}

	ErrorTypes = null;
	ErrorReturn = null;				
	XML = null;

	return(nNumFound);
}
</script>
