<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      OverrideAttribs.asp
Copyright:     Copyright © 2004 Marlborough Stirling

Description:   Override Freeze data for certain tasks attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DRC     03/02/2004  Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:


DRC		03/02/04	BMIDS692 Data freeze overrides
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
*/
<% /* Created by DRC  Feb 2004  

	For customisation follow the exmples below
	Note that for each new override task, there must be an entry in the switch statement below
	in the ClientFreezeOverRide function and in the ClearAllOverrides function.
	
	The positions for insertion of extra overrides is indicated by the sample "AnOtherOveride" task shown (an commented out

*/ %>

function ClientFreezeOverRide(sTaskID, strFreezeDataInd)
<% /* Specify client code for freeze overrides */ %>

{
  var  bOverride = false;
  var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
 	
  <% /* Clear out any existing overrides first - overides are not cumulative;
   only the last completed override task will apply   */ %>  
  ClearAllOverrides();
  
  if (strFreezeDataInd=="1" )
  {
	switch (sTaskID)
	{
     <% /* If Offer is issued or reissued or transfer of equity sent then all overides are set to 0 (false)    */ %>        
          
		case XML.GetGlobalParameterString(document, "TMFreezeOveride_Shortfall"):
		   scScreenFunctions.SetContextParameter(window,"idFreezeOveride_Shortfall", "1");
		   bOverride = true; 
		   break;
<% /* Further Freeze overrides go in here
		case XML.GetGlobalParameterString(document, "TM_AnOtherOveride"):
		scScreenFunctions.SetContextParameter(window,"idFreezeOveride_AnOtherOveride", "1");
		   bOverride = true; 
		   break;
*/ %>		
  
		default: bOverride = false; break;
	}
  }
  
  XML = null;
  return bOverride;
  
}
function ClearAllOverrides()
{
	scScreenFunctions.SetContextParameter(window,"idFreezeOveride_Shortfall", "0");
	<% /* Further Freeze overrides go in here
		
		scScreenFunctions.SetContextParameter(window,"idFreezeOveride_AnOtherOveride", "0");
		   
	*/ %>
}

</SCRIPT>
