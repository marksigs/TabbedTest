<SCRIPT LANGUAGE="JavaScript">

function SetMasks() 
{
	frmScreen.txtLivingRooms.setAttribute("filter", "[0-9]");
	frmScreen.txtBedrooms.setAttribute("filter", "[0-9]");
	frmScreen.txtBathrooms.setAttribute("filter", "[0-9]");
	frmScreen.txtWCs.setAttribute("filter", "[0-9]");
	frmScreen.txtHabitableRooms.setAttribute("filter", "[0-9]");
	frmScreen.txtGarages.setAttribute("filter", "[0-9]");
	frmScreen.txtKitchens.setAttribute("filter", "[0-9]");
	frmScreen.txtParkingSpaces.setAttribute("filter", "[0-9]");
	//EP2_2 Add extra field
	frmScreen.txtFloors.setAttribute("filter", "[0-9]");
}

</SCRIPT>