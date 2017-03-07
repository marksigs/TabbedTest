var omPC = new ActiveXObject("omPC.PCAttributesBO");

function getComputerName()
{
	var computerName = "";
	if (omPC != null)
	{
		computerName = omPC.NameOfPC();
	}
	return computerName;
}

function getMACAddress()
{
	var macAddress = "";
	if (omPC != null)
	{
		macAddress = omPC.GetMACAddress();
	}
	return macAddress;
}

function getPrinters()
{
	var printers = "";
	if (omPC != null)
	{
		printers = omPC.FindLocalPrinterList("<REQUEST ACTION='CREATE'/>");
	}
	return printers;
}

function sleep(milliseconds)
{
	if (omPC != null)
	{
		omPC.Sleep(milliseconds);
	}
}


WScript.Echo("FindLocalPrinterList: " + getPrinters());
WScript.Echo("GetMACAddress: " + getMACAddress());
WScript.Echo("NameOfPC: " + getComputerName());
WScript.Echo("Sleeping for 5 seconds.");
sleep(5000);
WScript.Echo("Sleep over.");
