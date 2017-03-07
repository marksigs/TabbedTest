function getComputerName()
{
      var computerName = "";
      var omPC = new ActiveXObject("omPC.PCAttributesBO");
      if (omPC != null)
      {
            computerName = omPC.NameOfPC();
            omPC = null;
      }
      return computerName;
}

WScript.Echo(getComputerName());
