function getMACAddress()
{
      var macAddress = "";
      var omPC = new ActiveXObject("omPC.PCAttributesBO");
      if (omPC != null)
      {
            macAddress = omPC.GetMACAddress();
            omPC = null;
      }
      return macAddress;
}

WScript.Echo(getMACAddress());
