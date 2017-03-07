function sleep(milliseconds)
{
      var omPC = new ActiveXObject("omPC.PCAttributesBO");
      if (omPC != null)
      {
            omPC.Sleep(milliseconds);
            omPC = null;
      }
}

WScript.Echo("Sleeping for 10 seconds.");
sleep(10000);
WScript.Echo("Sleep over.");