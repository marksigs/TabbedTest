function getPrinters()
{
      var printers = "";
      var omPC = new ActiveXObject("omPC.PCAttributesBO");
      if (omPC != null)
      {
            printers = omPC.FindLocalPrinterList("<REQUEST ACTION='CREATE'/>");
            omPC = null;
      }
      return printers;
}

WScript.Echo(getPrinters());
