var obj = new ActiveXObject("DmsCompression.dmsCompression1");
var strOriginal = "123456789123456789123456789123456789123456789123456789";
WScript.Echo(strOriginal);
var strCompressed = obj.BSTRCompressToBSTR(strOriginal);
var strDecompressed = obj.BSTRDecompressToBSTR(strCompressed);

WScript.Echo(strDecompressed);