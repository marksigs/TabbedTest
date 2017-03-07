set xmlSrcDoc = CreateObject("Msxml2.DOMDocument")
xmlSrcDoc.async = false

'xmlSrcDoc.load("axword.xml")
'xmlSrcDoc.load("axword.twopagedocument.xml")
'xmlSrcDoc.load("axword.notrackedchanges.xml")
'xmlSrcDoc.load("axword.trackedchanges.xml")

'xmlSrcDoc.load("axword.kfi.doc.xml")
'strCompressionMethod = ""
'strDeliveryType = "10"
'strExtn = ".doc"

'xmlSrcDoc.load("axword.kfi.pdf.xml")
'strCompressionMethod = ""
'strDeliveryType = "20"
'strExtn = ".pdf"

xmlSrcDoc.load("axword.kfi.rtf.xml")
strCompressionMethod = ""
strDeliveryType = "30"
'strExtn = ".rtf"

'xmlSrcDoc.load("axword.kfi.doc.zlib.xml")
'strCompressionMethod = "ZLIB"
'strDeliveryType = "10"

'xmlSrcDoc.load("axword.kfi.pdf.zlib.xml")
'strCompressionMethod = "ZLIB"
'strDeliveryType = "20"

'xmlSrcDoc.load("axword.kfi.rtf.zlib.xml")
'strCompressionMethod = "ZLIB"
'strDeliveryType = "30"

set xmlTgtDoc = CreateObject("Msxml2.DOMDocument")
xmlTgtDoc.appendChild(xmlTgtDoc.createNode(1, "REQUEST", ""))
set root = xmlTgtDoc.documentElement

if len(strCompressionMethod) > 0 then
	root.setAttribute "COMPRESSIONMETHOD", strCompressionMethod
end if

if len(strExtn) > 0 then
	root.setAttribute "FILEEXTENSION", strExtn
end if

set newNode = xmlTgtDoc.createNode(1, "CONTROLDATA", "")
root.appendChild(newNode)
newNode.setAttribute "DOCUMENTID", "10"
newNode.setAttribute "DELIVERYTYPE", strDeliveryType
newNode.setAttribute "PRINTER", "\\PRINTSERV1\JH_G_HPLJ4050_1"
newNode.setAttribute "COPIES", "1"
newNode.setAttribute "FIRSTPAGEPRINTERTRAY", "0"
newNode.setAttribute "OTHERPAGESPRINTERTRAY", "0"

set newNode = xmlTgtDoc.createNode(1, "DOCUMENTCONTENTS", "")
root.appendChild(newNode)
set nodePrintDocumentDetails = xmlSrcDoc.selectSingleNode("//RESPONSE/PRINTDOCUMENTDETAILS")
strFileContents = nodePrintDocumentDetails.getAttribute("PRINTDOCUMENT")

newNode.setAttribute "FILECONTENTS", strFileContents

strRequest = xmlTgtDoc.xml

set objAxword = CreateObject("axword.axwordclass")

'defaults to false
objAxword.ViewAsWord = true
objAxword.ViewAsPDF = false
objAxword.ReadOnly = false
objAxword.ResizeableFrame = true
objAxword.ShowFindFreeText = false
'defaults to true
objAxword.SpellCheckOnSave = true
'defaults to true
objAxword.SpellCheckWhileEditing = false
objAxword.ShowCommandBars = false
'defaults to false
objAxword.PersistState = true
objAxword.ShowPrint = true
objAxword.ShowPrintDialog = true
objAxword.ShowTrackedChanges = false
'objAxword.DisablePrintOut = true

'strResponse = objAxword.DisplayDocument(strRequest)
strResponse = objAxword.DisplayDocumentNative(strRequest)
'objAxword.PrintDocument strRequest
'strPrinters = objAxword.GetPrintersAsXML
set xmlPrintersDoc = CreateObject("Msxml2.DOMDocument")
call xmlPrintersDoc.loadXML(strPrinters)
call xmlPrintersDoc.save("C:\projects\DMS Control\Code\Black Box\AXWord_Viewer\AXWORD\TestScripts\Printers.xml")

'msgbox Left(strFileContents, 64) 
'msgbox Left(strResponse, 64)

if Len(strResponse) > 0 then
	if strDeliveryType = "10" then
		strExtn = ".doc"
	elseif strDeliveryType = "20" then
		strExtn = ".pdf"
	elseif strDeliveryType = "30" then
		strExtn = ".rtf"
	end if
	Call objAxword.ConvertBase64ToFile(strResponse, "C:\projects\DMS Control\Code\Black Box\AXWord_Viewer\AXWORD\TestScripts\KFI.new" & strExtn, strCompressionMethod)
end if

if objAxword.DocumentEdited then
	msgbox "Document has been edited"
end if
