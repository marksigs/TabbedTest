strSrcFile = "C:\projects\DMS Control\Code\Black Box\AXWord_Viewer\AXWORD\TestScripts\KFI.rtf"
strTgtFile = "C:\projects\DMS Control\Code\Black Box\AXWord_Viewer\AXWORD\TestScripts\axword.kfi.rtf.compapi.xml"
strCompressionMethod = "COMPAPI"

set objAxword = CreateObject("axword.axwordclass")
if not objAxword is nothing then
	strFileContents = objAxword.ConvertFileToBase64(strSrcFile, strCompressionMethod)
	
	'msgbox strFileContents

	set xmlTgtDoc = CreateObject("Msxml2.DOMDocument")
	xmlTgtDoc.appendChild(xmlTgtDoc.createNode(1, "RESPONSE", ""))
	set root = xmlTgtDoc.documentElement
	root.setAttribute "TYPE", "SUCCESS"

	set newNode = xmlTgtDoc.createNode(1, "PRINTDOCUMENTDETAILS", "")
	root.appendChild(newNode)
	newNode.setAttribute "PRINTDOCUMENT", strFileContents

	xmlTgtDoc.save strTgtFile

	set xmlTgtDoc = Nothing
	set objAxword = Nothing
end if