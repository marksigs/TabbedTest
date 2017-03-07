/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:			$Workfile: Documents.js $
Copyright:			Copyright © 2005 Marlborough Stirling
Description:		Document production and archive common functions.
Current Version:	$Revision: 1 $
Last Modified:		$JustDate: 19/04/05 $
Modified By:		$Author: Astanley $
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
GHun	14/07/2005	MAR7	Integrate local printing
JD		09/12/2005	MAR838	Added method getDocumentFromFileNet called from DMS105 and DMS107
PSC		16/01/2006	MAR1054 Add filenet userid as a parameter to getDocumentFromFileNet
PSC		06/02/2006	MAR1197	Amend to use file extension as well as delivery type 
PSC		13/03/2006	MAR1337	Amend printDocument to set showProgressBar to false
IK		25/04/2006	EP462	amend for omGemini interface
AS		04/08/2006	EP1068	omGemini: Add support for scanned documents
AS		09/08/2006	EP1078	omGemini: viewing an edited document in DMS107 does not return the document
AS		12/12/2006	EP1262	Gemini Printing.
AS		15/12/2006	EP1269	Gemini Printing: second release to system test.
AS		16/01/2007	EP1288	Gemini Printing: Replaced error checking in getArchivedDocument.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

// AS 15/12/2006 EP1269 Gemini Printing: second release to system test.

var GEMINIPRINTSTATUS_AWAITINGAPPROVAL = 10;
var GEMINIPRINTSTATUS_NOTAPPROVED = 20;
var GEMINIPRINTSTATUS_APPROVED = 30;
var GEMINIPRINTSTATUS_GEMINIPRINTED = 40;

function getLocalPrinters(xmlLocalPrinters)
{		
	// AS 02/02/05 BBGRETAIL Use Axword - no need for omPC.
	if (xmlLocalPrinters == null)
	{
		// Only call Axword if do not already have the local printers.
		xmlLocalPrinters = new ActiveXObject("Microsoft.XMLDOM");
		if (xmlLocalPrinters != null)
		{
			var omPC = new ActiveXObject("omPC.PCAttributesBO");
			if (omPC != null)
			{
				xmlLocalPrinters.loadXML(omPC.FindLocalPrinterList("<?xml version='1.0'?><REQUEST ACTION='CREATE'></REQUEST>"));
				omPC = null;
			}
		}
	}
	
	return xmlLocalPrinters;
}

function printDocumentData(success, fileContents, fileSize, fileEdited, isLastErr, lastErrNumber, lastErrSource, lastErrDescription)
{
	var _success = success;
	var _fileContents = fileContents;
	var _fileSize = fileSize;
	var _fileEdited = fileEdited;
	var _isLastErr = isLastErr;
	var _lastErrNumber = lastErrNumber;
	var _lastErrSource = lastErrSource;
	var _lastErrDescription = lastErrDescription;
	
	this.get_success = get_success;
	this.get_fileContents = get_fileContents;
	this.get_fileSize = get_fileSize;
	this.get_fileEdited = get_fileEdited;
	this.get_isLastErr = get_isLastErr;
	this.get_lastErrNumber = get_lastErrNumber;
	this.get_lastErrSource = get_lastErrSource;
	this.get_lastErrDescription = get_lastErrDescription;
	
	function get_success() { return _success; }
	function get_fileContents() { return _fileContents; }
	function get_fileSize() { return _fileSize; }
	function get_fileEdited() { return _fileEdited; }
	function get_isLastErr() { return _isLastErr; }
	function get_lastErrNumber() { return _lastErrNumber; }
	function get_lastErrSource() { return _lastErrSource; }
	function get_lastErrDescription() { return _lastErrDescription; }
}

function printDocument(fileContents, documentId, documentTitle, deliveryType, compressionMethod, printerType, copies, showPrintDialog, showProgressBar, readOnly, printOnly, fileExtension, geminiPrintMode)
{		
	var printDocumentDataObj = null;

	var ax = new ActiveXObject("axword.axwordclass");

	if (ax != null)
	{
		// The following properties control the apppearance of the Axword window. 
		// These properties will be ignored if printerType = "W" (Workstation Printing)
		// as this does not use the Axword word, but displays documents in the native application, 
		// e.g., Word/Acrobat.
		ax.ViewAsWord = true;
		ax.ViewAsPDF = false;
		ax.ResizeableFrame = true;
		ax.PersistState = true;
		ax.ShowFindFreeText = false;
		ax.SpellCheckOnSave = false;
		ax.SpellCheckWhileEditing = true;
		ax.ShowCommandBars = false;
		ax.ShowPrint = false;
		ax.ShowPrintDialog = showPrintDialog;
		ax.ShowTrackedChanges = false;
		ax.DisablePrintOut = false;

		ax.ReadOnly = readOnly;
		showProgressBar = false;

		var xmlAxwordRequest = new ActiveXObject("Microsoft.XMLDOM");
		var xmlRequestElement = xmlAxwordRequest.createElement("REQUEST");
		xmlAxwordRequest.appendChild(xmlRequestElement);
		xmlElement = xmlAxwordRequest.createElement("CONTROLDATA");
		xmlRequestElement.appendChild(xmlElement);
		if (documentId != null && documentId != "") xmlElement.setAttribute("DOCUMENTID", documentId);
		if (documentTitle != null && documentTitle != "") xmlElement.setAttribute("DOCUMENTTITLE", documentTitle);
		
		if (fileExtension != null && fileExtension != "")
			xmlRequestElement.setAttribute("FILEEXTENSION", fileExtension);
		else
			if (deliveryType != null && deliveryType != "") xmlElement.setAttribute("DELIVERYTYPE", deliveryType);
		
		if (compressionMethod != null && compressionMethod != "") xmlElement.setAttribute("COMPRESSIONMETHOD", compressionMethod);
		if (copies != null && copies != "") xmlElement.setAttribute("COPIES", copies);
		
		xmlElement = xmlAxwordRequest.createElement("DOCUMENTCONTENTS");
		xmlRequestElement.appendChild(xmlElement);					
		xmlElement.setAttribute("FILECONTENTS", fileContents);

		var success = false;
		var fileSize = 0;
		var fileEdited = false;
					
		// This flag could be configured as a global parameter on the database.
		// For now, use a hard coded value.
		var displayDocumentNative = true;
		
		if (printerType == "L" || printerType == "R")
		{
			if (displayDocumentNative)
			{
				// Document will be displayed in native application (Word/Acrobat). 
				// User can choose to print from the native application.
				// The document will also always be printed from the central print server.
				// For "L", this will be to the printer the user has choosen on PM010.
				// For "R", this will be to the remote printer held against the template record.
				fileContents = ax.DisplayDocumentNative(xmlAxwordRequest.xml);
			}
			else
			{					
				// Traditional Omiga4 printing. Document is displayed in Axword window
				// wrapping Word/Acrobat. Document is not printed from Axword, but from central
				// print server.
				fileContents = ax.DisplayDocument(xmlAxwordRequest.xml);
			}					
			success = fileContents != null && fileContents != "";
		}
		else
		{
			if (printOnly)
			{
				// Print type is "W", and not being viewed or edited.
				
				// AS 12/12/2006 EP1262 Gemini Printing.
				if (geminiPrintMode == "20" || geminiPrintMode == "30")
				{
					// GeminiPrintMode is Immediate or On-Hold, so do nothing.
					success = true;
				}
				else
				{
					// Document is to be printed, so display Axword print dialog.
					ax.ShowPrintDialog = showPrintDialog;
					ax.ShowProgressBar = showProgressBar;
					// Get Axword to print document. 
					// For DOC and RTF, document will be printed via Word on the workstation.
					// For PDF, document will be printed via wpCubed printing software.
					success = ax.PrintDocument(xmlAxwordRequest.xml);
				}
			}
			else
			{
				if (displayDocumentNative)
				{
					// Document is to be view/edited in native application, e.g., Word/Acrobat. 
					// This also allows the document to be printed from the native application. 
					// This applies to Workstation Printing, Email, File, and Fax.
					fileContents = ax.DisplayDocumentNative(xmlAxwordRequest.xml);
				}
				else
				{
					// Document is to be viewed/edited in Axword wrapper around Word/Acrobat.
					// If Workstation Printing, then show the Print button on the Axword window,
					// so we can print directly from Axword to a workstation printer.
					ax.ShowPrint = printerType == "W";
					if (ax.ShowPrint)
					{
						ax.ShowPrintDialog = showPrintDialog;
						ax.ShowProgressBar = showProgressBar;
					}
					fileContents = ax.DisplayDocument(xmlAxwordRequest.xml);							
				}
				success = fileContents != null && fileContents != "";
			}
		}
		fileSize = ax.FileSize;
		fileEdited = ax.FileSaved || ax.DocumentEdited;

		printDocumentDataObj = new printDocumentData(success, fileContents, fileSize, fileEdited, ax.IsLastErr(), ax.LastErrNumber, ax.LastErrSource, ax.LastErrDescription);
	}
	
	return printDocumentDataObj;
}

function getArchivedDocument(
	window, userId, unitId, machineId, channelId, userAuthorityLevel, applicationNumber, 
	documentGuid, fileGuid, version, readOnly, printOnly) 
{
	var xmlRequestObj = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
	var xmlRequest = xmlRequestDocument.createElement("REQUEST");
	xmlRequestDocument.appendChild(xmlRequest);
		
	xmlRequest.setAttribute("OPERATION", "GETDOCUMENTARCHIVE");
	xmlRequest.setAttribute("USERID", userId);
	xmlRequest.setAttribute("UNITID", unitId);
	xmlRequest.setAttribute("MACHINENAME", machineId);
	xmlRequest.setAttribute("CHANNELID", channelId);
	xmlRequest.setAttribute("USERAUTHORITYLEVEL", userAuthorityLevel);
	xmlRequest.setAttribute("APPLICATIONNUMBER", applicationNumber);
	xmlRequest.setAttribute("EDITDOCUMENT", readOnly ? "FALSE" : "TRUE");
	xmlRequest.setAttribute("REPRINTDOCUMENT", printOnly ? "TRUE" : "FALSE");
	xmlRequest.setAttribute("DOCUMENTGUID", documentGuid);
	
	var xmlElement = xmlRequestDocument.createElement("GETCRITERIA");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("FILEGUID", fileGuid);
	xmlElement.setAttribute("FILEVERSION", version);

	switch (ScreenRules())
	{
		case 1:
		case 0:
			xmlRequestObj.RunASP(document, "omPMRequest.asp");
			break;
		default:
			xmlRequestObj.SetErrorResponse();
	}

	var errorTypes = new Array("COMMANDFAILED");
	errorReturn = xmlRequestObj.CheckResponse(errorTypes);
	if (errorReturn[0])
	{
		var attribute1 = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS");
		var attribute2 = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@FILECONTENTSURL");
		if ((attribute1 == null || attribute1.text.length == 0) && (attribute2 == null || attribute2.text.length == 0))
		{
			alert("The document archive details are not yet available, please try again later.");
		}
		else
		{
			xmlRequestObj.SelectTag(null, "RESPONSE");	
		}
	}
	else
	{
		xmlRequestObj = null;
		if (errorReturn[1] == "COMMANDFAILED")
		{	
			var items = errorReturn[2].split(":");
			alert("The document is currently locked by user '" + items[items.length - 1] + "', please try again later.");
		}
	}
	
	return xmlRequestObj;
}

var m_fileContents = null;
var m_fileContentsUrl = null;
var m_archiveDeliveryType = null;
var m_archivedFileExtension = null;

function getArchivedDocumentForRow(table, selectedRow, readOnly, printOnly) 
{
	var version = table.rows(selectedRow).getAttribute("Version");
	var fileGuid = table.rows(selectedRow).getAttribute("FileGuid");
	var documentGuid = table.rows(selectedRow).getAttribute("DocGuid");
	
	// AS 09/08/2006 EP1078 omGemini: viewing an edited document in DMS107 does not return the document
	m_fileContents = null;
	m_fileContentsUrl = null;
	m_archiveDeliveryType = null;

	var xmlRequestObj = 
		getArchivedDocument(
			m_BaseNonPopupWindow, m_sUserId, m_sUnitId, m_sMachineId, m_sDistributionChannelId, "10", m_sApplicationNumber,
			documentGuid, fileGuid, version, readOnly, printOnly);
			
	if (xmlRequestObj != null)
	{
		var attribute = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS");
		if (attribute != null)
		{
			m_fileContents = attribute.text;
		}

		// AS 04/08/2006 EP1068 omGemini: Add support for scanned documents		
		var attribute = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@FILECONTENTSURL");
		if (attribute != null)
		{
			m_fileContentsUrl = attribute.text;
		}
		
		attribute = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@DELIVERYTYPE");
		if (attribute != null)
		{
			m_archiveDeliveryType = attribute.text;
		}
	}
	
	if ((m_fileContents == null && m_fileContentsUrl == null) || m_archiveDeliveryType == null)
	{
		xmlRequestObj = null;
	}
	
	return xmlRequestObj;
}

function displayArchivedDocument(archivedDocument, documentHostTemplateID, documentTitle, documentDeliveryType, printerType, readOnly, printOnly, documentFileExtension) 
{
	var printDocumentData = null;

	if (archivedDocument != null)
	{			
		var compressionMethod = "";
		printDocumentData = printDocument(archivedDocument, documentHostTemplateID, documentTitle, documentDeliveryType, compressionMethod, printerType, 1, true, true, readOnly, printOnly, documentFileExtension);
	}
	
	return printDocumentData;
}

function savePrintedDocument(
	window, userId, unitId, machineId, channelId, userAuthorityLevel, 
	xmlControlDataNode, xmlPrintDataNode, xmlTemplateDataNode, 
	fileSize, fileContents, compressionMethod)
{
	var xmlRequestObj = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
	var xmlRequest = xmlRequestDocument.createElement("REQUEST");
	xmlRequestDocument.appendChild(xmlRequest);
		
	xmlRequest.setAttribute("OPERATION", "PRINTDOCUMENT");
	if (userId != null) xmlRequest.setAttribute("USERID", userId);
	if (unitId != null) xmlRequest.setAttribute("UNITID", unitId);
	if (machineId != null) xmlRequest.setAttribute("MACHINENAME", machineId);
	if (channelId != null) xmlRequest.setAttribute("CHANNELID", channelId);
	if (userAuthorityLevel != null) xmlRequest.setAttribute("USERAUTHORITYLEVEL", userAuthorityLevel);
		
	if (xmlControlDataNode != null) xmlRequest.appendChild(xmlControlDataNode);		
	if (xmlPrintDataNode != null) xmlRequest.appendChild(xmlPrintDataNode);	
	if (xmlTemplateDataNode != null) xmlRequest.appendChild(xmlTemplateDataNode);
	
	var xmlElement = xmlRequestDocument.createElement("PRINTDOCUMENTDATA");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("FILEVERSION", "1");
	xmlElement.setAttribute("FILESIZE", fileSize);
	xmlElement.setAttribute("FILECONTENTS_TYPE", "BIN.BASE64");
	xmlElement.setAttribute("FILECONTENTS", fileContents);
	xmlElement.setAttribute("COMPRESSIONMETHOD", compressionMethod);
	xmlElement.setAttribute("COMPRESSED", compressionMethod == "" ? "0" : "1");
	
	xmlElement = xmlRequestDocument.createElement("AXWORD");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("RETURNIND", "1");
			
	switch (ScreenRules())
	{
		case 1: 
		case 0: 
			xmlRequestObj.RunASP(document, "PrintManager.asp");
			break;
		default:
			xmlRequestObj.SetErrorResponse();
	}

	return xmlRequestObj.IsResponseOK();
}

var m_xmlLocalPrinters = null;
var m_xmlComboPrinterDestination = null;
var m_xmlObjectComboPrinterDestination = null;

function getComboPrinterDestination(window, xmlComboPrinterDestination)
{
	if (xmlComboPrinterDestination == null)
	{
		m_xmlObjectComboPrinterDestination = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_xmlObjectComboPrinterDestination.GetComboLists(document, new Array("PrinterDestination"));
		xmlComboPrinterDestination = m_xmlObjectComboPrinterDestination.XMLDocument;
	}
	
	return xmlComboPrinterDestination;
}	

function getPrinterType(window, printerTypeId, userPrinterLocation, remotePrinterLocation)
{	
	var printerType = "L"; // Default to Local.
	
	m_xmlComboPrinterDestination = getComboPrinterDestination(window, m_xmlComboPrinterDestination);
	
	if (m_xmlComboPrinterDestination != null && printerTypeId != null)
	{	
		var xmlNodes = m_xmlComboPrinterDestination.selectNodes("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID='" + printerTypeId + "']/VALIDATIONTYPELIST/VALIDATIONTYPE");
		if (xmlNodes.length > 0)
		{
			printerType = xmlNodes.item(0).text;
			if	(
					xmlNodes.length == 2 &&
					(m_xmlComboPrinterDestination.selectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID='" + printerTypeId + "']/VALIDATIONTYPELIST/VALIDATIONTYPE[text()='L']") != null) &&
					(m_xmlComboPrinterDestination.selectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID='" + printerTypeId + "']/VALIDATIONTYPELIST/VALIDATIONTYPE[text()='R']") != null)
				)
			{
				// Template has a printer destination of "Local and Remote".
				if (userPrinterLocation != null && remotePrinterLocation != null && userPrinterLocation == remotePrinterLocation)
				{
					// User has chosen same printer as remote printer defined on template record, so print as remote.
					printerType = "R";
				}
				else							
				{
					printerType = "L";
				}
			}
		}
	}
		
	return printerType;
}

function getPrinterTypeId(window, printerType)
{
	var printTypeId = null;
	
	m_xmlComboPrinterDestination = getComboPrinterDestination(window, m_xmlComboPrinterDestination);

	if (m_xmlComboPrinterDestination != null)
	{
		var xmlNode = m_xmlComboPrinterDestination.selectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY/VALIDATIONTYPE/[VALIDATIONTYPE='" + printerType + "']/VALUEID");
		if (xmlNode != null)
		{
			printTypeId = xmlNode.text;
		}		
	}
	
	return printTypeId;
}

function getDefaultPrinterFromType(window, printerType)
{
	return getDefaultPrinterFromTypeId(window, getPrinterTypeId(window, printerType));
}

function getDefaultPrinterFromTypeId(window, printerTypeId)
{
	var defaultPrinter = null;
	
	var printerType = getPrinterType(window, printerTypeId);
	
	if (printerType == "L")
	{				
		m_xmlLocalPrinters = getLocalPrinters(m_xmlLocalPrinters);				
		if (m_xmlLocalPrinters != null)
		{	
			var xmlNode = m_xmlLocalPrinters.selectSingleNode("RESPONSE/PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
			if (xmlNode != null)
			{
				defaultPrinter = xmlNode.text;
			}
		}
		if (defaultPrinter == null || defaultPrinter == "")
		{					
			alert("You do not have a default printer set on your PC.");			
		}
	}
	else if (printerType == "W")
	{
		m_xmlComboPrinterDestination = getComboPrinterDestination(window, m_xmlComboPrinterDestination);

		if (m_xmlComboPrinterDestination != null)
		{
			var xmlNode = m_xmlComboPrinterDestination.selectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID='" + printerTypeId + "']/VALUENAME");
			if (xmlNode != null)
			{
				defaultPrinter = xmlNode.text;
			}
		}
	}
	else
	{				
		alert("The document template printer destination has been defined incorrectly. See your system administrator.");
	}
	
	return defaultPrinter;
}

function getPrintDocumentAttributes(sDocumentID)
{
	var attribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	attribsXML.CreateRequestTag(window , "GetPrintAttributes");
	attribsXML.CreateActiveTag("FINDATTRIBUTES");			
	attribsXML.SetAttribute("HOSTTEMPLATEID", sDocumentID);
									
	attribsXML.RunASP(document, "PrintManager.asp");
				
	if(attribsXML.IsResponseOK())
	{
		return attribsXML;
	}
	else
	{
		alert("Could not find Print Document Attributes. Please contact your System Administrator.");
		return false;
	}
}

// JD MAR838
// PSC 16/01/2006 MAR1054 - Start 
function getDocumentFromFileNet(tblTable, iCurrentRow, sFileNetUserId)
{

	var documentVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
	var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGUID");
	
	var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
	var xmlRequest = xmlRequestDocument.createElement("REQUEST");
	
	xmlRequestDocument.appendChild(xmlRequest);

	xmlRequest.setAttribute("OPERATION", "GETFILENETRECORD");
	xmlRequest.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xmlRequest.setAttribute("UNITID", m_sUnitId);
	
	// PSC 16/01/2006 MAR1054
	xmlRequest.setAttribute("USERID", sFileNetUserId);
	xmlRequest.setAttribute("FILENETIMAGEREF", fileNetGuid );

	xmlRequestObj.RunASP(document, "omPackRequest.asp");
	
	xmlRequestObj.SelectTag(null, "RESPONSE");
	if (xmlRequestObj.IsResponseOK()) 
	{
		var attribute = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@FILECONTENTS");
		if (attribute != null)
		{
			m_fileContents = attribute.text;
			attribute = xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/DOCUMENTCONTENTS/@CONTENTTYPE");
			if (attribute != null)
				m_archiveFileExtension = "." + attribute.text;
			return xmlRequestObj;
		}
	}
	else
	{
		return null;
	}
}

// EP462 IK
function ArchiveAvailable(iCurrentRow)
{
	var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	var xmlElem = xmlRequestDocument.createElement("REQUEST");
	xmlElem.setAttribute("CRUD_OP","READ");
	xmlElem.setAttribute("SCHEMA_NAME","omFVS");
	xmlElem.setAttribute("ENTITY_REF","FVFILE");
	var xmlNode = xmlRequestDocument.appendChild(xmlElem);
	xmlElem = xmlRequestDocument.createElement("FVFILE");
	xmlElem.setAttribute("FILEGUID",tblTable.rows(iCurrentRow).getAttribute("FileGuid"));
	xmlElem.setAttribute("FILEVERSION",tblTable.rows(iCurrentRow).getAttribute("Version"));
	xmlNode.appendChild(xmlElem);

	xmlRequestObj.RunASP(document, "omCRUDIf.asp");
	
	if(!xmlRequestObj.IsResponseOK()) 
		return(false);
	
	if(xmlRequestObj.XMLDocument.selectSingleNode("RESPONSE/FVFILE/FVFILECONTENTS[@FILECONTENTS_TYPE='GEMINI'][not(@FILECONTENTS)]"))
		return(false);
	
	return(true);	
}
// EP462 IK ends
