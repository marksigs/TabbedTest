/*
--------------------------------------------------------------------------------------------
Workfile:			Gemini.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Gemini helper functions
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		13/12/2006  First version.
AS		20/12/2006  CORE325 omPM: Gemini printing pack handling
AS		04/01/2007  CORE327 HasValue not trying to print pack document outside of pack.
AW		25/01/2007  EP1308 HasValue locked status of documents.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from Gemini.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

// Ambiguous reference in cref attribute.
#pragma warning disable 419

namespace Vertex.Fsd.Omiga.VisualBasicPort.Printing
{
	/// <summary>
	/// Support routines for Gemini interface.
	/// </summary>
	public static class Gemini
	{
		private const int GEMINIPRINTSTATUS_AWAITINGAPPROVAL = 10;
		private const int GEMINIPRINTSTATUS_NOTAPPROVED = 20;
		private const int GEMINIPRINTSTATUS_APPROVED = 30;
		private const int GEMINIPRINTSTATUS_GEMINIPRINTED = 40;
		private const int GEMINIPRINTMODE_NEVER = 10;
		private const int GEMINIPRINTMODE_IMMEDIATE = 20;
		private const int GEMINIPRINTMODE_ONHOLD = 30;
		private const string GEMINILOCATION_DMS = "DMS";
		private const string GEMINILOCATION_MOBIUSPENDING = "Mobius: pending";
		private const string m_fulfillmentRequestNamespace = "http://Omiga.Fsd.Vertex/Gemini.Fulfillment.Request";
		private const string m_geminiQueueName = "GeminiInterface";
		private const string m_geminiComponentName = "omGemini.FileVersioningBO";
		private const string MODULENAME = "Gemini";
		private const int ERRORNUMBERBASE = -2147221504 + 512 + (int)OMIGAERROR.RecordNotFound;
		//private static TraceAssist _traceAssist = new TraceAssist("omGemini");

		/// <summary>
		/// Sends a one or more packs to Gemini for fulfillment (printing at Gemini).
		/// </summary>
		/// <param name="pack">The input pack.</param>
		/// <returns>A list of packs.</returns>
		public static List<GeminiPack> GeminiSendToFulfillment(GeminiPack pack) 
		{
			return GeminiSendToFulfillment(pack, false, false, false, false, false);
		}

		/// <summary>
		/// Sends a one or more packs to Gemini for fulfillment (printing at Gemini).
		/// </summary>
		/// <param name="pack">The input pack.</param>
		/// <param name="mustExist">If true then each document must already exist in DMS or on Mobius, 
		/// i.e., Omiga must have received the notification of the documents location in 
		/// Mobius.
		/// </param>
		/// <param name="printOnHold">If true then an Gemini print a document even if the associated template is set to
		/// "On-hold" (or "Immediate"), i.e., document can be created but is only sent to Gemini printing
		/// by using the Gemini Print button on DMS110.
		/// </param>
		/// <param name="errorIfGeminiPrintNever">If true then an error is raised if attempting to Gemini print a
		/// document where the associated template is set to "Never", i.e., never Gemini print the
		/// document.
		/// If false then no error is raised, but the document(s) is not Gemini Printed.
		/// This should be set to true if printing a pack or packs, and false if printing a single document from PM010.
		/// </param>
		/// <param name="errorIfGeminiPrinted">If true then an error is raised if attempting to Gemini print a
		/// document that has already been Gemini Printed.
		/// </param>
		/// <param name="errorIfDocumentNotInPack">If true then an error is raised if attempting to Gemini print a document that is
		/// not in a pack and the template for the document is part of a pack definition.
		/// If false then no error is raised, but the document is not Gemini Printed.
		/// This should be set to true when Gemini Printing from DMS110, and false when Gemini Printing from PM010.
		/// </param>
		/// <returns>A list of packs.</returns>
		/// <remarks>
		/// <para>
		/// On input, the pack parameter should be defined as follows:
		/// <list type="bullet">
		///		<item><description>
		///		pack.PackFulfillmentGuid - specify to include all documents in this pack, and to not
		///		include any other associated packs.
		///		</description></item>
		///		<item><description>
		///		pack.ApplicationNumber - set to the mortgage application number.
		///		</description></item>
		///		<item><description>
		///		pack.UserId - set to the user initiating the request.
		///		</description></item>
		///		<item><description>
		///		pack.UnitId - set to the unit id for the user.
		///		</description></item>
		///		<item><description>
		///		Documents - set as follows: it is only necessary to specify details for one document 
		///		in the pack (the details for the other documents will be read from the database), UNLESS
		///		you are supplying the file contents for a document, in which case include its
		///		details as follows.
		///		<para>
		///		For each document specify either the FileContentsGuid or the
		///		DocumentGuid field. If supplying the document contents, then specify all the sub-fields in the
		///		DocumentDetails field. If these sub-fields are not supplied, then they will be
		///		retrieved either from DMS or Mobius. Do not specify the other document fields as these will be 
		///		read from the database.
		///		</para>
		///		</description></item>
		/// </list>
		/// </para>
		/// <para>
		/// For the specified pack, the fulfillment request will contain
		/// all the associated documents: if pack.PackFulfillmentGuid is specified, then includes all 
		/// documents in that pack, but does not include any other packs. If the document is part of a 
		/// pack, include all documents in the pack. If the document is in more than one pack, include 
		/// all the packs, and all the documents in these packs, UNLESS pack.PackFulfillmentGuid is 
		/// specified. If any of the documents in any of the packs is not available for Gemini printing
		/// (e.g., it does not exist in Mobius, or it is on-hold), then do not include any
		/// of these packs.
		/// </para>
		/// </remarks>
		public static List<GeminiPack> GeminiSendToFulfillment(GeminiPack pack, bool mustExist, bool printOnHold, bool errorIfGeminiPrintNever, bool errorIfGeminiPrinted, bool errorIfDocumentNotInPack) 
		{
			const string cstrFunctionName = "GeminiSendToFulfillment";
			List<GeminiPack> packs = null;

			ContextUtility.SetComplete();

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "(mustExist = " + Convert.ToString(mustExist) + ", printOnHold = " + Convert.ToString(printOnHold) + ", errorIfGeminiPrintNever = " + Convert.ToString(errorIfGeminiPrintNever) + ")");

			// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

			try
			{
				CheckInputPack(pack);

				packs = GetPacks(pack);

				if (CheckOutputPacks(packs, mustExist, printOnHold, errorIfGeminiPrintNever, errorIfGeminiPrinted, errorIfDocumentNotInPack))
				{
					// Packs are valid for sending to Gemini Printing
					GeminiSendMessagesToQueue(ToFulfillmentRequestXml(packs), "Fulfillment");
					CreateEventAuditDetails(packs, PrintManagerConstants.EVENTKEY_GEMINIPRINTED, GEMINIPRINTSTATUS_GEMINIPRINTED);
				}
			}
			catch
			{
				ContextUtility.SetAbort();
				throw;
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");
			
			return packs;
		}

		/// <summary>
		/// Sends a one or more packs to Gemini for fulfillment (printing at Gemini).
		/// </summary>
		/// <param name="xmlRequestNode">The input pack in xml format.</param>
		/// <remarks>
		/// This method is an xml wrapper for <see cref="GeminiSendToFulfillment"/>.
		/// </remarks>
		public static void GeminiSendToFulfillmentXml(XmlNode xmlRequestNode) 
		{
			const string cstrFunctionName = "GeminiSendToFulfillmentXml";
			GeminiPack pack = new GeminiPack();

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");
			// _traceAssist.TraceXML(xmlRequestNode.OuterXml, MODULENAME + "_" + cstrFunctionName + "_request");
			// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

			bool mustExist = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "MUSTEXIST", "0");
			bool printOnHold = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "PRINTONHOLD", "0");
			bool errorIfGeminiPrintNever = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "ERRORIFGEMINIPRINTNEVER", "0");
			bool errorIfGeminiPrinted = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "ERRORIFGEMINIPRINTED", "0");
			bool errorIfDocumentNotInPack = XmlAssist.GetAttributeAsBoolean(xmlRequestNode, "ERRORIFDOCUMENTNOTINPACK", "0");

			XmlNode xmlPackNode = XmlAssist.GetMandatoryNode(xmlRequestNode, "PACK");

			pack.PackFulfillmentGuid = XmlAssist.GetAttributeText(xmlPackNode, "PACKFULFILLMENTGUID");
			pack.ApplicationNumber = XmlAssist.GetMandatoryAttributeText(xmlPackNode, "APPLICATIONNUMBER");
			pack.UserId = XmlAssist.GetMandatoryAttributeText(xmlRequestNode, "USERID");
			pack.UnitId = XmlAssist.GetMandatoryAttributeText(xmlRequestNode, "UNITID");

			XmlNodeList xmlDocumentNodes = XmlAssist.GetMandatoryNodeList(xmlPackNode, "DOCUMENTLIST/DOCUMENT");
			pack.Documents = new List<GeminiDocument>(xmlDocumentNodes.Count);
			int documentIndex = 0;
			foreach (XmlNode xmlDocumentNode in xmlDocumentNodes)
			{
				pack.Documents[documentIndex].DocumentGuid = XmlAssist.GetAttributeText(xmlDocumentNode, "DOCUMENTGUID");
				pack.Documents[documentIndex].FileContentsGuid = XmlAssist.GetAttributeText(xmlDocumentNode, "FILECONTENTSGUID");

				XmlNode xmlDocumentDetailsNode = XmlAssist.GetNode(xmlDocumentNode, "DOCUMENTDETAILS");
				if (xmlDocumentDetailsNode != null)
				{
					pack.Documents[documentIndex].DocumentDetails.CompressionMethod = XmlAssist.GetAttributeText(xmlDocumentDetailsNode, "COMPRESSIONMETHOD");
					pack.Documents[documentIndex].DocumentDetails.DeliveryType = Convert.ToInt32(XmlAssist.GetAttributeText(xmlDocumentDetailsNode, "DELIVERYTYPE"));
					pack.Documents[documentIndex].DocumentDetails.FileContentsType = XmlAssist.GetAttributeText(xmlDocumentDetailsNode, "FILECONTENTS_TYPE");
					pack.Documents[documentIndex].DocumentDetails.FileContents = XmlAssist.GetAttributeText(xmlDocumentDetailsNode, "FILECONTENTS");
				}
				documentIndex++;
			}

			List<GeminiPack> packs = GeminiSendToFulfillment(pack, mustExist, printOnHold, errorIfGeminiPrintNever, errorIfGeminiPrinted, errorIfDocumentNotInPack);

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");
		}

		private static void CheckInputPack(GeminiPack pack) 
		{
			const string cstrFunctionName = "CheckInputPack";
			string reference = "";

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");

			if (pack.ApplicationNumber.Length == 0)
			{
				throw new ErrAssistException(ERRORNUMBERBASE, "Invalid pack.ApplicationNumber");
			}

			reference = "Details: ApplicationNumber=" + pack.ApplicationNumber;

			if (pack.PackFulfillmentGuid.Length == 0)
			{
				if (pack.Documents.Count == 0)
				{
					throw new ErrAssistException(ERRORNUMBERBASE, "Invalid pack Document: " + reference);
				}

				for (int documentIndex = 0; documentIndex < pack.Documents.Count; documentIndex++)
				{
					string refDocument = reference + ", documentIndex=" + Convert.ToString(documentIndex);

					if (pack.Documents[documentIndex].DocumentGuid.Length == 0 && pack.Documents[documentIndex].FileContentsGuid.Length == 0)
					{
						throw new ErrAssistException(ERRORNUMBERBASE, "Invalid DocumentGuid or FileContentsGuid: " + refDocument);
					}
				}
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
		}

		private static bool CheckOutputPacks(List<GeminiPack> packs, bool mustExist, bool printOnHold, bool errorIfGeminiPrintNever, bool errorIfGeminiPrinted, bool errorIfDocumentNotInPack) 
		{
			const string cstrFunctionName = "CheckOutputPacks";
			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");

			bool success = true;

			if (packs.Count == 0)
			{
				throw new ErrAssistException(ERRORNUMBERBASE, "Invalid packs");
			}

			// At least one pack.
			for (int packIndex = 0; packIndex < packs.Count; packIndex++)
			{
				if (success)
				{
					GeminiPack pack = packs[packIndex];
					string refPack = GetPackRef(pack);

					if (pack.Documents.Count == 0)
					{
						throw new ErrAssistException(ERRORNUMBERBASE, "Invalid pack Documents: " + refPack);
					}

					for (int documentIndex = 0; documentIndex < pack.Documents.Count; documentIndex++)
					{
						if (success)
						{
							GeminiDocument document = pack.Documents[documentIndex];

							string refDocument = GetDocumentRef(document, refPack);

							string strUserId = null;
							if (IsFileVersionLocked(document, out strUserId))
							{
								throw new ErrAssistException(ERRORNUMBERBASE, "The following document is curently being amended by " + strUserId + " " + ": " + refDocument);
							}

							if (
								document.InputDocument &&
								document.PrintStatus == GEMINIPRINTSTATUS_GEMINIPRINTED &&
								errorIfGeminiPrinted)
							{
								// The input document has already been Gemini Printed and it is an error
								// to attempt to Gemini Print a document more than once.
								// Ignore print status of other documents in the pack as these may have been
								// Gemini Printed in other packs.
								throw new ErrAssistException(ERRORNUMBERBASE, "Invalid Gemini Print Status (" + GeminiPrintStatusToString(document.PrintStatus) + "): " + refDocument);
							}

							if (document.PrintMode == GEMINIPRINTMODE_NEVER)
							{
								// Document is Gemini Print Never.
								if (errorIfGeminiPrintNever)
								{
									// And this should be reported as an error.
									throw new ErrAssistException(ERRORNUMBERBASE, "Invalid Gemini Print Mode (" + GeminiPrintModeToString(document.PrintMode) + "): " + refDocument);
								}
								else
								{
									// Otherwise do not continue with Gemini Printing.
									success = false;
								}
							}

							if (success && document.PrintMode == GEMINIPRINTMODE_ONHOLD)
							{
								// Document is Gemini Print On-Hold, so only Gemini Print if printOnHold is true, i.e.,
								// called from DMS110.
								success = printOnHold;
							}

							if (success &&
								document.PrintStatus != GEMINIPRINTSTATUS_APPROVED &&
								document.PrintStatus != GEMINIPRINTSTATUS_GEMINIPRINTED)
							{
								// Document is not Approved or not Gemini Printed.
								throw new ErrAssistException(ERRORNUMBERBASE, "Invalid Gemini Print Status (" + GeminiPrintStatusToString(document.PrintStatus) + "): " + refDocument);
							}

							if (success &&
								mustExist &&
								(document.DocumentLocation).ToUpper() == GEMINILOCATION_MOBIUSPENDING.ToUpper() &&
								document.DocumentDetails.FileContents.Length == 0)
							{
								// Document has been sent to Mobius but no notification has been received back as to its
								// location in Mobius, and contents of the document where not included in the input pack.
								throw new ErrAssistException(ERRORNUMBERBASE, "Invalid Document Location (" + document.DocumentLocation + "): " + refDocument);
							}

							if (success && pack.PackFulfillmentGuid.Length == 0 && document.TemplatePackMember)
							{
								// Document is not part of a pack and the document template is part of a pack definition.
								if (errorIfDocumentNotInPack)
								{
									// And this should be reported as an error.
									throw new ErrAssistException(ERRORNUMBERBASE, "Document is not part of a pack: " + refDocument);
								}
								else
								{
									// Otherwise do not continue with Gemini Printing.
									success = false;
								}
							}
						}
					}
				}
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");

			return success;
		}

		private static string GetPackRef(GeminiPack pack) 
		{
			string refPack = "Details: ApplicationNumber=" + pack.ApplicationNumber;

			if (pack.PackControlName.Length > 0)
			{
				refPack = refPack + ", PackName=" + pack.PackControlName;
			}

			if (pack.PackFulfillmentGuid.Length > 0)
			{
				refPack = refPack + ", PackCreationDate=" + Convert.ToString(pack.PackCreationDate);
				refPack = refPack + ", PackFulfillmentGuid=" + pack.PackFulfillmentGuid;
			}

			return refPack;
		}

		private static string GetDocumentRef(GeminiDocument document, string refPack) 
		{
			string refDocument = refPack;

			if (document.DocumentVersion.Length > 0)
			{
				refDocument = refDocument + ", DocumentVersion=" + document.DocumentVersion;
			}

			if (document.DocumentName.Length > 0)
			{
				refDocument = refDocument + ", DocumentName=" + document.DocumentName;
			}

			if (document.DocumentVersion.Length > 0)
			{
				refDocument = refDocument + ", DocumentDate=" + Convert.ToString(document.DocumentDate);
			}

			if (document.DocumentGuid.Length > 0)
			{
				refDocument = refDocument + ", DocumentGuid=" + document.DocumentGuid;
			}

			if (document.FileContentsGuid.Length > 0)
			{
				refDocument = refDocument + ", FileContentsGuid=" + document.FileContentsGuid;
			}

			return refDocument;
		}

		private static string GeminiPrintModeToString(int geminiPrintMode) 
		{
			string GeminiPrintModeToString = "";
			switch (geminiPrintMode)
			{
				case GEMINIPRINTMODE_NEVER:
					GeminiPrintModeToString = "Never";
					break;
				case GEMINIPRINTMODE_IMMEDIATE:
					GeminiPrintModeToString = "Immediate";
					break;
				case GEMINIPRINTMODE_ONHOLD:
					GeminiPrintModeToString = "On-hold";
					break;
				default:
					GeminiPrintModeToString = "Unknown";
					break; 
			}
			return GeminiPrintModeToString;
		}

		private static string GeminiPrintStatusToString(int geminiPrintStatus) 
		{
			string GeminiPrintStatusToString = "";
			switch (geminiPrintStatus)
			{
				case GEMINIPRINTSTATUS_AWAITINGAPPROVAL:
					GeminiPrintStatusToString = "Awaiting approval";
					break;
				case GEMINIPRINTSTATUS_NOTAPPROVED:
					GeminiPrintStatusToString = "Not approved";
					break;
				case GEMINIPRINTSTATUS_APPROVED:
					GeminiPrintStatusToString = "Approved";
					break;
				case GEMINIPRINTSTATUS_GEMINIPRINTED:
					GeminiPrintStatusToString = "Gemini printed";
					break;
				default:
					GeminiPrintStatusToString = "Unknown";
					break; 
			}
			return GeminiPrintStatusToString;
		}

		/// <summary>
		/// Converts a list of GeminiPack items into a fulfillment request.
		/// </summary>
		/// <param name="packs">The list of GeminiPack items to convert into the fulfillment request.</param>
		/// <returns>The fulfillment request, as an Xml string.</returns>
		/// <remarks>
		/// Each pack is a separate request. This enables each request to be put onto the
		/// queue as a separate message. If one of the requests fails, it will not affect the others,
		/// and it can be resubmitted to the queue.
		/// </remarks>
		private static string[] ToFulfillmentRequestXml(List<GeminiPack> packs) 
		{
			const string cstrFunctionName = "ToFulfillmentRequestXml";

			string[] requests = new string[packs.Count];

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

			for (int packIndex = 0; packIndex < packs.Count; packIndex++)
			{
				XmlDocument xmlRequestDocument = new XmlDocument();

				XmlNode xmlRequest = xmlRequestDocument.CreateNode(XmlNodeType.Element, "FULFILLMENTREQUEST", m_fulfillmentRequestNamespace);
				xmlRequestDocument.AppendChild(xmlRequest);

				XmlElement xmlPackList = (XmlElement)xmlRequestDocument.CreateNode(XmlNodeType.Element, "PACKLIST", m_fulfillmentRequestNamespace);
				xmlRequest.AppendChild(xmlPackList);

				XmlElement xmlPack = (XmlElement)xmlRequestDocument.CreateNode(XmlNodeType.Element, "PACK", m_fulfillmentRequestNamespace);
				xmlPackList.AppendChild(xmlPack);
				xmlPack.SetAttribute("APPLICATIONNUMBER", packs[packIndex].ApplicationNumber);

				XmlElement xmlDocumentList = (XmlElement)xmlRequestDocument.CreateNode(XmlNodeType.Element, "DOCUMENTLIST", m_fulfillmentRequestNamespace);
				xmlPack.AppendChild(xmlDocumentList);

				for (int documentIndex = 0; documentIndex < packs[packIndex].Documents.Count; documentIndex++)
				{
					if (packs[packIndex].Documents[documentIndex].PrintMode != GEMINIPRINTMODE_NEVER)
					{
						XmlElement xmlDocument = (XmlElement)xmlRequestDocument.CreateNode(XmlNodeType.Element, "DOCUMENT", m_fulfillmentRequestNamespace);
						xmlDocumentList.AppendChild(xmlDocument);
						xmlDocument.SetAttribute("FILECONTENTSGUID", packs[packIndex].Documents[documentIndex].FileContentsGuid);

						if (packs[packIndex].Documents[documentIndex].DocumentDetails.FileContents.Length > 0)
						{
							XmlElement xmlDocumentDetails = (XmlElement)xmlRequestDocument.CreateNode(XmlNodeType.Element, "DOCUMENTDETAILS", m_fulfillmentRequestNamespace);
							xmlDocument.AppendChild(xmlDocumentDetails);
							xmlDocumentDetails.SetAttribute("COMPRESSIONMETHOD", packs[packIndex].Documents[documentIndex].DocumentDetails.CompressionMethod);
							xmlDocumentDetails.SetAttribute("DELIVERYTYPE", Convert.ToString(packs[packIndex].Documents[documentIndex].DocumentDetails.DeliveryType));
							xmlDocumentDetails.SetAttribute("FILECONTENTS_TYPE", packs[packIndex].Documents[documentIndex].DocumentDetails.FileContentsType);
							xmlDocumentDetails.SetAttribute("FILECONTENTS", packs[packIndex].Documents[documentIndex].DocumentDetails.FileContents);
						}
					}
				}

				requests[packIndex] = xmlRequestDocument.OuterXml;
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");
			
			return requests;
		}

		/// <summary>
		/// Gets a list of GeminiPack items for a specified document or file contents.
		/// </summary>
		/// <param name="pack">The input pack.</param>
		/// <returns>A list of packs.</returns>
		/// <remarks>
		/// For the specified document or file contents, the fulfillment request will contain
		/// all the associated documents. If the document is part of a pack, include all documents 
		/// in the pack. If the document is in more than one pack, include all the packs, and all the
		/// documents in these packs. If any of the documents in any of the packs is not available 
		/// for Gemini printing (e.g., it does not exist in Mobius, or it is on-hold), then do not 
		/// include any of these packs.
		/// </remarks>
		private static List<GeminiPack> GetPacks(GeminiPack pack) 
		{
			const string cstrFunctionName = "GetPacks";

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

			List<GeminiPack> packs = new List<GeminiPack>();

			using (SqlConnection adoConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
			{
				using (SqlCommand adoCommand = new SqlCommand("usp_GeminiGetPacks", adoConnection))
				{
					adoCommand.CommandType = CommandType.StoredProcedure;

					SqlParameter adoParameter = null;
					adoParameter = adoCommand.Parameters.Add("applicationNumber", SqlDbType.NVarChar, 12);
					adoParameter.Value = pack.ApplicationNumber;

					adoParameter = adoCommand.Parameters.Add("packFulfillmentGuid", SqlDbType.Binary, 16);
					adoParameter.IsNullable = true;
					if (pack.PackFulfillmentGuid != null && pack.PackFulfillmentGuid.Length != 0)
					{
						adoParameter.Value = SqlAssist.GuidStringToByteArray(pack.PackFulfillmentGuid);
					}
					else
					{
						adoParameter.Value = System.DBNull.Value;
					}

					adoParameter = adoCommand.Parameters.Add("documentGuid", SqlDbType.Binary, 16);
					adoParameter.IsNullable = true;
					if (pack.Documents.Count > 0 && pack.Documents[0].DocumentGuid != null && pack.Documents[0].DocumentGuid.Length != 0)
					{
						adoParameter.Value = SqlAssist.GuidStringToByteArray(pack.Documents[0].DocumentGuid);
					}
					else
					{
						adoParameter.Value = System.DBNull.Value;
					}

					adoParameter = adoCommand.Parameters.Add("fileContentsGuid", SqlDbType.Binary, 16);
					adoParameter.IsNullable = true;
					if (pack.Documents.Count > 0 && pack.Documents[0].FileContentsGuid != null && pack.Documents[0].FileContentsGuid.Length != 0)
					{
						adoParameter.Value = SqlAssist.GuidStringToByteArray(pack.Documents[0].FileContentsGuid);
					}
					else
					{
						adoParameter.Value = System.DBNull.Value;
					}

					adoConnection.Open();
					using (SqlDataReader adoRecordSet = adoCommand.ExecuteReader(CommandBehavior.CloseConnection))
					{
						int packIndex = 0;
						while (adoRecordSet.HasRows)
						{
							packs.Add(new GeminiPack());
							packs[packIndex].ApplicationNumber = pack.ApplicationNumber;
							packs[packIndex].UserId = pack.UserId;
							packs[packIndex].UnitId = pack.UnitId;

							int documentIndex = 0;
							while (adoRecordSet.Read())
							{
								packs[packIndex].Documents.Add(new GeminiDocument());
								if (adoRecordSet.FieldCount >= 14)
								{
									int fieldIndex = 0;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].PackFulfillmentGuid = GuidAssist.ToString((byte[])adoRecordSet[fieldIndex]);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].PackControlName = adoRecordSet.GetString(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].PackCreationDate = adoRecordSet.GetDateTime(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].DocumentGuid = GuidAssist.ToString((byte[])adoRecordSet[fieldIndex]);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].FileGuid = GuidAssist.ToString((byte[])adoRecordSet[fieldIndex]);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].FileContentsGuid = GuidAssist.ToString((byte[])adoRecordSet[fieldIndex]);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].DocumentVersion = adoRecordSet.GetString(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].DocumentDate = adoRecordSet.GetDateTime(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].DocumentName = adoRecordSet.GetString(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].HostTemplateId = adoRecordSet.GetString(fieldIndex);
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].TemplatePackMember = Convert.ToBoolean(adoRecordSet.GetValue(fieldIndex));
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].PrintStatus = Convert.ToInt32(adoRecordSet.GetValue(fieldIndex));
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].PrintMode = Convert.ToInt32(adoRecordSet.GetValue(fieldIndex));
									}

									fieldIndex++;
									if (!adoRecordSet.IsDBNull(fieldIndex))
									{
										packs[packIndex].Documents[documentIndex].DocumentLocation = adoRecordSet.GetString(fieldIndex);
									}

									if (packs[packIndex].Documents[documentIndex].FileContentsGuid.Length > 0 && pack.Documents.Count > 0)
									{
										for (int packDocumentIndex = 0; packDocumentIndex < pack.Documents.Count; packDocumentIndex++)
										{
											if (
												String.Compare(packs[packIndex].Documents[documentIndex].DocumentGuid, pack.Documents[packDocumentIndex].DocumentGuid, StringComparison.OrdinalIgnoreCase) == 0 ||
												String.Compare(packs[packIndex].Documents[documentIndex].FileContentsGuid, pack.Documents[packDocumentIndex].FileContentsGuid, StringComparison.OrdinalIgnoreCase) == 0)
											{
												// This pack document is the same document that was passed in.
												packs[packIndex].Documents[documentIndex].InputDocument = true;

												if (pack.Documents[packDocumentIndex].DocumentDetails.FileContents.Length > 0)
												{
													// Document details have been passed in so copy them in to the pack document.
													packs[packIndex].Documents[documentIndex].DocumentDetails = pack.Documents[packDocumentIndex].DocumentDetails;
												}
											}
										}
									}
								}

								documentIndex++;
							}

							adoRecordSet.NextResult();
							packIndex++;
						}
					}
				}
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");
			
			return packs;
		}

		/// <summary>
		/// Send messages to the Gemini queue.
		/// </summary>
		/// <param name="messages">The xml messages.</param>
		/// <param name="messageType">The type of the messages.</param>
		/// <returns>
		/// The responses to sending the messages.
		/// </returns>
		public static string[] GeminiSendMessagesToQueue(string[] messages, string messageType) 
		{
			string[] responses = new string[messages.Length];

			for (int messageIndex = 0; messageIndex < messages.Length; messageIndex++)
			{
				responses[messageIndex] = GeminiSendMessageToQueue(messages[messageIndex], messageType, messageIndex);
			}

			return responses;
		}

		/// <summary>
		/// Send a message to the Gemini queue.
		/// </summary>
		/// <param name="messageText">The xml message text.</param>
		/// <param name="messageType">The type of the message.</param>
		/// <param name="messageIndex">The index for the message.</param>
		/// <returns>
		/// The response to sending the message.
		/// </returns>
		public static string GeminiSendMessageToQueue(string messageText, string messageType, int messageIndex) 
		{
			const string cstrFunctionName = "GeminiSendMessageToQueue";
			omMQ.IOmigaToMessageQueue messageQueueObject = null;

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "(messageText=" + messageText + ")");

			string response = "";

			try
			{
				// _traceAssist.TraceXML(messageText, MODULENAME + "_" + messageType + Convert.ToString(messageIndex) + "_request");
				// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

				int messageQueueType = GlobalAssist.GetGlobalParamAmount("MessageQueueType");

				switch (messageQueueType)
				{
					case 1:
						// SQL Server MSMQ
						messageQueueObject = (omMQ.IOmigaToMessageQueue)new omToMSMQ.OmigaToMessageQueue();
						break;
					case 2:
					case 3:
						// SQL Server OMMQ Oracle OMMQ
						messageQueueObject = (omMQ.IOmigaToMessageQueue)new omToOMMQ.OmigaToMessageQueue();
						break;
					default:
						// Error Message Queue type not supported.
						throw new InvalidMessageQueueTypeException("Message Queue Type not supported, check global parameter MessageQueueType");
				}

				string request =
					"<REQUEST>" +
					"<MESSAGEQUEUE>" +
					"<QUEUENAME>" + m_geminiQueueName + "</QUEUENAME>" +
					"<PROGID>" + m_geminiComponentName + "</PROGID>" +
					"</MESSAGEQUEUE>" +
					"</REQUEST>";

				System.Diagnostics.Debug.WriteLine("->messageQueueObject.AsyncSend(request=" + request + ", messageText=" + messageText + ")");
				// _traceAssist.TraceMethodEntry(MODULENAME, "messageQueueObject.AsyncSend", request);
				response = messageQueueObject.AsyncSend(request, messageText);
				System.Diagnostics.Debug.WriteLine("<-messageQueueObject.AsyncSend() = " + response);
				// _traceAssist.TraceMethodExit(MODULENAME, "messageQueueObject.AsyncSend", response);

				ErrAssistException.CheckXmlResponse(response, true, null);
			}
			finally
			{
				if (messageQueueObject != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(messageQueueObject);
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "() = " + response);
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");

			return response;
		}

		private static bool CreateEventAuditDetails(List<GeminiPack> packs, short eventKey, short geminiPrintStatus) 
		{
			const string cstrFunctionName = "CreateEventAuditDetails";

#if !TEST
			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodEntry(MODULENAME, cstrFunctionName, "");

			omPM.PrintManagerBO objPM = null;

			try
			{
				objPM = new omPM.PrintManagerBO();

				for (int packIndex = 0; packIndex < packs.Count; packIndex++)
				{
					for (int documentIndex = 0; documentIndex < packs[packIndex].Documents.Count; documentIndex++)
					{
						string request =
							"<REQUEST " +
							"APPLICATIONNUMBER='" + packs[packIndex].ApplicationNumber + "' " +
							"USERID='" + packs[packIndex].UserId + "' " +
							"UNITID='" + packs[packIndex].UnitId + "' " +
							"DOCUMENTGUID='" + packs[packIndex].Documents[documentIndex].DocumentGuid + "' " +
							"OPERATION='CREATEAUDITTRAIL'>" +
							"<EVENTDETAIL " +
							"EVENTKEY='" + Convert.ToString(eventKey) + "' " +
							"DOCUMENTVERSION='" + packs[packIndex].Documents[documentIndex].DocumentVersion + "' " +
							"FILEGUID='" + packs[packIndex].Documents[documentIndex].FileGuid + "' " +
							"PACKFULFILLMENTGUID='" + packs[packIndex].PackFulfillmentGuid + "' " +
							"HOSTTEMPLATEID='" + packs[packIndex].Documents[documentIndex].HostTemplateId + "'>" +
							"</EVENTDETAIL>" +
							"<DOCUMENTDETAILS " +
							"GEMINIPRINTSTATUS='" + Convert.ToString(geminiPrintStatus) + "'>" +
							"</DOCUMENTDETAILS>" +
							"</REQUEST>";
						string response = objPM.omRequest(request);
					}
				}

			}
			finally
			{
				if (objPM != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(objPM);
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");
			// _traceAssist.TraceMethodExit(MODULENAME, cstrFunctionName, "");
#endif 
			return true;
		}

		private static bool IsFileVersionLocked(GeminiDocument document, out string strUserId) 
		{
			const string cstrFunctionName = "IsFileVersionLocked";
			bool blnIsLocked = false;
			strUserId = "";

			System.Diagnostics.Debug.WriteLine("->" + cstrFunctionName + "()");

			using (SqlConnection adoConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
			{
				using (SqlCommand adoCommand = new SqlCommand("usp_IsFvFileLocked", adoConnection))
				{
					adoCommand.CommandType = CommandType.StoredProcedure;

					SqlParameter adoParameter = null;
					adoParameter = adoCommand.Parameters.Add("FileGuid", SqlDbType.Binary, 16);
					adoParameter.IsNullable = true;
					if (document.FileContentsGuid != null && document.FileContentsGuid.Length != 0)
					{
						adoParameter.Value = SqlAssist.GuidStringToByteArray(document.FileContentsGuid);
					}
					else
					{
						adoParameter.Value = System.DBNull.Value;
					}

					adoParameter = adoCommand.Parameters.Add("FileVersion", SqlDbType.NVarChar, 20);
					adoParameter.IsNullable = true;
					if (document.DocumentVersion != null && document.DocumentVersion.Length != 0)
					{
						adoParameter.Value = document.DocumentVersion;
					}
					else
					{
						adoParameter.Value = System.DBNull.Value;
					}

					SqlParameter isLockedParameter = adoCommand.Parameters.Add("IsLocked", SqlDbType.TinyInt, 0);
					isLockedParameter.Direction = ParameterDirection.Output;

					SqlParameter userIdParameter = adoCommand.Parameters.Add("UserId", SqlDbType.NVarChar, 25);
					userIdParameter.Direction = ParameterDirection.Output;

					adoConnection.Open();
					adoCommand.ExecuteNonQuery();

					if (isLockedParameter.Value != System.DBNull.Value)
					{
						blnIsLocked = Convert.ToBoolean(isLockedParameter.Value);
					}

					if (userIdParameter.Value != System.DBNull.Value)
					{
						strUserId = Convert.ToString(userIdParameter.Value);
					}
				}
			}

			System.Diagnostics.Debug.WriteLine("<-" + cstrFunctionName + "()");

			return blnIsLocked;
		}
	}

	/// <summary>
	/// Represents a document pack being sent to Gemini.
	/// </summary>
	public class GeminiPack
	{
		/// <summary>
		/// A unique identifier for the pack.
		/// </summary>
		/// <remarks>
		/// Represents the APPLICATIONPACK.PACKFULFILLMENTGUID field.
		/// Specify on input or specify at least one document.
		/// </remarks>
		public string PackFulfillmentGuid = "";
		/// <summary>
		/// The pack control name.
		/// </summary>
		/// <remarks>
		/// Represents the PACKCONTROL.PACKCONTROLNAME field. 
		/// Read from the database.
		/// </remarks>
		public string PackControlName = "";
		/// <summary>
		/// The pack creation date.
		/// </summary>
		/// <remarks>
		/// Represents the APPLICATIONPACK.CREATIONDATE field. 
		/// Read from the database.
		/// </remarks>
		public DateTime PackCreationDate;
		/// <summary>
		/// The mortgage application number.
		/// </summary>
		/// <remarks>
		/// Represents the APPLICATIONPACK.APPLICATIONNUMBER field. 
		/// Specify on input.
		/// </remarks>
		public string ApplicationNumber = "";
		/// <summary>
		/// The Omiga user id.
		/// </summary>
		/// <remarks>
		/// Specify on input.
		/// </remarks>
		public string UserId = "";
		/// <summary>
		/// The Omiga user unit id.
		/// </summary>
		/// <remarks>
		/// Specify on input.
		/// </remarks>
		public string UnitId = "";
		/// <summary>
		/// The list of documents in this pack.
		/// </summary>
		/// <remarks>
		/// If <see cref="PackFulfillmentGuid"/> is not specified then specify at least one document on input.
		/// </remarks>
		public List<GeminiDocument> Documents = new List<GeminiDocument>();
	}

	/// <summary>
	/// Represents a document being sent to Gemini.
	/// </summary>
	public class GeminiDocument
	{
		/// <summary>
		/// The unique identifier for the document.
		/// </summary>
		/// <remarks>
		/// Represents the DOCUMENTAUDITDETAILS.DOCUMENTGUID field. 
		/// Specify on input or specify <see cref="FileContentsGuid"/>. 
		/// </remarks>
		public string DocumentGuid = "";
		/// <summary>
		/// The unique identifier for the file in FVS.
		/// </summary>
		/// <remarks>
		/// Represents the FVFILE.FILEGUID field.
		/// Read from the database.
		/// </remarks>
		public string FileGuid = "";
		/// <summary>
		/// The unique identifier for the file contents in FVS.
		/// </summary>
		/// <remarks>
		/// Represents the FVFILECONTENTS.FILECONTENTSGUID field. 
		/// Specify on input or specify <see cref="DocumentGuid"/>.
		/// </remarks>
		public string FileContentsGuid = "";
		/// <summary>
		/// The version of the document.
		/// </summary>
		/// <remarks>
		/// Represents the FVFILE.FILEVERSION field. 
		/// Read from the database.
		/// </remarks>
		public string DocumentVersion = "";
		/// <summary>
		/// The date of the document.
		/// </summary>
		/// <remarks>
		/// Represents the DOCUMENTAUDITDETAILS.CREATIONDATE field. 
		/// Read from the database.
		/// </remarks>
		public DateTime DocumentDate = DateTime.Now;
		/// <summary>
		/// The name of the document template.
		/// </summary>
		/// <remarks>
		/// Represents the HOSTTEMPLATE.HOSTTEMPLATENAME field. 
		/// Read from the database.
		/// </remarks>
		public string DocumentName = "";
		/// <summary>
		/// The unique identifier for the document template.
		/// </summary>
		/// <remarks>
		/// Represents the HOSTTEMPLATE.HOSTTEMPLATEID field. 
		/// Read from the database.
		/// </remarks>
		public string HostTemplateId = "";
		/// <summary>
		/// Indicates if the document is part of a pack.
		/// </summary>
		/// <remarks>
		/// Read from the database.
		/// </remarks>
		public bool TemplatePackMember = false;
		/// <summary>
		/// Indicates the print status for the document.
		/// </summary>
		/// <remarks>
		/// Represents the DOCUMENTAUDITDETAILS.GEMINIPRINTSTATUS field. 
		/// Read from the database.
		/// </remarks>
		public int PrintStatus = 0;
		/// <summary>
		/// Indicates the print mode for the document.
		/// </summary>
		/// <remarks>
		/// Represents the HOSTTEMPLATE.GEMINIPRINTMODE field. 
		/// Read from the database.
		/// </remarks>
		public int PrintMode = 0;
		/// <summary>
		/// The location of the document.
		/// </summary>
		/// <remarks>
		/// Read from the database.
		/// </remarks>
		public string DocumentLocation = "";
		/// <summary>
		/// Indicates if the document was supplied in the input.
		/// </summary>
		public bool InputDocument = false;
		/// <summary>
		/// Details for the document.
		/// </summary>
		/// <remarks>
		/// Specify on input if supplying document contents.
		/// </remarks>
		public GeminiDocumentDetails DocumentDetails = new GeminiDocumentDetails();
	}

	/// <summary>
	/// Represents the details of a document being sent to Gemini.
	/// </summary>
	public class GeminiDocumentDetails
	{
		/// <summary>
		/// The type of the file contents.
		/// </summary>
		/// <remarks>
		/// Represents the FVFILECONTENTS.FILECONTENTS_TYPE field. 
		/// Generally one of "BIN.BASE64" or "GEMINI".
		/// Specify on input if supplying document contents.
		/// </remarks>
		public string FileContentsType = "";
		/// <summary>
		/// The contents of the file.
		/// </summary>
		/// <remarks>
		/// Represents the FVFILECONTENTS.FILECONTENTS field. 
		/// Specify on input if supplying document contents.
		/// </remarks>
		public string FileContents = "";
		/// <summary>
		/// The compression method for the <see cref="FileContents"/>. 
		/// Represents the FVFILECONTENTS.COMPRESSIONMETHOD field. 
		/// One of "ZLIB", "COMPAPI", or "". 
		/// Specify on input if supplying document contents.
		/// </summary>
		public string CompressionMethod = "";
		// Specify on input if supplying document contents.
		/// <summary>
		/// The delivery type. 
		/// Represents the FVFILECONTENTS.DELIVERYTYPE field. 
		/// One of 10, 20, 30, 40 etc.
		/// Specify on input if supplying document contents.
		/// </summary>
		public int DeliveryType = 0;
	}

}
