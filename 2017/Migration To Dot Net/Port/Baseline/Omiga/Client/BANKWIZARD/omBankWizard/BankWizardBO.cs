/*
--------------------------------------------------------------------------------------------
Workfile:			BankWizardBO.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Validate bank details
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HMA		18/09/2006	Created
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;
using System.Collections;
using omLogging = Vertex.Fsd.Omiga.omLogging;
using EigerSystems.BankWizard;

namespace Vertex.Fsd.Omiga.omBankWizard
{
	public class BankWizardBO
	{
		
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		private const int defaultErrorCode = 500;		

		public BankWizardBO()
		{
		}

		public string GetBankDetails(string request)
		{
			string functionName ="GetBankDetails";
			string sortCode = "";
			string accountNumber = "";
			string rollNumber = "";
			string bankName = "";
			string branchName = "";
			string errorString = "";
			Boolean transferred = false;
			Boolean fatalErrorOccurred = false; 

			XmlElement responseElement = null;

			Boolean fatalErrorsExist = false;
			Boolean errorsExist = false;

			string[] fatalMessages;
			string[] errorMessages;

			long[] fatalConditions;
			long[] errorConditions;

			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(functionName + ": Request: " + request);
			}

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(request);

			sortCode = doc.SelectSingleNode("//ValidateBankDetailsRequestType/NewSortCode").InnerText;
			accountNumber = doc.SelectSingleNode("//ValidateBankDetailsRequestType/NewAccountNumber").InnerText;

			/* Initialise Bank Wizard */

			/* Initialise the country */
			BankWizardInt.setRuntime(BankWizardInt.R_INC_COUNTRY,BWIHandle.C_UK);

			/* Create a handle, which will store any condition codes which arise 
			 * from initialisation. The handle does not need to have an parameters
			 * set within it, it is used purely as a store for conditions. */
			BWIHandle initHandle = new BWIHandle();

			/* The core must be initialised before it can be used for validation. */
			BankWizardInt.init(initHandle);
			
			/*
			 * Display the results, normally nothing will be shown unless there is
			 * an error, or warning. Note that when a licence key is getting close to
			 * expiry a 30 day warning is output. An application should normally 
			 * handle this by logging the message in a system log, the message would
			 * not normally be displayed to a user, unless appropriate to the application.
			 */

			fatalErrorsExist = IterateErrorConditions(initHandle, BWIHandle.N_FATAL, out fatalConditions, out fatalMessages);
			errorsExist = IterateErrorConditions(initHandle, BWIHandle.N_ERROR, out errorConditions, out errorMessages);

			/* The handle used to initialise MUST not be used for any other tasks
				as it deals with initialisation conditions in a special way. */
			initHandle.free();

			/* If any errors have occurred during initialisation, report them and return */
			if(fatalErrorsExist == true) 
			{
				responseElement = SetErrorMessage ("APPERR", defaultErrorCode.ToString(), "GetBankDetails", fatalMessages[0]);
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + ": Response: " + responseElement.OuterXml);
				}

				return responseElement.OuterXml;
			}
			else if (errorsExist == true)
			{
				responseElement = SetErrorMessage ("APPERR", defaultErrorCode.ToString(), "GetBankDetails", errorMessages[0]);
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + ": Response: " + responseElement.OuterXml);
				}
				return responseElement.OuterXml;
			}

			/* If the initialisation has succeeded continue to perform the Bank Wizard check */

			/* Create the handle to perform the checks */
			BWIHandle checkHandle = new BWIHandle();
			
			/* Set up the country */
			checkHandle.setValue(BWIHandle.V_COUNTRY_CHECK,BWIHandle.C_UK);

			/* Perform the validation on the Sort Code and Account Number */
			checkHandle.setString(BWIUK.SORT_CODE, sortCode);
			checkHandle.setString(BWIUK.ACCOUNT_NO, accountNumber); 

			checkHandle.check();

			fatalConditions = new long[0];
			errorConditions = new long[0];

			fatalMessages = new string[0];
			errorMessages = new string[0];

			fatalErrorsExist = false;
			errorsExist = false;

			/* Check the response */
			fatalErrorsExist = IterateErrorConditions(checkHandle, BWIHandle.N_FATAL, out fatalConditions, out fatalMessages);
			errorsExist = IterateErrorConditions(checkHandle, BWIHandle.N_ERROR, out errorConditions, out errorMessages);

			/* If a fatal error has occurred, report it and return */
			if(fatalErrorsExist == true) 
			{
				responseElement = SetErrorMessage ("APPERR", defaultErrorCode.ToString(), "GetBankDetails", fatalMessages[0]);
				if(_logger.IsDebugEnabled) 
				{
					_logger.Debug(functionName + ": Response: " + responseElement.OuterXml);
				}
				checkHandle.free();	
				return responseElement.OuterXml;
			}

			/* If an error has been reported, process it */
			if(errorsExist == true) 
			{
				/* If the only error is BWI_UK_SORT_TRANSFERRED, search for the new Sort Code details */
				/* Otherwise set up the error message in the response element */

				responseElement = checkErrors(checkHandle, errorConditions, errorMessages, out transferred, out fatalErrorOccurred);

				/* Continue processing if no fatal error has been reported. */
				if (fatalErrorOccurred == true)
				{
					if(_logger.IsDebugEnabled) 
					{
						_logger.Debug(functionName + ": Response: " + responseElement.OuterXml);
					}
					checkHandle.free();	
					return responseElement.OuterXml;
				}
				
				if (responseElement != null)
				{
					errorString = responseElement.SelectSingleNode("//DESCRIPTION").InnerText;
				}
			}

			if ((responseElement == null) || (errorString != ""))
			{
				/* Return bank details */
				bankName = checkHandle.getString(BWIUK.FULL_NAME_1);
				branchName = checkHandle.getString(BWIUK.BRANCH_TITLE);
			
				/* Create the response */ 
				responseElement = doc.CreateElement ("RESPONSE");
				responseElement.SetAttribute ("TYPE","SUCCESS");

				XmlNode bankDetails = doc.CreateElement ("AccountDetails");

				XmlNode newNode = doc.CreateElement ("SortCode");
				newNode.InnerText = sortCode;
				bankDetails.AppendChild (newNode);

				newNode = doc.CreateElement ("AccountNumber");
				newNode.InnerText = accountNumber;
				bankDetails.AppendChild (newNode);

				newNode = doc.CreateElement ("BankName");
				newNode.InnerText = bankName;
				bankDetails.AppendChild (newNode);

				newNode = doc.CreateElement ("BranchName");
				newNode.InnerText = branchName;
				bankDetails.AppendChild (newNode);

				newNode = doc.CreateElement ("TransposedIndicator");
				newNode.InnerText = "0";
				if(checkHandle.isCondSet(BWIHandle.N_WARNING,BWIUK.TRANSPOSED)) 
				{
					newNode.InnerText = "1";
				}
				bankDetails.AppendChild (newNode);

				newNode = doc.CreateElement ("RedirectedIndicator");
				newNode.InnerText = "0";
				if(transferred == true) 
				{
					newNode.InnerText = "1";
				}
				bankDetails.AppendChild (newNode);

				XmlNode addressNode = doc.CreateElement("Address");
				newNode = doc.CreateElement ("Address1");
				newNode.InnerText = checkHandle.getString(BWIUK.ADDRESS_1);
				addressNode.AppendChild (newNode);

				newNode = doc.CreateElement ("Address2");
				newNode.InnerText = checkHandle.getString(BWIUK.ADDRESS_2);
				addressNode.AppendChild (newNode);

				newNode = doc.CreateElement ("Address3");
				newNode.InnerText = checkHandle.getString(BWIUK.ADDRESS_3);
				addressNode.AppendChild (newNode);

				newNode = doc.CreateElement ("Address4");
				newNode.InnerText = checkHandle.getString(BWIUK.ADDRESS_4);
				addressNode.AppendChild (newNode);

				newNode = doc.CreateElement ("PostCode");
				newNode.InnerText = checkHandle.getString(BWIUK.POST_CODE_MAJOR) + checkHandle.getString(BWIUK.POST_CODE_MINOR);
				addressNode.AppendChild (newNode);

				bankDetails.AppendChild (addressNode);

				/* If no serious error has occurred, continue to check the roll number if necessary */
				if (errorString.Length == 0) 
				{
					if(checkHandle.isCondSet(BWIHandle.N_WARNING,BWIUK.ROLL_NUM_REQUIRED)) 
					{
						if (doc.SelectSingleNode("//ValidateBankDetailsRequestType/NewRollNumber") == null)
						{
							errorString = "Collection account requires a roll account number";
						}
						else
						{
							/* Perform a check on the roll number */

							rollNumber = doc.SelectSingleNode("//ValidateBankDetailsRequestType/NewRollNumber").InnerText;

							checkHandle.setValue(BWIUK.RN_CHECK, 1);
							checkHandle.setValue(BWIUK.RN_CONV_BACS, 1);
							checkHandle.setString(BWIUK.RN, rollNumber); 

							checkHandle.check();

							if(checkHandle.isCondSet(BWIHandle.N_ERROR,BWIUK.RN_FORMAT)) 
							{
								errorString = "Roll number format is incorrect";
							}
							else if(checkHandle.isCondSet(BWIHandle.N_ERROR,BWIUK.RN_LENGTH))
							{
								errorString = "Roll number length is incorrect";
							}
						}
					}
				}

				if(errorString != "") 
				{
					newNode = doc.CreateElement ("BankWizardError");
					newNode.InnerText = errorString;
					bankDetails.AppendChild (newNode);
				}
	
				responseElement.AppendChild (bankDetails);
			}

			/* Free the handle after use, to release externally managed resources. */
			checkHandle.free();	
			if(_logger.IsDebugEnabled) 
			{
				_logger.Debug(functionName + ": Response: " + responseElement.OuterXml);
			}
			return responseElement.OuterXml;
		}

		private XmlElement checkErrors(BWIHandle h, long[] errorConditions, string[] errorMessages, out Boolean transferred, out Boolean fatalErrorOccurred)
		{
			transferred = false;
			fatalErrorOccurred = false;
			int index = 0;

			int numError = 0;
			string sortCodeReDir = null;
			XmlElement nodeResp = null;

			Boolean fatalErrorsExist = false;
			Boolean errorsExist = false;

			string[] newfatalMessages;
			string[] newerrorMessages;

			long[] newFatalConditions;
			long[] newErrorConditions;

			XmlDocument doc = new XmlDocument();
	
			foreach (long errorCon in errorConditions)
			{
				/* Store the first error (other than branch redirection) and keep a count of the errors */

				if (errorCon == BWIUK.SORT_TRANSFERRED)
				{
					sortCodeReDir = h.getString(BWIUK.REDIR_SC);
				}
				else if ((errorCon == BWIUK.ACCOUNT_NO_FORMAT) || 
					(errorCon == BWIUK.SORT_CODE_FORMAT) ||
					(errorCon == BWIUK.NO_BRANCH))
				{
					numError = numError + 1;
					if (numError == 1)
					{
						nodeResp = SetErrorMessage ("SUCCESS", defaultErrorCode.ToString(), "GetBankDetails", errorMessages[index]);
					}
				}
				else if (errorCon == BWIUK.MODULUS_CHECK)
				{
					numError = numError + 1;
					if (numError == 1)
					{
						nodeResp = SetErrorMessage ("SUCCESS", defaultErrorCode.ToString(), "GetBankDetails", "Account number may be invalid. Please check");
					}
				}
				index = index + 1;
			}
			
			/* If the branch has closed and the accounts redirected to another branch
			 * use BWI_UK_REDIR_SC to access this new branch if no other errors have been reported. */

			if ((sortCodeReDir != null) && (numError == 0))
			{	
				transferred = true;

				h.setString(BWIUK.SORT_CODE, sortCodeReDir);
				h.check();

				fatalErrorsExist = IterateErrorConditions(h, BWIHandle.N_FATAL, out newFatalConditions, out newfatalMessages);
				errorsExist = IterateErrorConditions(h, BWIHandle.N_ERROR, out newErrorConditions, out newerrorMessages);
				
				/* Return any errors */
				if(fatalErrorsExist == true) 
				{
					nodeResp = SetErrorMessage ("APPERR", defaultErrorCode.ToString(), "GetBankDetails", newfatalMessages[0]);
					fatalErrorOccurred = true;
				}
				else if (errorsExist == true)
				{
					index = 0;
					numError = 0;
					foreach (long errorCon in newErrorConditions)
					{
						if ((errorCon == BWIUK.ACCOUNT_NO_FORMAT) || 
							(errorCon == BWIUK.SORT_CODE_FORMAT) ||
							(errorCon == BWIUK.NO_BRANCH))
						{
							numError = numError + 1;
							if (numError == 1)
							{
								nodeResp = SetErrorMessage ("SUCCESS", defaultErrorCode.ToString(), "GetBankDetails", newerrorMessages[index]);
							}
						}
						else if (errorCon == BWIUK.MODULUS_CHECK)
						{
							numError = numError + 1;
							if (numError == 1)
							{
								nodeResp = SetErrorMessage ("SUCCESS", defaultErrorCode.ToString(), "GetBankDetails", "Account number may be invalid. Please check");
							}
						}

						index = index + 1;
					}
				}	
			}
			return nodeResp;
		}

		private Boolean IterateErrorConditions(BWIHandle h, int iSeverity, out long[] conditions, out string[] messages)
		{
			/* Returns a boolean value to indicate whether errors of the given type exist */

			Boolean errorExists = false;
			messages = new string[0];
			conditions = new long[0];
            
			/* Loop over the errors */
			while(h.next(iSeverity)) 
			{
				long cond = h.getValue(BWIHandle.V_NEXT);
				String mess = h.getString(BWIHandle.V_NEXT);

				/* Add the new values to the arrays */
				messages = AddToArray(messages, mess);
				conditions = AddToArray(conditions, cond);

				errorExists = true;
			}

			return errorExists;
		}

		public String[] AddToArray(String[] pOldArray, String pNewItem)
		{
			ArrayList originalList = new ArrayList(pOldArray);
			originalList.Add(pNewItem);
			return(String[])originalList.ToArray(typeof(String));
		}

		public long[] AddToArray(long[] pOldArray, long pNewItem)
		{
			ArrayList originalList = new ArrayList(pOldArray);
			originalList.Add(pNewItem);
			return(long[])originalList.ToArray(typeof(long));
		}

		private XmlElement SetErrorMessage (string errType, string errNo, string source, string errormsg)
		{
			XmlDocument doc = new XmlDocument();
			XmlElement nodeResp = doc.CreateElement ("RESPONSE");
			nodeResp.SetAttribute ("TYPE","");
			XmlNode errNode = doc.CreateElement ("ERROR");
			errNode.AppendChild (doc.CreateElement ("NUMBER"));
			errNode.AppendChild (doc.CreateElement ("SOURCE"));
			errNode.AppendChild (doc.CreateElement ("DESCRIPTION"));
			nodeResp.AppendChild (errNode);

			nodeResp.Attributes.GetNamedItem ("TYPE").Value = errType; 
			nodeResp.SelectSingleNode("//NUMBER").InnerText = errNo;
			nodeResp.SelectSingleNode("//DESCRIPTION").InnerText = errormsg;
			nodeResp.SelectSingleNode("//SOURCE").InnerText="BankWizard: " + source;
			return nodeResp;
		}
	}
}
