/*
--------------------------------------------------------------------------------------------
Workfile:			Introducer.cs
Copyright:			Copyright ©2006 Vertex Financial Services

Description:		Encapsulates an Omiga introducer
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		23/10/2006	EP2_22   Created
PSC		21/11/2006	EP2_138	 Add GetIntroducerPipeline
SR		25/01/2007	EP2_1005 modified validateBroker and AuthenticateFirm. Read the datareader till the  
							 EOF before checking the value of return param. Also, add the return params
							 to the command object first before any other parameters.
SR		26/01/2007	EP2_987  validateIntroducer, validateIntroducerAndGetFrims
SR		26/01/2007	EP2_987  ValidateIntroducerAndGetFirms. Add MESSAGETEXT to MESSAGE element
SR		05/02/2007	EP2_1163 added new method ValidatePasswordFormat()
SR		07/02/2007	EP2_103	 modified ValidateIntroducer to cover more error conditions
SR		12/02/2007	EP2_1295 
IK		27/03/2007	EP2_1162 ValidateIntroducer: log-in attempts no longer input, additional message 
--------------------------------------------------------------------------------------------
*/

using System;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using Vertex.Fsd.Omiga.Core;
using Vertex.Fsd.Omiga.omLogging;

namespace Vertex.Fsd.Omiga.omIntroducer
{
	public enum FirmType
	{
		ARFirm,
		PrincipalFirm
	}

	/// <summary>
	/// Encapsulates an Omiga introducer
	/// </summary>
	public class Introducer
	{
		private static readonly Logger _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		/// <summary>
		/// Creates an instance of the Introducer class
		/// </summary>
		public Introducer()
		{
		}

		/// <summary>
		/// Validates an introducer's login credentials
		/// </summary>
		/// <param name="userName">Introducer's user name</param>
		/// <param name="password">Introducer's password</param>
		/// <param name="loginAttempts">Number of login attempts</param>
		/// <returns>The corresponding introducer id if the credentials are valid</returns>
		public string ValidateIntroducer(string userName, string password, int loginAttempts, out short pwdExpiryin5DaysInd, out byte changePwdInd)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			string introducerId ;
			int returnCode = 0;
			
			pwdExpiryin5DaysInd = -1; // holds number of days password is valid.  if number of days is <=5, else -1
			changePwdInd = 0;

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("USP_VALIDATEINTRODUCER", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;
					sqlCommand.Parameters.Add("@pUserName", SqlDbType.NVarChar, userName.Length).Value = userName;
					sqlCommand.Parameters.Add("@pPassword", SqlDbType.NVarChar, password.Length).Value = password;
					
					SqlParameter returnValue = new SqlParameter("@ReturnValue", SqlDbType.Int);
					returnValue.Direction = ParameterDirection.ReturnValue;
					sqlCommand.Parameters.Add(returnValue);
					// SR : EP2_987
					SqlParameter pIntroducerId = new SqlParameter("@pIntroducerId", SqlDbType.NChar, 28);
					pIntroducerId.Direction = ParameterDirection.Output;
					sqlCommand.Parameters.Add(pIntroducerId);

					SqlParameter paramChangePasswordInd = new SqlParameter("@pChangePasswordIndicator", SqlDbType.SmallInt ,1);
					paramChangePasswordInd.Direction = ParameterDirection.Output;
					sqlCommand.Parameters.Add(paramChangePasswordInd);

					SqlParameter paramPasswordExpiryIn5Days = new SqlParameter("@pPasswordExpiryIn5Days", SqlDbType.SmallInt, 1);
					paramPasswordExpiryIn5Days.Direction = ParameterDirection.Output;
					sqlCommand.Parameters.Add(paramPasswordExpiryIn5Days);
					// SR : EP2_987 - End
					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);

					}
					// SR : EP2_987
					sqlCommand.ExecuteScalar();
					returnCode = (int)returnValue.Value;
					introducerId = pIntroducerId.Value.ToString().Trim();
					if(!(paramChangePasswordInd.Value is System.DBNull)) changePwdInd =  Convert.ToByte(paramChangePasswordInd.Value); 
					if(!(paramPasswordExpiryIn5Days.Value is System.DBNull)) pwdExpiryin5DaysInd = Convert.ToInt16(paramPasswordExpiryIn5Days.Value);
					// SR : EP2_987 - End
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Stored procedure return code {1}", functionName, returnCode);
				}

				int messageNumber = 0;
				
				if (returnCode != 0)
				{
					switch (returnCode)
					{
						case 1:		// Exceeded login attempts
							messageNumber = 7801;
							break;
						case 2:		// User name and password combination not found
							messageNumber = 7802;
							break;
						case 3:		// Password expired
							messageNumber = 7803;
							break;
						case 4:		//Inactive user  //SR : EP2_103
							messageNumber = 7804;
							break;
						case 5:		// Outside working hours	//SR : EP2_103
							messageNumber = 7805;
							break;
						case 6:		// User Name (email address) not registered	//IK : EP2_1162
							messageNumber = 7826;
							break;
					}

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Throwing OmigaException {1}", functionName, messageNumber);
					}

					throw new OmigaException(messageNumber);
				}

					
				return introducerId;
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Gets the firms linked to an introducer
		/// </summary>
		/// <param name="introducerId">The introducer's id</param>
		/// <returns>An XML string containing the introducer's firms</returns>
		public string GetIntroducerFirms(string introducerId)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			XmlReader introducerFirms = null;
			
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("USP_GETINTRODUCERUSERFIRMS", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;
					sqlCommand.Parameters.Add("@pIntroducerId", SqlDbType.NVarChar, introducerId.Length).Value = introducerId;
					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);
					}

					StringBuilder outputString = new StringBuilder();

					introducerFirms = sqlCommand.ExecuteXmlReader();

					introducerFirms.Read();

					while (introducerFirms.ReadState != ReadState.EndOfFile)
					{
						outputString.Append(introducerFirms.ReadOuterXml());
					}


					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Returned firms: {1}.", functionName, outputString.ToString());
					}

					return outputString.ToString();
				}
			}
			finally
			{
				if (introducerFirms != null)
				{
					introducerFirms.Close();
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Validates an introducer's credentials and returns their linked firms
		/// </summary>
		/// <param name="userName">Introducer's user name</param>
		/// <param name="password">Introducer's password</param>
		/// <param name="loginAttempts">Number of login attempts</param>
		/// <returns>An XML string containing the introducer's firms</returns>
		public string ValidateIntroducerAndGetFirms(string userName, string password, int loginAttempts)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Validating introducer credentials.", functionName);
				}

				short pwdExpiryIn5DaysInd ; // holds number of days password is valid. if number of days is <=5, else -1
				byte changePwdInd ;
				string introduceId = ValidateIntroducer(userName, password, loginAttempts, out pwdExpiryIn5DaysInd, out changePwdInd);

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Getting introducer firms.", functionName);
				}
				// SR : EP2_987
				string respIntroducerFirms = GetIntroducerFirms(introduceId);
				XmlDocumentEx responseXmlDoc = new XmlDocumentEx();
				responseXmlDoc.LoadXml(respIntroducerFirms);
				XmlElementEx responseElem  = (XmlElementEx)responseXmlDoc.CreateElement("RESPONSE");				
				if(pwdExpiryIn5DaysInd != -1) 
				{
					responseElem.SetAttribute("TYPE", "WARNING");
					
					XmlElementEx messageElem = (XmlElementEx)responseXmlDoc.CreateElement("MESSAGE");
					XmlElementEx messageTextElem = (XmlElementEx)responseXmlDoc.CreateElement("MESSAGETEXT");
					messageTextElem.InnerXml = "Your Password will expire in " + Convert.ToString(pwdExpiryIn5DaysInd, 10) + " days.";
					messageElem.AppendChild(messageTextElem);
					XmlElementEx messageTypeElem = (XmlElementEx)responseXmlDoc.CreateElement("MESSAGETYPE");
					messageTypeElem.InnerXml = "Warning";
					messageElem.AppendChild(messageTypeElem);
					
					responseElem.AppendChild (messageElem);
				}
				else responseElem.SetAttribute("TYPE", "SUCCESS");

				if(changePwdInd	== 1)
				{
					XmlElementEx introducerUserElem = (XmlElementEx)responseXmlDoc.CreateElement("INTRODUCERUSER");
					introducerUserElem.SetAttribute("CHANGEPASSWORDINDICATOR", "1");
					responseElem.AppendChild (introducerUserElem);
				}
				
				responseElem.AppendChild(responseXmlDoc.DocumentElement);
				return responseElem.OuterXml;
				// SR : EP2_987 - End
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Authenticates an introducer against a firm 
		/// </summary>
		/// <param name="introducerId">The inttroducer's id</param>
		/// <param name="firmId">The id of the firm to be checked</param>
		/// <returns>An XML string containing the list of permissions for the supplied firm</returns>
		public string AuthenticateFirm(string introducerId, string firmId, FirmType firmType)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			XmlReader firmPermissions = null;
			StringBuilder outputString = new StringBuilder();
			int returnCode = 0;
		
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("USP_AUTHENTICATEFIRM", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;

					SqlParameter returnValue = new SqlParameter("@ReturnValue", SqlDbType.Int);
					returnValue.Direction = ParameterDirection.ReturnValue;
					sqlCommand.Parameters.Add(returnValue);
					sqlCommand.Parameters.Add("@pIntroducerId", SqlDbType.NVarChar, introducerId.Length).Value = introducerId;
					sqlCommand.Parameters.Add("@pFirmId", SqlDbType.NVarChar, firmId.Length).Value = firmId;
	
					switch (firmType)
					{
						case FirmType.ARFirm:
							sqlCommand.Parameters.Add("@pFirmType", SqlDbType.NVarChar, 1).Value = "A";
							break;
						case FirmType.PrincipalFirm:
							sqlCommand.Parameters.Add("@pFirmType", SqlDbType.NVarChar, 1).Value = "P";
							break;
					}					

					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);
					}

					firmPermissions = sqlCommand.ExecuteXmlReader();

					if (firmPermissions.Read())
					{
						while (firmPermissions.ReadState != ReadState.EndOfFile)
						{
							outputString.Append(firmPermissions.ReadOuterXml());
						}

						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Returned permissions: {1}.", functionName, outputString.ToString());
						}
					}
					returnCode = (int)returnValue.Value;
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Stored procedure return code {1}", functionName, returnCode);
					}
				}

				if (returnCode != 0)
				{
					int messageNumber = 0;

					switch (returnCode)
					{
						case 1:		// Inactive user		
							messageNumber = 7804;
							break;
						case 2:		// Outside working hours		
							messageNumber = 7805;
							break;
						case 3:		// Inactive user role
							messageNumber = 7806;
							break;
						case 4:		// Not a web user
							messageNumber = 7807;
							break;
						case 5:		// Firm not found
							messageNumber = 7808;
							break;
						case 6:		// Firm is blacklisted
							messageNumber = 7809;
							break;
						case 7:		// Firm not applicable
							messageNumber = 7810;
							break;
						case 8:		// Unit is not active
							messageNumber = 7811;
							break;
					}
					
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Throwing OmigaException {1}", functionName, messageNumber);
					}

					throw new OmigaException(messageNumber);
				}

				return outputString.ToString();
			}
			finally
			{
				if (firmPermissions != null)
				{
					firmPermissions.Close();
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		/// <summary>
		/// Validates a brokers credentials
		/// </summary>
		/// <param name="userName">The broker's user name</param>
		/// <param name="FsaRef">The FSA Reference</param>
		/// <returns>An XML string containing a list of permissions for the borkers associated firms</returns>
		public string ValidateBroker(string userName, string FsaRef)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			XmlReader firmPermissions = null;
			StringBuilder outputString = new StringBuilder();
			int returnCode = 0;
		
			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("USP_VALIDATEBROKER", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;

					SqlParameter returnValue = new SqlParameter("@ReturnValue", SqlDbType.Int);
					returnValue.Direction = ParameterDirection.ReturnValue;
					sqlCommand.Parameters.Add(returnValue);

					if (userName != null && userName.Length > 0)
					{
						sqlCommand.Parameters.Add("@pUserName", SqlDbType.NVarChar, userName.Length).Value = userName;
					}
					else
					{
						sqlCommand.Parameters.Add("@pUserName", SqlDbType.NVarChar, 1).Value = DBNull.Value;
					}

					if (FsaRef != null && FsaRef.Length > 0)
					{
						sqlCommand.Parameters.Add("@pFSARef", SqlDbType.NVarChar, FsaRef.Length).Value = FsaRef;
					}
					else
					{
						sqlCommand.Parameters.Add("@pFSARef", SqlDbType.NVarChar, 1).Value = DBNull.Value;
					}

					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);
					}

					firmPermissions = sqlCommand.ExecuteXmlReader();				

					if (firmPermissions.Read())
					{
						while (firmPermissions.ReadState != ReadState.EndOfFile)
						{
							outputString.Append(firmPermissions.ReadOuterXml());
						}

						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: Returned permissions: {1}.", functionName, outputString.ToString());
						}
					}

					returnCode = (int)returnValue.Value;
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Stored procedure return code {1}", functionName, returnCode);
					}
				}
					
				int messageNumber = 0;			
				if (returnCode != 0)
				{
					switch (returnCode)
					{
						case 1:		// User name unknown		
							messageNumber = 7812;
							break;
					}
				
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Throwing OmigaException {1}", functionName, messageNumber);
					}

					throw new OmigaException(messageNumber);								
				}

				return outputString.ToString();
			}
			finally
			{
				if (firmPermissions != null)
				{
					firmPermissions.Close();
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}
		
		// PSC 21/11/2006 EP2_138 - Start
		/// <summary>
		/// Retrieves a list of pipeline cases linked to an introducer
		/// </summary>
		/// <param name="userId">The user id for the user performing the search</param>
		/// <param name="unitId">The unit id for the user performing the search</param>
		/// <param name="introducerId">The introducer id for the user performing the search</param>
		/// <param name="includePrevious">Whether cases owned previously should be included in the search</param>
		/// <param name="includeCancelled">Whether cancelled cases should be included in the search</param>
		/// <param name="includeFirm">Whether cases for the firm should be included in the search</param>
		/// <returns>An XML string containing a list of the introducer's pipeline cases </returns>
		public string GetIntroducerPipeline(string userId, 
											string unitId, 
											string introducerId, 
											bool includePrevious, 
											bool includeCancelled, 
											bool includeFirm)
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;
			string pipelineCases = string.Empty;
			
			XmlReader introducerPipeline = null;
			XPathDocument pipelineCasesDocument = null;
			StringWriter responseWriter = new StringWriter();
			XmlTextWriter responseXmlwriter = new XmlTextWriter(responseWriter);

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}

			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("USP_GETINTRODUCERPIPELINE", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;

					sqlCommand.Parameters.Add("@pUserId", SqlDbType.NVarChar, userId.Length).Value = userId;
					sqlCommand.Parameters.Add("@pUnitId", SqlDbType.NVarChar, unitId.Length).Value = unitId;
					sqlCommand.Parameters.Add("@pIntroducerId", SqlDbType.NVarChar, introducerId.Length).Value = introducerId;
					
					SqlParameter parameter = sqlCommand.Parameters.Add("@pIncludePrevious", SqlDbType.TinyInt);
					parameter.Value = includePrevious? 1 : 0;

					parameter = sqlCommand.Parameters.Add("@pIncludeCancelled", SqlDbType.TinyInt);
					parameter.Value = includeCancelled? 1 : 0;

					parameter = sqlCommand.Parameters.Add("@pIncludeFirm", SqlDbType.TinyInt);
					parameter.Value = includeFirm? 1 : 0;

					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);
					}

					introducerPipeline = sqlCommand.ExecuteXmlReader();


					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Returned from stored procedure", functionName);
					}

					pipelineCasesDocument = new XPathDocument(introducerPipeline);
					responseXmlwriter.Formatting = Formatting.None;

					string XsltFilePath = GetXMLPath() + "\\IntroducerPipeLineResponse.xslt";
					XslTransform xslt = new XslTransform();
					
					try
					{
						xslt.Load(XsltFilePath);
					}
					catch (System.IO.FileNotFoundException exp)
					{
						throw new OmigaException("Unable to apply transform to results. Xslt file " + XsltFilePath + " not found." , exp);
					}
					xslt.Transform(pipelineCasesDocument, null, responseXmlwriter, null);
					pipelineCases = responseWriter.ToString();
					responseWriter.Close();
					responseXmlwriter.Close();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Returned pipeline: {1}.", functionName, pipelineCases);
					}

					return pipelineCases;
				}
			}
			finally
			{
				if (introducerPipeline != null)
				{
					introducerPipeline.Close();
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}

		// SR 01/02/2007 EP2_1163 - Start
		/// <summary>
		/// validate Password format based on the setting in global parameters
		/// </summary>
		/// <param name="password">password string to be validated</param>
		/// <param name="emailAddress">Email Address of the user</param>
		/// <returns></returns>
		public bool ValidatePasswordFormat(string password, string emailAddress, out int message) 
		{			
			return ValidatePasswordFormat(password, emailAddress, out message, String.Empty) ;
		}

		/// <summary>
		/// validate Password format based on the setting in global parameters
		/// </summary>
		/// <param name="password">password string to be validated</param>
		/// <param name="emailAddress">Email Address of the user</param>
		/// <param name="userId">userId for which the password record is to be created.</param>
		/// <returns></returns>
		//SR EP2_1163 : new method ValidatePasswordFormat
		public bool ValidatePasswordFormat(string password, string emailAddress, out int message, string userId) 
		{
			string functionName = MethodInfo.GetCurrentMethod().Name;		

			if (_logger.IsDebugEnabled)
			{
				_logger.DebugFormat("{0}: Starting.", functionName);
			}			
			
			try
			{
				using (SqlConnection sqlConnection = new SqlConnection(Vertex.Fsd.Omiga.Core.Global.DatabaseConnectionString))
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Setting up parameters.", functionName);
					}

					SqlCommand sqlCommand = new SqlCommand("usp_GetParamsForPasswordValidation", sqlConnection);
					sqlCommand.CommandType = CommandType.StoredProcedure;
					string encryptedPwd = Encrypt(password);
					sqlCommand.Parameters.Add("@pPassword", SqlDbType.NVarChar, encryptedPwd.Length).Value = encryptedPwd; //SR EP2_1295
					sqlCommand.Parameters.Add("@puserId", SqlDbType.NVarChar, userId.Length).Value = userId;					
					SqlParameter pPasswordAlreadyUsed = createOutSQLParam(sqlCommand, "@pPasswordAlreadyUsed" , SqlDbType.SmallInt);
					SqlParameter pPasswordAlphaNumeric = createOutSQLParam(sqlCommand, "@pPasswordAlphaNumeric", SqlDbType.SmallInt);
					SqlParameter pPasswordCharacterDuplication = createOutSQLParam(sqlCommand, "@pPasswordCharacterDuplication", SqlDbType.SmallInt);
					SqlParameter pPasswordLowerAlpha = createOutSQLParam(sqlCommand, "@pPasswordLowerAlpha", SqlDbType.SmallInt);
					SqlParameter pPasswordMaximumLength = createOutSQLParam(sqlCommand, "@pPasswordMaximumLength", SqlDbType.SmallInt); 
					SqlParameter pPasswordMinimumLength = createOutSQLParam(sqlCommand, "@pPasswordMinimumLength", SqlDbType.SmallInt); 
					SqlParameter pPasswordNoUserID = createOutSQLParam(sqlCommand, "@pPasswordNoUserID", SqlDbType.SmallInt);
					SqlParameter pPasswordNumeric = createOutSQLParam(sqlCommand, "@pPasswordNumeric", SqlDbType.SmallInt);
					SqlParameter pPasswordSpecialCharacters = createOutSQLParam(sqlCommand, "@pPasswordSpecialCharacters", SqlDbType.SmallInt);
					SqlParameter pPasswordUpperAlpha = createOutSQLParam(sqlCommand, "@pPasswordUpperAlpha" , SqlDbType.SmallInt);
					sqlConnection.Open();

					if (_logger.IsDebugEnabled)
					{
						_logger.DebugFormat("{0}: Calling stored procedure.", functionName);
					}

					sqlCommand.ExecuteScalar();
					int messageNumber = 0;
					
					if(userId != "" && Convert.ToInt16(pPasswordAlreadyUsed.Value)==1)
					{
						messageNumber = 7825;  //SR EP2_1295
					}
					else if(Convert.ToInt16(pPasswordNoUserID.Value) == 1 && emailAddress == password)
					{
						messageNumber = 7821;
					}
					else if(password.Length >  Convert.ToInt16(pPasswordMaximumLength.Value))
					{
						messageNumber = 7819 ;
					}
					else if(password.Length < Convert.ToInt16(pPasswordMinimumLength.Value))
					{
						messageNumber = 7820 ;
					}
					else
					{
						if(Convert.ToInt16(pPasswordAlphaNumeric.Value) == 1)
						{
							if(!Regex.IsMatch(password, "[0-9]"))
							{
								messageNumber = 7816 ;  // must contain numerics
							}
							else
							{	// must contain alphabets 
								if(!Regex.IsMatch(password, "[a-zA-Z]")) messageNumber = 7816 ;  
							}
						}
						if(messageNumber==0 && Convert.ToInt16(pPasswordNumeric.Value) != 1)
						{	// must not contain numerics 
							if(Regex.IsMatch(password, "[0-9]")) messageNumber = 7822 ;
						}
						if(messageNumber==0 && Convert.ToInt16(pPasswordLowerAlpha.Value) != 1)
						{	// must not contain lower case alphabets 
							if(Regex.IsMatch(password, "[a-z]")) messageNumber = 7818 ; 
						}
						if(messageNumber==0 && Convert.ToInt16(pPasswordUpperAlpha.Value) != 1)
						{	// must not contain upper case alphabets 
							if(Regex.IsMatch(password, "[A-Z]")) messageNumber = 7824 ;
						}
						if(messageNumber==0 && Convert.ToInt16(pPasswordSpecialCharacters.Value) != 1)
						{	// must not contain special characters
							if(Regex.IsMatch(password, "[^a-zA-Z0-9]")) messageNumber = 7823 ;
						}
						if(messageNumber==0 && Convert.ToInt16(pPasswordCharacterDuplication.Value) == 1) //SR EP2_1295
						{	//must not contain duplicate characters
							if(HasDuplicateChars(password, true)) messageNumber = 7817 ;
						}
					}

					message = messageNumber ;
					if(messageNumber != 0)
					{						
						if (_logger.IsDebugEnabled)
						{
							_logger.DebugFormat("{0}: returning OmigaException {1}", functionName, messageNumber);
						}
						return false;
					}
					else return true ;

				}	
			}
			finally
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.DebugFormat("{0}: Exiting.", functionName);
				}
			}
		}
		
		private SqlParameter createOutSQLParam(SqlCommand sqlCommand, string paramName, SqlDbType type)
		{
			SqlParameter param = new SqlParameter(paramName, type);
			param.Direction = ParameterDirection.Output;
			sqlCommand.Parameters.Add(param);
			return param;
		}

		private bool HasDuplicateChars(string strToCheck)
		{
			return HasDuplicateChars(strToCheck, false);

		}
		private bool HasDuplicateChars(string strToCheck, bool caseSensitive)
		{	// functionality here is same that in omORG component
			// if any of the character repeats consecutivel, raise an error
			// a12b12 - valid.   a112b2 - invalid
			string currChar, lastAlphaChar, lastDigitChar ;
			int numAlphas, numDigits ;
			int lenOfString = strToCheck.Length ;
			int index = 1;
			bool hasDuplicateChars, hasDuplicateAlphas, hasDuplicateDigits ;
			
			hasDuplicateChars = false;
			hasDuplicateAlphas = false;
			hasDuplicateDigits = false ;
			numAlphas = 0;
			numDigits = 0;
			lastAlphaChar = "";
			lastDigitChar = "";

			do 
			{
				currChar = strToCheck.Substring(index-1,1).ToString();
				if( isAlpha(currChar))
				{
					++numAlphas;
					if(!caseSensitive){ currChar = currChar.ToUpper(); }
					if(numAlphas > 1)
					{
						if(currChar == lastAlphaChar) { hasDuplicateAlphas = true; }
					}

					lastAlphaChar = currChar ;
				}
				else if(isDigit(currChar))
				{
					++numDigits;
					if(numDigits > 1)
					{
						if(currChar == lastDigitChar) { hasDuplicateDigits = true; }
					}
					lastDigitChar = currChar;
				}
				index = index + 1;

			}while(hasDuplicateAlphas==false &&  hasDuplicateDigits==false && index <= lenOfString);	

			if((hasDuplicateAlphas && numAlphas > 1) || (hasDuplicateDigits && numDigits > 1))
				{hasDuplicateChars = true; }

			return hasDuplicateChars ;
		}
		
		private bool isAlpha(string c)
		{
			return Regex.IsMatch(c, "[a-zA-Z]");
		}
		private bool isDigit(string c)
		{
			return Regex.IsMatch(c, "[0-9]");
		}

		public string Encrypt(string plainText)
		{
			int inputLength = plainText.Length;
			int charPosition = 0;
			string cypherText = string.Empty;
			byte increment;
			byte char1;
			byte char2;
			const byte asciiTilde = 126;

			byte[] random = new byte[30] {23, 12, 33, 40, 17, 31, 23, 25, 12, 43, 44, 43, 48, 43, 13, 35, 17, 23, 17, 37, 21, 34, 10, 38, 11, 12, 41, 21, 43, 23};
			byte randomIndex = 0;

			while (charPosition < inputLength)
			{
				increment = random[randomIndex++];
				char1 = (byte) plainText.ToCharArray(charPosition++, 1)[0];
				char2 = (byte) (char1 + increment);

				if (char2 >= asciiTilde)
				{
					cypherText += '~';
					char2 = (byte) (char1 - increment);
				}

				cypherText += (char) char2;
			}
			return cypherText;
		}
		private string GetXMLPath()
		{
			string thePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

			if (thePath.ToUpper().IndexOf("DLLDOTNET") >= 0)
			{
				thePath = thePath.Replace("\\DLLDOTNET", "\\XML");
			}
			else if (thePath.ToUpper().IndexOf("OMIGAWEBSERVICES") >= 0)
			{
				int position = thePath.ToUpper().IndexOf("OMIGAWEBSERVICES");
				thePath = thePath.Substring(0, position) + @"XML";
			}
			return(thePath);
		}
		// PSC 21/11/2006 EP2_138 - End
	}
}
