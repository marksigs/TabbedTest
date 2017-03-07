/*
--------------------------------------------------------------------------------------------
Workfile:			ErrAssistException.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Replacement for errAssist.bas and ErrAssist.cls.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		09/05/2007	First version.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Contains constants describing the different type of Omiga messages on the MESSAGE table.
	/// </summary>
	public enum OmigaMessageType
	{
		/// <summary>
		/// Undefined.
		/// </summary>
		None,
		/// <summary>
		/// A prompt.
		/// </summary>
		Prompt,
		/// <summary>
		/// An information message.
		/// </summary>
		Information,
		/// <summary>
		/// A warning message.
		/// </summary>
		Warning,
		/// <summary>
		/// An error (exception).
		/// </summary>
		Error
	}

	/// <summary>
	/// Contains constants describing the different type of action the 
	/// <see cref="ErrAssistException"/> constructor should take for any current COM+ context.
	/// </summary>
	public enum TransactionAction
	{
		/// <summary>
		/// No action.
		/// </summary>
		None,
		/// <summary>
		/// Sets the consistent bit and the done bit to true in the current COM+ context (if there is one). 
		/// </summary>
		SetComplete,
		/// <summary>
		/// Sets the consistent bit and the done bit to true in the current COM+ context (if there is one). 
		/// </summary>
		SetAbort,
		/// <summary>
		/// If the <see cref="ErrAssistException"/> object represents an Omiga system error then SetAbort, otherwise SetComplete.
		/// </summary>
		SetOnErrorType,
		/// <summary>
		/// If the <see cref="ErrAssistException"/> object represents an Omiga system error then SetComplete, otherwise do nothing.
		/// </summary>
		SetCompleteOnSystemError,
		/// <summary>
		/// If the <see cref="ErrAssistException"/> object represents an Omiga system error then SetAbort, otherwise do nothing.
		/// </summary>
		SetAbortOnSystemError,
	}

	/// <summary>
	/// The ErrAssistException class replaces the VB6 ErrAssist.cls errAssist.bas modules.
	/// </summary>
	public class ErrAssistException : Exception
	{
		private OMIGAERROR _omigaError = OMIGAERROR.UnspecifiedError;
		private int _omigaErrorNumber;
		private string _omigaMessageText;
		private OmigaMessageType _omigaMessageType = OmigaMessageType.None;
		private string _id;
		private string _errorSource;
		private string _exceptionType;
		private bool _logged;
		private TransactionAction _transactionAction = TransactionAction.None;

		/// <summary>
		/// Constructs a new ErrAssistException exception from an existing exception.
		/// </summary>
		/// <param name="exception">The existing exception.</param>
		/// <remarks>
		/// <para>
		/// If <i>exception</i> is an ErrAssistException then its fields are copied to this instance.
		/// </para>
		/// <para>
		/// If this is an Omiga system exception, then exception details are written to the event log.
		/// </para>
		/// </remarks>
		public ErrAssistException(Exception exception)
			: this(exception, TransactionAction.None, null)
		{
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception from an existing exception.
		/// </summary>
		/// <param name="exception">The existing exception.</param>
		/// <param name="xmlResponseElement">An xml element to which any warnings should be added.</param>
		/// <remarks>
		/// <para>
		/// If <i>exception</i> is an ErrAssistException then its fields are copied to this instance.
		/// </para>
		/// <para>
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga system error, then if there is a COM+ transaction it will be set to 
		/// aborted. 
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga application error, then if there is a COM+ transaction it will be set 
		/// to completed. 
		/// </para>
		/// <para>
		/// If this is an Omiga system exception, then exception details are written to the event log. 
		/// If this is an Omiga warning, then the warning 
		/// details are added to the <paramref name="xmlResponseElement"/> parameter.
		/// </para>
		/// </remarks>
		public ErrAssistException(Exception exception, XmlElement xmlResponseElement)
			: this(exception, TransactionAction.None, xmlResponseElement)
		{
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception from an existing exception.
		/// </summary>
		/// <param name="exception">The existing exception.</param>
		/// <param name="transactionAction">Indicates the action to take for any COM+ transaction.</param>
		/// <remarks>
		/// <para>
		/// If <i>exception</i> is an ErrAssistException then its fields are copied to this instance.
		/// </para>
		/// <para>
		/// If <paramref name="setTransactionStatus"/> is true and the exception is an Omiga system error, 
		/// then if there is a COM+ transaction it will be set to aborted. 
		/// If <paramref name="setTransactionStatus"/> is true and the exception is an Omiga 
		/// application error, then if there is a COM+ transaction it will be set to completed. 
		/// </para>
		/// </remarks>
		public ErrAssistException(Exception exception, TransactionAction transactionAction)
			: this(exception, transactionAction, null)
		{
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception from an existing exception.
		/// </summary>
		/// <param name="exception">The existing exception.</param>
		/// <param name="transactionAction">Indicates the action to take for any COM+ transaction.</param>
		/// <param name="xmlResponseElement">An xml element to which any warnings should be added.</param>
		/// <remarks>
		/// <para>
		/// If <i>exception</i> is an ErrAssistException then its fields are copied to this instance.
		/// </para>
		/// <para>
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga system error, then if there is a COM+ transaction it will be set to 
		/// aborted. 
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga application error, then if there is a COM+ transaction it will be set 
		/// to completed. 
		/// </para>
		/// <para>
		/// If this is an Omiga system exception, then exception details are written to the event log. 
		/// If this is an Omiga warning, then the warning 
		/// details are added to the <paramref name="xmlResponseElement"/> parameter.
		/// </para>
		/// </remarks>
		public ErrAssistException(Exception exception, TransactionAction transactionAction, XmlElement xmlResponseElement)
			: base(exception.Message, exception)
		{
			if (exception == null)
			{
				throw new ArgumentNullException("exception");
			}

			ErrAssistException errAssistException = exception as ErrAssistException;
			if (errAssistException != null)
			{
				_omigaError = errAssistException.OmigaError;
				_omigaErrorNumber = errAssistException.OmigaErrorNumber;
				_omigaMessageText = errAssistException._omigaMessageText;
				_omigaMessageType = errAssistException._omigaMessageType;
				_id = errAssistException.Id;
				_errorSource = errAssistException.ErrorSource;
				_exceptionType = errAssistException._exceptionType;
				_logged = errAssistException._logged;
				_transactionAction = errAssistException._transactionAction;
				// Do no log as will have already been logged.
			}
			else
			{
				_id = GuidAssist.CreateGUID();
				_errorSource = exception.StackTrace;
				_exceptionType = GetType().ToString();
				LogError();
			}

			Source = exception.Source;

			if (xmlResponseElement != null && IsWarning())
			{
				AddWarning(xmlResponseElement);
			}

			if (_transactionAction != transactionAction)
			{
				SetTransactionState(transactionAction);
				_transactionAction = transactionAction;
			}
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception with a specified description.
		/// </summary>
		/// <param name="description">The exception description.</param>
		/// <remarks>
		/// <para>
		/// The <i>description</i> parameter is assigned to the base exception Message property.
		/// </para>
		/// <para>
		/// The exception details are written to the event log.
		/// </para>
		/// </remarks>
		public ErrAssistException(string description)
			: base(description)
		{
			_id = GuidAssist.CreateGUID();
			_errorSource = MakeStackTrace();
			_exceptionType = GetType().ToString();
			LogError();
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception with a specified description and inner exception.
		/// </summary>
		/// <param name="description">The exception description.</param>
		/// <param name="innerException">The inner exception.</param>
		/// <remarks>
		/// <para>
		/// The <i>description</i> parameter is assigned to the base exception Message property. 
		/// The <i>innerException</i> parameter is assigned to the base exception InnerException property.
		/// </para>
		/// <para>
		/// The exception details are written to the event log.
		/// </para>
		/// </remarks>
		public ErrAssistException(string description, Exception innerException)
			: base(description, innerException)
		{
			_id = GuidAssist.CreateGUID();
			_errorSource = MakeStackTrace();
			_exceptionType = GetType().ToString();
			LogError();
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception for a specified OMIGAERROR and message.
		/// </summary>
		/// <param name="omigaError">One of the constants in the <see cref="OMIGAERROR"/> enum.</param>
		/// <param name="parameters">Any additional parameters to be included in the message.</param>
		/// <remarks>
		/// <para>
		/// <i>omigaError</i> is used to read a message from the database, 
		/// where MESSAGE.MESSAGENUMBER = <i>omigaError</i>. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// <i>parameters</i> is optional. If it is defined it can consist of 1 or more parameters. 
		/// The parameters are used to substitute for any %s placeholders in the message for 
		/// this ErrAssistException object. If any parameters are any left over from this substitution 
		/// then the first one is appended to the end of the message. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// If <i>omigaError</i> is a system error (as defined by <see cref="IsSystemError"/>), 
		/// then the exception details are written to the event log.
		/// </para>
		/// </remarks>
		public ErrAssistException(OMIGAERROR omigaError, params string[] parameters) 
			: this((int)omigaError, TransactionAction.None, parameters)
		{
			_omigaError = omigaError;
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception for a specified OMIGAERROR and message.
		/// </summary>
		/// <param name="omigaError">One of the constants in the <see cref="OMIGAERROR"/> enum.</param>
		/// <param name="transactionAction">Indicates the action to take for any COM+ transaction.</param>
		/// <param name="parameters">Any additional parameters to be included in the message.</param>
		/// <remarks>
		/// <para>
		/// <i>omigaError</i> is used to read a message from the database, 
		/// where MESSAGE.MESSAGENUMBER = <i>omigaError</i>. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// <i>parameters</i> is optional. If it is defined it can consist of 1 or more parameters. 
		/// The parameters are used to substitute for any %s placeholders in the message for 
		/// this ErrAssistException object. If any parameters are any left over from this substitution 
		/// then the first one is appended to the end of the message. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// If <i>omigaError</i> is a system error (as defined by <see cref="IsSystemError"/>), 
		/// then the exception details are written to the event log.
		/// </para>
		/// <para>
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga system error, then if there is a COM+ transaction it will be set to 
		/// aborted. 
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga application error, then if there is a COM+ transaction it will be set 
		/// to completed. 
		/// </para>
		/// </remarks>
		public ErrAssistException(OMIGAERROR omigaError, TransactionAction transactionAction, params string[] parameters)
			: this((int)omigaError, transactionAction, parameters)
		{
			_omigaError = omigaError;
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception for a specified an error number and message.
		/// </summary>
		/// <param name="errorNumber">An error number.</param>
		/// <param name="parameters">Any additional parameters to be included in the message.</param>
		/// <remarks>
		/// <para>
		/// <i>errorNumber</i> is converted to an <see cref="OMIGAERROR"/> constant, and this is 
		/// used to read a message from the database, 
		/// where MESSAGE.MESSAGENUMBER = <i>omigaError</i>. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// <i>parameters</i> is optional. If it is defined it can consist of 1 or more parameters. 
		/// The parameters are used to substitute for any %s placeholders in the message for 
		/// this ErrAssistException object. If any parameters are any left over from this substitution 
		/// then the first one is appended to the end of the message. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// If <i>omigaError</i> is a system error (as defined by <see cref="IsSystemError"/>), 
		/// then the exception details are written to the event log.
		/// </para>
		/// </remarks>
		public ErrAssistException(int errorNumber, params string[] parameters) : this(errorNumber, TransactionAction.None, parameters)
		{
		}

		/// <summary>
		/// Constructs a new ErrAssistException exception for a specified an error number and message.
		/// </summary>
		/// <param name="errorNumber">An error number.</param>
		/// <param name="transactionAction">Indicates the action to take for any COM+ transaction.</param>
		/// <param name="parameters">Any additional parameters to be included in the message.</param>
		/// <remarks>
		/// <para>
		/// <i>errorNumber</i> is converted to an <see cref="OMIGAERROR"/> constant, and this is 
		/// used to read a message from the database, 
		/// where MESSAGE.MESSAGENUMBER = <i>omigaError</i>. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// <i>parameters</i> is optional. If it is defined it can consist of 1 or more parameters. 
		/// The parameters are used to substitute for any %s placeholders in the message for 
		/// this ErrAssistException object. If any parameters are any left over from this substitution 
		/// then the first one is appended to the end of the message. The message becomes the value of the 
		/// Message property for this ErrAssistException object.
		/// </para>
		/// <para>
		/// If <i>omigaError</i> is a system error (as defined by <see cref="IsSystemError"/>), 
		/// then the exception details are written to the event log.
		/// </para>
		/// <para>
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga system error, then if there is a COM+ transaction it will be set to 
		/// aborted. 
		/// If <paramref name="transactionAction"/> is TransactionAction.SetOnErrorType and the 
		/// exception is an Omiga application error, then if there is a COM+ transaction it will be set 
		/// to completed. 
		/// </para>
		/// </remarks>
		public ErrAssistException(int errorNumber, TransactionAction transactionAction, params string[] parameters)
		{
			_omigaErrorNumber = GetOmigaErrorNumber(errorNumber);
			_id = GuidAssist.CreateGUID();
			_errorSource = MakeStackTrace();
			_exceptionType = GetType().ToString();
			GetMessage(_omigaErrorNumber, out _omigaMessageText, out _omigaMessageType, parameters);

			if (IsSystemError())
			{
				LogError();
			}

			SetTransactionState(transactionAction);
			_transactionAction = transactionAction;
		}

		/// <summary>
		/// The <see cref="OMIGAERROR"/> enum constant for this exception.
		/// </summary>
		/// <remarks>
		/// Defaults to OMIGAERROR.UnspecifiedError.
		/// </remarks>
		public OMIGAERROR OmigaError
		{
			get { return _omigaError; }
		}

		/// <summary>
		/// The Omiga error number for this exception. 
		/// </summary>
		/// <remarks>
		/// Defaults to 0.
		/// </remarks>
		public int OmigaErrorNumber
		{
			get { return _omigaErrorNumber; }
		}

		/// <summary>
		/// The description for this exception.
		/// </summary>
		public override string Message
		{
			get { return _omigaMessageText != null ? _omigaMessageText : base.Message; }
		}

		/// <summary>
		/// The description for this exception. This provides a method of setting the exception message text 
		/// after the exception object has been created, because the Exception.Message property does not support 
		/// a set accessor.
		/// </summary>
		public string OmigaMessageText
		{
			get { return _omigaMessageText != null ? _omigaMessageText : base.Message; }
			set { _omigaMessageText = value; }
		}

		/// <summary>
		/// The Omiga message type for this exception.
		/// </summary>
		/// <remarks>
		/// Equivalent to "select MESSAGETYPE from MESSAGE where MESSAGENUMBER = <i>OmigaError</i>".
		/// </remarks>
		public OmigaMessageType OmigaMessageType
		{
			get { return _omigaMessageType; }
		}

		/// <summary>
		/// The complete description for this exception and recursively any inner exceptions.
		/// </summary>
		public string FullMessage
		{
			get { return GetFullMessage(this); }
		}

		private static string GetFullMessage(Exception exception)
		{
			// For readibility ensure all exception messages end in ".".
			StringBuilder message = new StringBuilder();
			do
			{
				// If the constructor ErrAssistException(Exception exception) is used then the InnerException.Message
				// will be the same as this.Message, so skip this.Message.
				if (exception.InnerException == null || exception.Message != exception.InnerException.Message)
				{
					message.Append(PrettifyMessage(exception.Message));
				}
				exception = exception.InnerException;
			} while (exception != null);

			return message.ToString().Trim();
		}

		private static string PrettifyMessage(string message)
		{
			if (message != null && message.Length > 0)
			{
				message = message.Trim();
				if (message.Length > 0 && message[message.Length - 1] != '.')
				{
					message += '.';
				}
				message += ' ';
			}

			return message;
		}

		/// <summary>
		/// The unique identifier for this exception.
		/// </summary>
		public string Id
		{
			get { return _id; }
		}

		/// <summary>
		/// The source of this exception (normally a stack trace).
		/// </summary>
		public string ErrorSource
		{
			get { return _errorSource; }
			set { _errorSource = value; }
		}

		/// <summary>
		/// The type of this exception.
		/// </summary>
		public string ExceptionType
		{
			get { return _exceptionType; }
			set { _exceptionType = value; }
		}

		/// <summary>
		/// Gets the message text for this exception.
		/// </summary>
		/// <param name="parameters">Any optional parameters for the description.</param>
		/// <returns>
		/// The exception's message text, i.e., 
		/// "select MESSAGETEXT from MESSAGE where MESSAGENUMBER=<i>OmigaErrorNumber</i>", with 
		/// %s place holders replaced by <i>parameters</i>.
		/// </returns>
		public string GetMessageText(params string[] parameters)
		{
			if (_omigaMessageText == null)
			{
				_omigaMessageText = GetMessageText((int)_omigaErrorNumber, parameters);
			}

			return _omigaMessageText;
		}

		private static string GetMessageText(OMIGAERROR omigaErrorNumber, params string[] parameters)
		{
			return GetMessageText((int)omigaErrorNumber, parameters);
		}

		private static string GetMessageText(int errorNumber, params string[] parameters)
		{
			string omigaMessageText = null;
			OmigaMessageType omigaMessageType = OmigaMessageType.None;
			GetMessage(errorNumber, out omigaMessageText, out omigaMessageType, parameters);
			return omigaMessageText;
		}

		/// <summary>
		/// Gets the message type for this exception.
		/// </summary>
		/// <returns>
		/// The exception's message type.
		/// </returns>
		/// <remarks>
		/// The exception's message type, i.e., 
		/// "select MESSAGETYPE from MESSAGE where MESSAGENUMBER=<i>OmigaErrorNumber</i>".
		/// </remarks>
		public OmigaMessageType GetMessageType()
		{
			if (_omigaMessageType == OmigaMessageType.None)
			{
				// Assume message type has not already been read from the database.
				_omigaMessageType = GetMessageType(_omigaErrorNumber);
			}

			return _omigaMessageType;
		}

		private static OmigaMessageType GetMessageType(int errorNumber)
		{
			string omigaMessageText = null;
			OmigaMessageType omigaMessageType = OmigaMessageType.None;
			GetMessage(errorNumber, out omigaMessageText, out omigaMessageType);
			return omigaMessageType;
		}

		private static void GetMessage(int errorNumber, out string omigaMessageText, out OmigaMessageType omigaMessageType, params string[] parameters)
		{
			// First see if the error is not on the database.
			int omigaErrorNumber = GetOmigaErrorNumber(errorNumber);
			omigaMessageText = GetDefaultMessageText(omigaErrorNumber);
			omigaMessageType = OmigaMessageType.Error;
			if (omigaErrorNumber > 0 && omigaMessageText == null)
			{
				try
				{
					using (SqlConnection sqlConnection = new SqlConnection(GlobalProperty.DatabaseConnectionString + "Enlist=false;"))	//CORE264 GHun
					{
						sqlConnection.Open();
						using (SqlCommand sqlCommand = new SqlCommand("select messagetext, messagetype from message where messagenumber = @MessageNumber", sqlConnection))
						{
							sqlCommand.Parameters.Add("@MessageNumber", SqlDbType.Int);
							sqlCommand.Parameters["@MessageNumber"].Value = omigaErrorNumber;
							using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader())
							{
								sqlDataReader.Read();
								omigaMessageText = ExpandMessageText(sqlDataReader.GetString(0), parameters);
								omigaMessageType = ParseOmigaMessageType(sqlDataReader.GetString(1));
							}
						}
					}
				}
				catch (Exception e)
				{
					omigaMessageText = "Unable to get error message " + Convert.ToString(omigaErrorNumber) + " from the database (" + e.Message + ")";
				}
			}
		}

		private static string ExpandMessageText(string omigaMessageText, params string[] parameters)
		{
			if (omigaMessageText != null && parameters != null && parameters.Length > 0)
			{
				// Find %s and substitute it with the substitution parameters
				StringBuilder newMessageText = new StringBuilder();
				int position = omigaMessageText.IndexOf("%s", StringComparison.OrdinalIgnoreCase);
				int parameterIndex = 0;
				while (position != -1 && parameterIndex < parameters.Length)
				{
					newMessageText.Append(omigaMessageText.Substring(0, position));
					// Substitute parameter if present
					if (parameters[parameterIndex] != null)
					{
						newMessageText.Append(parameters[parameterIndex]);
					}
					else
					{
						newMessageText.Append("*** MISSING PARAMETER ***");
					}
					omigaMessageText = omigaMessageText.Substring(position + 2);
					position = omigaMessageText.IndexOf("%s", StringComparison.OrdinalIgnoreCase);
					parameterIndex++;
				}
				newMessageText.Append(omigaMessageText);
				if (parameters.Length > parameterIndex && parameters[parameterIndex] != null)
				{
					// There are additional parameters. 
					// The first additional parameter is the additional error text.
					if (newMessageText.Length > 0)
					{
						newMessageText.Append("; ");
					}
					newMessageText.Append("Details: ");
					newMessageText.Append(parameters[parameterIndex]);
				}

				omigaMessageText = newMessageText.ToString();
			}

			return omigaMessageText;
		}

		private static OmigaMessageType ParseOmigaMessageType(string omigaMessageTypeText)
		{
			OmigaMessageType omigaMessageType = OmigaMessageType.None;

			switch (omigaMessageTypeText.ToLower())
			{
				case "information":
					omigaMessageType = OmigaMessageType.Information;
					break;
				case "warning":
					omigaMessageType = OmigaMessageType.Warning;
					break;
				case "error":
					omigaMessageType = OmigaMessageType.Error;
					break;
			}

			return omigaMessageType;
		}

		private static string GetDefaultMessageText(int omigaErrorNumber)
		{
			string defaultMessageText = null;

			switch (omigaErrorNumber)
			{
				case 556:	//OMIGAERROR.UnableToConnect
					defaultMessageText = "Unable to establish database connection.";
					break;
				case 901:	// OMIGAERROR.InvalidMessageNo
					defaultMessageText = "Error message not found.";
					break;
			}

			return defaultMessageText;
		}

		/// <summary>
		/// Gets an Omiga error number from an Omiga style xml response.
		/// </summary>
		/// <param name="xmlResponseElement">The xml response.</param>
		/// <returns>The Omiga error number or 0 if <i>xmlResponseElement</i> does not contain an error.</returns>
		public static int GetErrorNumberFromResponse(XmlElement xmlResponseElement)
		{
			return GetErrorNumberFromResponse(xmlResponseElement, false);
		}

		/// <summary>
		/// Gets an Omiga error number from an Omiga style xml response.
		/// </summary>
		/// <param name="xmlResponseElement">The xml response.</param>
		/// <param name="isOmigaError">If true then the error number is in native Omiga error number format, 
		/// and not in COM HRESULT format.</param>
		/// <returns>The Omiga error number or 0 if <i>xmlResponseElement</i> does not contain an error.</returns>
		public static int GetErrorNumberFromResponse(XmlElement xmlResponseElement, bool isOmigaError)
		{
			int errorNumber = 0;

			if (xmlResponseElement == null)
			{
				throw new ErrAssistException(OMIGAERROR.InvalidParameter, "Response to HasValue missing");
			}

			if (xmlResponseElement.Name != "RESPONSE")
			{
				throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, "RESPONSE must be top level tag");
			}

			string responseType = xmlResponseElement.GetAttribute("TYPE");
			if (responseType != "SUCCESS" && responseType != "WARNING")
			{
				errorNumber = Convert.ToInt32(xmlResponseElement.GetElementsByTagName("NUMBER").Item(0).InnerText);
				if (isOmigaError)
				{
					errorNumber = GetOmigaErrorNumber(errorNumber);
				}
			}

			return errorNumber;
		}

		/// <summary>
		/// Determines whether this ErrAssistException is for an Omiga application error.
		/// </summary>
		/// <returns>
		/// Returns true if this ErrAssistException object is for an Omiga application error.
		/// </returns>
		public bool IsApplicationError()
		{
			bool isApplicationError = false;

			if (_omigaErrorNumber > 0 && _omigaErrorNumber <= errConstants.clngMAX_ERROR_NO)
			{
				if (_omigaErrorNumber < errConstants.clngADO_START_ERROR_NO || _omigaErrorNumber > errConstants.clngADO_END_ERROR_NO)
				{
					isApplicationError = true;
				}
			}

			return isApplicationError;
		}

		/// <summary>
		/// Determines whether this ErrAssistException object is for an Omiga system error.
		/// </summary>
		/// <returns>
		/// Returns true if this ErrAssistException object is for an Omiga system error.
		/// </returns>
		public bool IsSystemError()
		{
			return !IsApplicationError();
		}

		/// <summary>
		/// Determines whether this ErrAssistException object is for an Omiga warning.
		/// </summary>
		/// <returns>
		/// Returns true if this ErrAssistException object is for an Omiga warnign.
		/// </returns>
		public virtual bool IsWarning()
		{
			bool isWarning = false;

			if (IsApplicationError())
			{
				OmigaMessageType omigaMessageType = GetMessageType(_omigaErrorNumber);
				if (omigaMessageType == OmigaMessageType.Warning ||
					omigaMessageType == OmigaMessageType.Prompt)
				{
					isWarning = true;
				}
			}
 
			return isWarning;
		}

		/// <summary>
		/// Converts an error number to an Omiga number. 
		/// </summary>
		/// <param name="errorNumber">Source error number.</param>
		/// <returns>Omiga error number.</returns>
		/// <remarks>
		/// When errors are raised by VB6
		/// the VB constant 'vbObjectError + 512' is added to them, so to get the
		/// Omiga4 error number orginally raised we need to subtract this. 
		/// Note: system error numbers will end up much larger than Omiga 4 error numbers.
		/// </remarks>
		public static int GetOmigaErrorNumber(int errorNumber)
		{
			if (errorNumber < 0)
			{
				// This is a error in VB6 format.
				errorNumber = errorNumber - errConstants.vbObjectError - 512;
			}

			return errorNumber;
		}

		/// <summary>
		/// Throws this ErrAssistException object if it describes an Omiga error as opposed to an Omiga warning.
		/// </summary>
		/// <returns>The Omiga error type.</returns>
		public OMIGAERRORTYPE CheckError()
		{
			XmlElement xmlResponse = null;
			return CheckError(xmlResponse);
		}

		/// <summary>
		/// Throws this ErrAssistException object if it describes an Omiga error as opposed to an Omiga warning.
		/// </summary>
		/// <param name="xmlResponse">If this ErrAssistException object describes an Omiga warning, then this parameter will be set to an Omiga style warning response.</param>
		/// <returns>The Omiga error type.</returns>
		public OMIGAERRORTYPE CheckError(XmlElement xmlResponse)
		{
			OMIGAERRORTYPE errorType = OMIGAERRORTYPE.omNO_ERR;

			if (_omigaErrorNumber != 0)
			{
				if (IsWarning())
				{
					if (xmlResponse == null)
					{
						throw new ErrAssistException(OMIGAERROR.XMLMissingElement, "Missing RESPONSE element");
					}
					// Add warning to the response node.
					AddWarning(xmlResponse);
					errorType = OMIGAERRORTYPE.omWARNING;
				}
				else
				{
					throw this;
				}
			}

			return errorType;
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="xmlResponseNodeToCheck">The xml response to check.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponseNode(XmlElement xmlResponseNodeToCheck)
		{
			return CheckXmlResponseNode((XmlNode)xmlResponseNodeToCheck, null, false);
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="xmlResponseNodeToCheck">The xml response to check.</param>
		/// <param name="xmlResponseNodeToAdd">The xml response to add any Omiga warning to.</param>
		/// <param name="throwError">If true then an exception will be thrown if the response contains an error.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponseNode(XmlElement xmlResponseNodeToCheck, XmlElement xmlResponseNodeToAdd, bool throwError)
		{
			return CheckXmlResponseNode((XmlNode)xmlResponseNodeToCheck, xmlResponseNodeToAdd, throwError);
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="xmlResponseNodeToCheck">The xml response to check.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponseNode(XmlNode xmlResponseNodeToCheck)
		{
			return CheckXmlResponseNode(xmlResponseNodeToCheck, null, false);
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="xmlResponseNodeToCheck">The xml response to check.</param>
		/// <param name="xmlResponseNodeToAdd">The xml response to add any Omiga warning to.</param>
		/// <param name="throwError">If true then an exception will be thrown if the response contains an error.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponseNode(XmlNode xmlResponseNodeToCheck, XmlElement xmlResponseNodeToAdd, bool throwError)
		{
			int errorNumber = (int)OMIGAERROR.UnspecifiedError;

			if (xmlResponseNodeToCheck == null)
			{
				throw new ErrAssistException(OMIGAERROR.XMLMissingElement, "Missing RESPONSE element");
			}
			if (xmlResponseNodeToCheck.Name != "RESPONSE")
			{
				throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, "RESPONSE must be the top level tag");
			}
			if (xmlResponseNodeToAdd != null && xmlResponseNodeToAdd.Name != "RESPONSE")
			{
				throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, "RESPONSE must be the top level tag");
			}
			XmlNode xmlResponseTypeNode = xmlResponseNodeToCheck.Attributes.GetNamedItem("TYPE");
			if (xmlResponseTypeNode != null)
			{
				if (xmlResponseTypeNode.InnerText == "WARNING")
				{
					if (xmlResponseNodeToAdd != null)
					{
						xmlResponseNodeToAdd.SetAttribute("TYPE", "WARNING");
						XmlElement xmlFirstChild = (XmlElement)xmlResponseNodeToAdd.FirstChild;
						XmlNodeList xmlMessageList = xmlResponseNodeToCheck.SelectNodes("MESSAGE");
						// Insert messages at the top of the response to add element.
						foreach (XmlElement xmlMessageElem in xmlMessageList)
						{
							xmlResponseNodeToAdd.InsertBefore(xmlMessageElem.CloneNode(true), xmlFirstChild);
						}
					}
				}
				else if (xmlResponseTypeNode.InnerText != "SUCCESS")
				{
					XmlNode xmlResponseErrorNumber = xmlResponseNodeToCheck.SelectSingleNode("ERROR/NUMBER");
					if (xmlResponseErrorNumber != null)
					{
						Int32.TryParse(xmlResponseErrorNumber.InnerText, out errorNumber);
					}
					xmlResponseErrorNumber = null;
					if (throwError)
					{
						string errorSource = MakeStackTrace();
						string description = String.Empty;
						XmlNode xmlResponseErrorSource = xmlResponseNodeToCheck.SelectSingleNode("ERROR/SOURCE");
						if (xmlResponseErrorSource != null)
						{
							if (xmlResponseErrorSource.InnerText.Length > 0)
							{
								errorSource = xmlResponseErrorSource.InnerText;
							}
						}
						xmlResponseErrorSource = null;
						XmlNode xmlResponseErrorDescription = xmlResponseNodeToCheck.SelectSingleNode("ERROR/DESCRIPTION");
						if (xmlResponseErrorDescription != null)
						{
							if (xmlResponseErrorDescription.InnerText.Length > 0)
							{
								description = xmlResponseErrorDescription.InnerText;
							}
						}
						xmlResponseErrorDescription = null;
						if (description.Length == 0)
						{
							description = GetMessageText(OMIGAERROR.UnspecifiedError);
						}

						Exception exception = new ErrAssistException(errorNumber, description);
						exception.Source = errorSource;
						throw exception;
					}
				}
				else
				{
					errorNumber = 0;
				}
			}

			return errorNumber;
		}

		/// <summary>
		/// Checks an Omiga xml response; does not throw an exception if it contains an error.
		/// </summary>
		/// <param name="response">The xml response to check, as a string.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponse(string response)
		{
			return CheckXmlResponse(response, false, null);
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="response">The xml response to check, as a string.</param>
		/// <param name="throwError">If true then an exception will be thrown if the response contains an error.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponse(string response, bool throwError)
		{
			return CheckXmlResponse(response, throwError, null);
		}

		/// <summary>
		/// Checks an Omiga xml response and optionally throws an exception if it contains an error.
		/// </summary>
		/// <param name="response">The xml response to check, as a string.</param>
		/// <param name="throwError">If true then an exception will be thrown if the response contains an error.</param>
		/// <param name="xmlInResponseElement">The xml response to add any Omiga warning to.</param>
		/// <returns>
		/// The error number in the xml response, or OMIGAERROR.UnspecifiedError if the response does not contain an error.
		/// </returns>
		public static int CheckXmlResponse(string response, bool throwError, XmlElement xmlInResponseElement)
		{
			XmlDocument xmlXmlIn = XmlAssist.Load(response);
			XmlElement xmlResponseElement = (XmlElement)xmlXmlIn.SelectSingleNode("RESPONSE");
			if (xmlResponseElement == null)
			{
				throw new ErrAssistException(OMIGAERROR.XMLMissingElement, "Missing RESPONSE element");
			}

			return CheckXmlResponseNode(xmlResponseElement, xmlInResponseElement, throwError);
		}

		/// <summary>
		/// Throws an exception if the specified Omiga xml response contains an error.
		/// </summary>
		/// <param name="response">The Omiga xml response to check, as a string.</param>
		public static void RaiseXmlResponse(string response)
		{
			XmlDocument xmlXmlIn = XmlAssist.Load(response);
			RaiseXmlResponseNode(xmlXmlIn.DocumentElement);
		}

		/// <summary>
		/// Throws an exception if the specified Omiga xml response contains an error.
		/// </summary>
		/// <param name="xmlResponseElement">The Omiga xml response to check.</param>
		public static void RaiseXmlResponseNode(XmlElement xmlResponseElement)
		{
			if (xmlResponseElement == null)
			{
				throw new ErrAssistException(OMIGAERROR.InvalidParameter, "Response to check is missing");
			}

			if (xmlResponseElement.Name != "RESPONSE")
			{
				throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, "RESPONSE must be top level tag");
			}

			if (xmlResponseElement.GetAttributeNode("TYPE") != null)
			{
				string responseType = xmlResponseElement.GetAttribute("TYPE");
				if (responseType != "SUCCESS" && responseType != "WARNING")
				{
					XmlNode xmlResponseErrorNumber = xmlResponseElement.SelectSingleNode("ERROR/NUMBER");
					int errorNumber = 0;
					if (xmlResponseErrorNumber != null)
					{
						Int32.TryParse(xmlResponseErrorNumber.InnerText, out errorNumber);
					}
					xmlResponseErrorNumber = null;

					string exceptionType = GetExceptionTypeFromResponse(xmlResponseElement);

					if (errorNumber != 0 || exceptionType.Length > 0)
					{
						string errorSource = String.Empty;
						string errorDescription = String.Empty;

						XmlNode xmlResponseErrorSource = xmlResponseElement.SelectSingleNode("ERROR/SOURCE");
						if (xmlResponseErrorSource != null)
						{
							if (xmlResponseErrorSource.InnerText.Length > 0)
							{
								errorSource = xmlResponseErrorSource.InnerText;
							}
						}
						XmlNode xmlResponseErrorDesc = xmlResponseElement.SelectSingleNode("ERROR/DESCRIPTION");
						if (xmlResponseErrorDesc != null)
						{
							if (xmlResponseErrorDesc.InnerText.Length > 0)
							{
								errorDescription = xmlResponseErrorDesc.InnerText;
							}
						}
						if (errorDescription.Length == 0)
						{
							errorDescription = GetMessageText(OMIGAERROR.UnspecifiedError);
						}

						ErrAssistException errAssistException = new ErrAssistException(errorNumber, errorDescription);
						errAssistException.ErrorSource = errorSource;
						errAssistException.ExceptionType = exceptionType;
						throw errAssistException;
					}
				}
			}
		}

		/// <summary>
		/// Gets the exception type from an Omiga xml response.
		/// </summary>
		/// <param name="xmlResponseElement">The Omiga xml response.</param>
		/// <returns>The exception type.</returns>
		public static string GetExceptionTypeFromResponse(XmlElement xmlResponseElement)
		{
			if (xmlResponseElement == null)
			{
				throw new ArgumentNullException("xmlResponseElement");
			}

			string exceptionType = null;
			XmlNode xmlResponseErrorExceptionType = xmlResponseElement.SelectSingleNode("ERROR/EXCEPTIONTYPE");
			if (xmlResponseErrorExceptionType != null)
			{
				if (xmlResponseErrorExceptionType.InnerText.Length > 0)
				{
					exceptionType = xmlResponseErrorExceptionType.InnerText;
				}
			}
			return exceptionType;
		}

		/// <summary>
		/// Creates an Omiga xml error response from this ErrAssistException object.
		/// </summary>
		/// <returns>
		/// The Omiga xml error response as a string, in the format
		/// <code>
		/// &lt;RESPONSE TYPE="APPERR|SYSERR"&gt;
		///		&lt;ERROR&gt;
		///			&lt;NUMBER&gt;The error number in HRESULT format.&lt;/NUMBER&gt;
		///			&lt;EXCEPTIONTYPE&gt;The type of the .Net exception.&lt;/NUMBER&gt;
		///			&lt;ID&gt;A unique identifier for this exception.&lt;/ID&gt;
		///			&lt;DESCRIPTION&gt;The full error message.&lt;/DESCRIPTION&gt;
		///			&lt;SOURCE&gt;The source of the exception.&lt;/SOURCE&gt;
		///			&lt;VERSION&gt;The assembly version number.&lt;/VERSION&gt;
		///		&lt;/ERROR&gt;
		/// &lt;/RESPONSE&gt;
		/// </code>
		/// </returns>
		public string CreateErrorResponse()
		{
			XmlElement xmlErrorElement = (XmlElement)CreateErrorResponseNode(false);
			if (IsSystemError())
			{
				LogError(null, EventLogEntryType.Error);
			}

			return xmlErrorElement.OuterXml;
		}

		/// <summary>
		/// Creates an Omiga xml error response from this ErrAssistException object.
		/// </summary>
		/// <returns>
		/// The Omiga xml error response, in the format
		/// <code>
		/// &lt;RESPONSE TYPE="APPERR|SYSERR"&gt;
		///		&lt;ERROR&gt;
		///			&lt;NUMBER&gt;The error number in HRESULT format.&lt;/NUMBER&gt;
		///			&lt;EXCEPTIONTYPE&gt;The type of the .Net exception.&lt;/NUMBER&gt;
		///			&lt;ID&gt;A unique identifier for this exception.&lt;/ID&gt;
		///			&lt;DESCRIPTION&gt;The full error message.&lt;/DESCRIPTION&gt;
		///			&lt;SOURCE&gt;The source of the exception.&lt;/SOURCE&gt;
		///			&lt;VERSION&gt;The assembly version number.&lt;/VERSION&gt;
		///		&lt;/ERROR&gt;
		/// &lt;/RESPONSE&gt;
		/// </code>
		/// </returns>
		public XmlNode CreateErrorResponseNode()
		{
			return CreateErrorResponseNode(true);
		}

		/// <summary>
		/// Creates an Omiga xml error response from this ErrAssistException object.
		/// </summary>
		/// <param name="logError">If true and the error is a system error, then it will be written to the Windows event log.</param>
		/// <returns>
		/// The Omiga xml error response, in the format
		/// <code>
		/// &lt;RESPONSE TYPE="APPERR|SYSERR"&gt;
		///		&lt;ERROR&gt;
		///			&lt;NUMBER&gt;The error number in HRESULT format.&lt;/NUMBER&gt;
		///			&lt;EXCEPTIONTYPE&gt;The type of the .Net exception.&lt;/NUMBER&gt;
		///			&lt;ID&gt;A unique identifier for this exception.&lt;/ID&gt;
		///			&lt;DESCRIPTION&gt;The full error message.&lt;/DESCRIPTION&gt;
		///			&lt;SOURCE&gt;The source of the exception.&lt;/SOURCE&gt;
		///			&lt;VERSION&gt;The assembly version number.&lt;/VERSION&gt;
		///		&lt;/ERROR&gt;
		/// &lt;/RESPONSE&gt;
		/// </code>
		/// </returns>
		public XmlNode CreateErrorResponseNode(bool logError)
		{
			XmlDocument xmlDocument = new XmlDocument();
			XmlElement xmlReponseElement = xmlDocument.CreateElement("RESPONSE");
			xmlDocument.AppendChild(xmlReponseElement);
			XmlElement xmlErrorElement = xmlDocument.CreateElement("ERROR");
			xmlReponseElement.AppendChild(xmlErrorElement);
			XmlElement xmlElement = xmlDocument.CreateElement("NUMBER");
			xmlElement.InnerText = _omigaErrorNumber.ToString();
			xmlErrorElement.AppendChild(xmlElement);
			xmlElement = xmlDocument.CreateElement("EXCEPTIONTYPE");
			xmlElement.InnerText = _exceptionType;
			xmlErrorElement.AppendChild(xmlElement);
			xmlElement = xmlDocument.CreateElement("ID");
			xmlElement.InnerText = _id;
			xmlErrorElement.AppendChild(xmlElement);
			XmlElement xmlDescriptionElement = xmlDocument.CreateElement("DESCRIPTION");
			xmlDescriptionElement.InnerText = FullMessage;
			xmlErrorElement.AppendChild(xmlDescriptionElement);
			xmlElement = xmlDocument.CreateElement("VERSION");
			xmlElement.InnerText = App.Comments;
			xmlErrorElement.AppendChild(xmlElement);
			xmlElement = xmlDocument.CreateElement("SOURCE");
			xmlElement.InnerText = _errorSource;
			xmlErrorElement.AppendChild(xmlElement);
			if (IsApplicationError())
			{
				xmlReponseElement.SetAttribute("TYPE", "APPERR");
				if (xmlDescriptionElement.InnerText.Length == 0)
				{
					xmlDescriptionElement.InnerText = GetMessageText();
				}
			}
			else
			{
				xmlReponseElement.SetAttribute("TYPE", "SYSERR");
			}

			if (logError && IsSystemError())
			{
				LogError(null, EventLogEntryType.Error);
			}

			return xmlReponseElement.CloneNode(true);
		}

		/// <summary>
		/// Adds a warning for this ErrAssistException object to an Omiga xml response.
		/// </summary>
		/// <param name="xmlResponse">The Omiga xml response to which to add the warning.</param>
		public void AddWarning(XmlElement xmlResponse)
		{
			AddWarning(xmlResponse, _omigaErrorNumber, this);
		}

		/// <summary>
		/// Adds a warning for a specified Omiga error number to an Omiga xml response.
		/// </summary>
		/// <param name="xmlResponse">The Omiga xml response to which to add the warning.</param>
		/// <param name="omigaErrorNumber">The Omiga error number for the warning.</param>
		public static void AddWarning(XmlElement xmlResponse, int omigaErrorNumber)
		{
			AddWarning(xmlResponse, omigaErrorNumber, null);
		}

		/// <summary>
		/// Adds a warning for a specified exception to an Omiga xml response.
		/// </summary>
		/// <param name="xmlResponse">The Omiga xml response to which to add the warning.</param>
		/// <param name="exception">The exception that contains the warning.</param>
		public static void AddWarning(XmlElement xmlResponse, Exception exception)
		{
			AddWarning(xmlResponse, 0, exception);
		}

		private static void AddWarning(XmlElement xmlResponse, int omigaErrorNumber, Exception exception)
		{
			if (xmlResponse.Name != "RESPONSE")
			{
				throw new ErrAssistException(OMIGAERROR.MissingPrimaryTag, "RESPONSE must be the top level tag");
			}

			string messageText = String.Empty;
			string messageType = String.Empty;
			if (exception != null)
			{
				ErrAssistException errAssistException = exception as ErrAssistException;
				if (errAssistException != null)
				{
					messageText = errAssistException.FullMessage;
					messageType = errAssistException.OmigaMessageType.ToString();
				}
				else
				{
					messageText = exception.Message;
					messageType = OmigaMessageType.Error.ToString();
				}
			}
			else if (omigaErrorNumber != 0)
			{
				OmigaMessageType omigaMessageType = OmigaMessageType.Error;
				GetMessage(omigaErrorNumber, out messageText, out omigaMessageType);
				messageType = omigaMessageType.ToString();
			}

			xmlResponse.SetAttribute("TYPE", "WARNING");
			XmlNode xmlFirstChild = xmlResponse.FirstChild;
			XmlElement xmlMessageElement = xmlResponse.OwnerDocument.CreateElement("MESSAGE");
			xmlResponse.InsertBefore(xmlMessageElement, xmlFirstChild);
			XmlElement xmlElement = xmlResponse.OwnerDocument.CreateElement("MESSAGETEXT");
			xmlElement.InnerText = messageText;
			xmlMessageElement.AppendChild(xmlElement);
			xmlElement = xmlResponse.OwnerDocument.CreateElement("MESSAGETYPE");
			xmlElement.InnerText = messageType;
			xmlMessageElement.AppendChild(xmlElement);
		}

		/// <summary>
		/// Gets an Omiga message response for this ErrAssistException object.
		/// </summary>
		/// <returns>The Omiga message response.</returns>
		public string FormatMessageNode()
		{
			XmlDocument xmlMessageDoc = new XmlDocument();
			XmlElement xmlMessageElement = xmlMessageDoc.CreateElement("MESSAGE");
			xmlMessageDoc.AppendChild(xmlMessageElement);
			XmlElement xmlElement = xmlMessageDoc.CreateElement("TEXT");
			xmlElement.InnerText = FullMessage;
			xmlMessageElement.AppendChild(xmlElement);
			xmlElement = xmlMessageDoc.CreateElement("TYPE");
			xmlElement.InnerText = "Warning";
			xmlMessageElement.AppendChild(xmlElement);
			return xmlMessageDoc.OuterXml;
		}

		/// <summary>
		/// Logs this ErrAssistException object to the Windows event log as an error.
		/// </summary>
		public void LogError()
		{
			LogError(null, EventLogEntryType.Error);
		}

		/// <summary>
		/// Logs this ErrAssistException object and the specific text to the Windows event log as an error.
		/// </summary>
		/// <param name="errorText">The text to include in the event log entry.</param>
		public void LogError(string errorText)
		{
			LogError(errorText, EventLogEntryType.Error);
		}

		/// <summary>
		/// Logs this ErrAssistException object and the specific text to the Windows event log as the specified type of event.
		/// </summary>
		/// <param name="errorText">The text to include in the event log entry.</param>
		/// <param name="eventLogEntryType">The type of event log entry to log.</param>
		/// <remarks>
		/// If this ErrAssistException object has already been logged to the event log it will not be 
		/// logged again.
		/// </remarks>
		public void LogError(string errorText, EventLogEntryType eventLogEntryType)
		{
			if (!_logged)
			{
				try
				{
					string description = FullMessage;
					if (errorText != null)
					{
						description += " Details: " + errorText;
					}
					string message =
						(IsApplicationError() ? "APPERR" : "SYSERR") + "\r" +
						"Error number: " + Convert.ToString(_omigaErrorNumber) + "\r" +
						"Exception type: " + _exceptionType + "\r" +
						"Id: " + _id + "\r" +
						"Description: " + description + "\r" +
						"Source:\r" + _errorSource;

					App.LogEvent(message, eventLogEntryType);
					Debug.WriteLine(message);
				}
				finally
				{
					_logged = true;
				}
			}
		}

		#region COM+ Transactions

		/// <summary>
		/// Sets the state of any current COM+ transaction.
		/// </summary>
		/// <param name="transactionAction">Determines how to set the transaction.</param>
		public void SetTransactionState(TransactionAction transactionAction)
		{
			switch (transactionAction)
			{
				case TransactionAction.SetComplete:
					SetComplete();
					break;
				case TransactionAction.SetAbort:
					SetAbort();
					break;
				case TransactionAction.SetOnErrorType:
					SetOnErrorType();
					break;
				case TransactionAction.SetCompleteOnSystemError:
					SetCompleteOnSystemError();
					break;
				case TransactionAction.SetAbortOnSystemError:
					SetAbortOnSystemError();
					break;
			}
		}

		/// <summary>
		/// Sets the state of any current COM+ transaction to complete.
		/// </summary>
		public static void SetComplete()
		{
			ContextUtility.SetComplete();
		}

		/// <summary>
		/// Sets the state of any current COM+ transaction to aborted.
		/// </summary>
		public static void SetAbort()
		{
			ContextUtility.SetAbort();
		}

		/// <summary>
		/// If the current error is a system error, then sets the state of 
		/// any COM+ transaction, to aborted, otherwise sets the state to completed.
		/// </summary>
		public void SetOnErrorType()
		{
			if (IsSystemError()) ContextUtility.SetAbort(); else ContextUtility.SetComplete();
		}

		/// <summary>
		/// If the current error is a system error, then sets the state of 
		/// any COM+ transaction to completed.
		/// </summary>
		public void SetCompleteOnSystemError()
		{
			if (IsSystemError()) ContextUtility.SetComplete();
		}

		/// <summary>
		/// If the current error is a system error, then sets the state of 
		/// any COM+ transaction to aborted.
		/// </summary>
		public void SetAbortOnSystemError()
		{
			if (IsSystemError()) ContextUtility.SetAbort(); 
		}
		#endregion

		/*
		#region Resume next
		public void Throw()
		{
			if (!ResumeNext.IsResumeNext(this))
			{
				throw this;
			}
		}
		#endregion
		*/

		private static string MakeStackTrace()
		{
			StackTrace stackTrace = new StackTrace(true);
			return stackTrace.ToString();
		}

		#region Obsolete

		// Missing XML comment for publicly visible type or member 
		#pragma warning disable 1591

		[Obsolete("Use CreateErrorResponse instead")]
		public XmlNode CreateErrorResponseEx()
		{
			return CreateErrorResponseNode();
		}

		[Obsolete("Use CheckXmlResponseNode instead")]
		public static int CheckResponse(XmlElement xmlResponseNodeToCheck, XmlElement xmlResponseNodeToAdd, bool throwError)
		{
			return CheckXmlResponseNode(xmlResponseNodeToCheck, xmlResponseNodeToAdd, throwError);
		}

		[Obsolete("Use CheckXmlResponse instead")]
		public static int CheckXMLResponse(string response, bool throwError, XmlElement xmlInResponseElement)
		{
			return CheckXmlResponse(response, throwError, xmlInResponseElement);
		}

		[Obsolete("Use System.Exception.ToString instead")]
		public static string FormatParserError(Exception exception)
		{
			return exception.ToString();
		}

		[Obsolete("Throw a new ErrAssistException instead")]
		public static void RaiseError(string objectName, string functionName, int errorNumber, params string[] options)
		{
			throw new System.NotImplementedException("RaiseError is not implemented - throw a new ErrAssistException instead");
		}

		[Obsolete("Throw an ErrAssistException instead")]
		public static void ReRaise()
		{
			throw new System.NotImplementedException("ReRaise is not implemented - throw an ErrAssistException instead");
		}

		[Obsolete("Use RaiseXmlResponseNode instead")]
		public static int ReRaiseResponseError(XmlElement xmlResponseElement)
		{
			RaiseXmlResponseNode(xmlResponseElement);
			return 0;
		}

		[Obsolete("Throw a new ErrAssistException instead")]
		public static void ThrowError(string objectName, string functionName, int omigaErrorNumber, params string[] options)
		{
			throw new System.NotImplementedException("ThrowError is not implemented - throw a new ErrAssistException instead");
		}

		[Obsolete("SaveErr is not implemented and should no longer be necessary")]
		public static void SaveErr()
		{
			throw new System.NotImplementedException("SaveErr is not implemented and should no longer be necessary");
		}

		[Obsolete("Use AddWarning instead")]
		public static void errAddOmigaWarning(XmlElement xmlResponse, int omigaErrorNumber)
		{
			AddWarning(xmlResponse, omigaErrorNumber, null);
		}

		[Obsolete("Use AddWarning instead")]
		public void errAddWarning(XmlElement xmlResponse)
		{
			AddWarning(xmlResponse);
		}

		[Obsolete("Use CheckError instead")]
		public int errCheckError(string functionName, string objectName, XmlElement xmlResponse)
		{
			return (int)CheckError(xmlResponse);
		}

		[Obsolete("Use CheckXmlResponse instead")]
		public static int errCheckXMLResponse(string response, bool throwError, XmlElement xmlResponse)
		{
			return CheckXmlResponse(response, throwError, xmlResponse);
		}

		[Obsolete("Use CheckXmlResponseNode instead")]
		public static int errCheckXMLResponseNode(XmlElement xmlResponseNodeToCheck, XmlElement xmlResponseNodeToAdd, bool throwError)
		{
			return CheckXmlResponseNode(xmlResponseNodeToCheck, xmlResponseNodeToAdd, throwError);
		}

		[Obsolete("Use CreateErrorResponse instead")]
		public string errCreateErrorResponse()
		{
			return CreateErrorResponse();
		}

		[Obsolete("Use GetMessageText instead")]
		public static string errGetMessageText(int errorNumber, int messageField)
		{
			return GetMessageText(errorNumber);
		}

		[Obsolete("Use GetOmigaErrorNumber instead")]
		public static int errGetOmigaErrorNumber(int errorNumber)
		{
			return GetOmigaErrorNumber(errorNumber);
		}

		[Obsolete("Use FormatMessageNode instead")]
		public string errFormatMessageNode()
		{
			return FormatMessageNode();
		}

		[Obsolete("Use IsApplicationError instead")]
		public bool errIsApplicationError()
		{
			return IsApplicationError();
		}

		[Obsolete("Use IsWarning instead")]
		public bool errIsWarning()
		{
			return IsWarning();
		}

		[Obsolete("Use IsSystemError instead")]
		public bool errIsSystemError()
		{
			return IsSystemError();
		}

		[Obsolete("Use RaiseXmlResponse instead")]
		public static void errRaiseXMLResponse(string response)
		{
			RaiseXmlResponse(response);
		}

		[Obsolete("Use RaiseXmlResponseNode instead")]
		public static void errRaiseXMLResponseNode(XmlElement xmlResponseElement)
		{
			RaiseXmlResponseNode(xmlResponseElement);
		}

		[Obsolete("Throw a new ErrAssistException instead")]
		public static void errThrowError(string functionName, int omigaErrorNumber, params string[] options)
		{
			throw new System.NotImplementedException("errThrowError is not implemented - throw a new ErrAssistException instead");
		}

		#pragma warning restore 1591

		#endregion

	}
}
