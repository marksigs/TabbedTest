/*
--------------------------------------------------------------------------------------------
Workfile:			TraceAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Tracing.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		21/06/2007	First .Net version. Ported from TraceAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.IO;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using Vertex.Fsd.Omiga.omLogging;

// Ambiguous reference in cref attribute.
#pragma warning disable 419

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Defines a method for initialising a trace object. 
	/// </summary>
	/// <remarks>
	/// This interface should be implemented by 
	/// any class that has a TraceAssist object which is called by another class with a TraceAssist object, 
	/// where the TraceAssist object in the called class should assume the properties of the TraceAssist 
	/// object in the calling class. The calling class calls <b>TraceInitialiseOffspring</b>, 
	/// passing the called class as the ITraceAssist parameter; this results in 
	/// <see cref="InitialiseTraceInterface"/> being called for the called class via the ITraceAssist 
	/// interface.
	/// </remarks>
	public interface ITraceAssist
	{
		/// <summary>
		/// Initialises a trace object.
		/// </summary>
		/// <param name="isTracingEnabled">Indicates whether tracing is enabled.</param>
		/// <param name="traceId">The unqiue identifier for the trace.</param>
		/// <param name="startTime">The date and time when the trace started.</param>
		/// <remarks>
		/// This method is generally implemented as follows:
		/// <code>
		/// public void InitialiseTraceInterface(bool isTracingEnabled, string traceId, DateTime startTime)
		/// {
		///	 if (isTracingEnabled)
		///	 {
		///		 _traceAssist.TraceInitialiseFromParent(isTracingEnabled, traceId, startTime);
		///	 }
		/// }
		/// </code>
		/// </remarks>
		void InitialiseTraceInterface(bool isTracingEnabled, string traceId, DateTime startTime);
	}

	/// <summary>
	/// Provides logging support. Ported from TraceAssist.cls.
	/// </summary>
	public sealed class TraceAssist
	{
		private Logger _logger;
		private string _loggerName;
		private string _traceFolder;
		private string _traceId;
		private DateTime _startTime;
		private bool _isTracingEnabled;
		private bool _isXmlTracingEnabled;

		/// <summary>
		/// Creates a new logger with a specified name.
		/// </summary>
		/// <param name="loggerName">The name of the logger.</param>
		/// <remarks>
		/// <para>
		/// A log4net logger is created with the name <paramref name="loggerName"/>. 
		/// The logger is configured either from:
		/// <list type="bullet">
		///		<item>
		///			<description>
		///			A log4net xml configuration file for the current assembly.
		///			<para>
		///			The location of the configration file is defined by one of:
		///			</para>
		///			<list type="bullet">
		///				<item>
		///					<description>
		///					A <b>LoggingConfigurationPath</b> entry in the assembly .config file 
		///					<b>appSettings</b> section.
		///					</description>
		///				</item>
		///				<item>
		///					<description>
		///					The registry entry 
		///					<b>HKEY_LOCAL_MACHINE\SOFTWARE\Vertex\Fsd\Logging\LoggingConfigurationPath</b>.
		///					</description>
		///				</item>
		///				<item>
		///					<description>
		///					The registry entry 
		///					<b>HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\SYSTEM CONFIGURATION\Logging\LoggingConfigurationPath</b>.
		///					</description>
		///				</item>
		///			</list>
		///			<para>
		///			The name of the configuration file is defined by one of:
		///			</para>
		///			<list type="bullet">
		///				<item>
		///					<description>
		///					A <b>LoggingConfigurationFile</b> entry in the assembly .config file 
		///					<b>appSettings</b> section.
		///					</description>
		///				</item>
		///				<item>
		///					<description>
		///					The registry entry 
		///					<b>HKEY_LOCAL_MACHINE\SOFTWARE\Vertex\Fsd\Logging\LoggingConfigurationFile</b>.
		///					</description>
		///				</item>
		///				<item>
		///					<description>
		///					The registry entry 
		///					<b>HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\SYSTEM CONFIGURATION\Logging\LoggingConfigurationFile</b>.
		///					</description>
		///				</item>
		///			</list>
		///			<para>
		///			Tracing can be switched off by setting the log4net level to "OFF".
		///			</para>
		///			<para>
		///			General tracing (e.g., <see cref="TraceMethodEntry"/>, <see cref="TraceMethodExit"/> and <see cref="TraceMessage"/>) 
		///			is enabled at log4net level "INFO". Lines are logged to the file named in the log4net 
		///			xml configuration file. Each line contains a unique trace id for this TraceAssist object.
		///			</para>
		///			<para>
		///			Xml tracing (<see cref="TraceRequest"/>, <see cref="TraceResponse"/> and <see cref="TraceXML"/>) 
		///			is enabled at log4net level "DEBUG" (this also enables general tracing). Xml documents are saved to 
		///			separate files in the same directory as the general trace file. The file name includes the unique trace id for this 
		///			TraceAssist object. Note: xml documents are written using the 
		///			<see cref="System.Xml.XmlDocument.Save"/> method and not log4net (although the location 
		///			of the xml file is defined in the log4net configuration file).
		///			</para>
		///			</description>
		///		</item>
		///		<item>
		///		From the registry; this is only used if configuration via log4net has failed. 
		///		The location of the log file is given by the registry entry 
		///		<b>HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\TRACE\Folder</b>. The value of the default registry 
		///		entry at <b>HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\TRACE</b> determines the amount of tracing: 
		///		0 is no tracing, 1 is general tracing, and 2 is general tracing plus xml tracing.
		///		<para>
		///		Note if configuration via log4net has failed, then it likely that only xml tracing will 
		///		be enabled, and general tracing will not work - the former does not use log4net whilst 
		///		the latter does, so general tracing will not function correctly if the log4net configuration 
		///		is not correct.
		///		</para>
		///		</item>
		/// </list>
		/// </para>
		/// </remarks>
		public TraceAssist(string loggerName)
		{
			StartTrace(loggerName);
		}

		/// <summary>
		/// Creates a new logger with a specified name.
		/// </summary>
		/// <param name="loggerName">The name of the logger.</param>
		/// <remarks>
		/// See TraceAssist constructor.
		/// </remarks>
		public void StartTrace(string loggerName) 
		{
			if (_loggerName == null)
			{
				_loggerName = loggerName;
				_logger = LogManager.GetLogger(_loggerName);

				if (!LoadConfigurationFromLogger())
				{
					LoadConfigurationFromRegistry();
				}

				if (_isTracingEnabled && _traceFolder != null && _traceFolder.Length > 0)
				{
					if (!Directory.Exists(_traceFolder))
					{
						lock (_traceFolder)
						{
							if (!Directory.Exists(_traceFolder))
							{
								Directory.CreateDirectory(_traceFolder);
							}
						}
					}

					_traceId = DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Thread.CurrentThread.ManagedThreadId.ToString() + "_" + Path.GetFileNameWithoutExtension(Path.GetTempFileName());
					_startTime = DateTime.Now;
				}
			}
		}

		private bool LoadConfigurationFromLogger()
		{
			bool loaded = false;

			log4net.Repository.ILoggerRepository loggerRepository = log4net.LogManager.GetRepository();
			log4net.Core.ILogger[] loggers = loggerRepository.GetCurrentLoggers();
			foreach (log4net.Core.ILogger logger in loggers)
			{
				if (logger.Name == _loggerName)
				{
					log4net.Appender.IAppender[] appenders = logger.Repository.GetAppenders();

					foreach (log4net.Appender.IAppender appender in appenders)
					{
						if (appender is log4net.Appender.FileAppender)
						{
							log4net.Appender.FileAppender fileAppender = (log4net.Appender.FileAppender)appender;
							if (fileAppender.Name.Contains(_loggerName))
							{
								_isTracingEnabled = _logger.IsInfoEnabled || _logger.IsDebugEnabled;
								_isXmlTracingEnabled = _logger.IsDebugEnabled;
								_traceFolder = Path.GetDirectoryName(fileAppender.File);
								loaded = true;
								break;
							}
						}
					}
					break;
				}
			}

			return loaded;
		}

		private bool LoadConfigurationFromRegistry()
		{
			bool loaded = false;

			using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("Software\\Omiga4\\TRACE"))
			{
				if (registryKey != null)
				{
					string traceLevel = Convert.ToString(registryKey.GetValue(null));

					if (traceLevel == "1" || traceLevel == "2")
					{
						_traceFolder = Convert.ToString(registryKey.GetValue("FOLDER"));

						if (_traceFolder != null && _traceFolder.Length > 0)
						{
							_isTracingEnabled = true;
							if (traceLevel == "2")
							{
								_isXmlTracingEnabled = true;
							}
							loaded = true;
						}
					}
				}
			}

			return loaded;
		}

		/// <summary>
		/// Initialises the trace object from the properties of an existing trace.
		/// </summary>
		/// <param name="isTracingEnabled">Indicates whether tracing is enabled.</param>
		/// <param name="traceId">The unique identifier for this trace.</param>
		/// <param name="startTime">The date and time the trace started.</param>
		/// <remarks>
		/// See <see cref="ITraceAssist"/>.
		/// </remarks>
		public void TraceInitialiseFromParent(bool isTracingEnabled, string traceId, DateTime startTime) 
		{
			_isTracingEnabled = isTracingEnabled;
			_traceId = traceId;
			_startTime = startTime;
		}

		/// <summary>
		/// Initialises the trace object in a called component.
		/// </summary>
		/// <param name="iTraceAssist">The called component (which must implement the ITraceAssist interface).</param>
		/// <remarks>
		/// See <see cref="ITraceAssist"/>.
		/// </remarks>
		public void TraceInitialiseOffspring(ITraceAssist iTraceAssist) 
		{
			if (_isTracingEnabled)
			{
				iTraceAssist.InitialiseTraceInterface(_isTracingEnabled, _traceId, _startTime);
			}
		}

		/// <summary>
		/// Traces the entry into a method.
		/// </summary>
		/// <param name="moduleName">The name of the module to which the method belongs.</param>
		/// <param name="methodName">The method name.</param>
		public void TraceMethodEntry(string moduleName, string methodName)
		{
			TraceMethodEntry(moduleName, methodName, null);
		}

		/// <summary>
		/// Traces the entry into a method.
		/// </summary>
		/// <param name="moduleName">The name of the module to which the method belongs.</param>
		/// <param name="methodName">The method name.</param>
		/// <param name="message">A message to include in the trace line.</param>
		public void TraceMethodEntry(string moduleName, string methodName, string message) 
		{
			TraceMessage(moduleName, methodName, "entry", message);
		}

		/// <summary>
		/// Traces the exit from a method.
		/// </summary>
		/// <param name="moduleName">The name of the module to which the method belongs.</param>
		/// <param name="methodName">The method name.</param>
		public void TraceMethodExit(string moduleName, string methodName)
		{
			TraceMethodExit(moduleName, methodName, null);
		}

		/// <summary>
		/// Traces the entry from a method.
		/// </summary>
		/// <param name="moduleName">The name of the module to which the method belongs.</param>
		/// <param name="methodName">The method name.</param>
		/// <param name="message">A message to include in the trace line.</param>
		public void TraceMethodExit(string moduleName, string methodName, string message) 
		{
			TraceMessage(moduleName, methodName, "exit", message);
		}

		/// <summary>
		/// Traces details about an error.
		/// </summary>
		/// <param name="moduleName">The name of the module in which the error occurred.</param>
		/// <param name="methodName">The name of the method in which the error occurred.</param>
		/// <param name="exception">The error.</param>
		public void TraceMethodError(string moduleName, string methodName, Exception exception)
		{
			TraceMethodError(moduleName, methodName, exception, null);
		}

		/// <summary>
		/// Traces details about an error.
		/// </summary>
		/// <param name="moduleName">The name of the module in which the error occurred.</param>
		/// <param name="methodName">The name of the method in which the error occurred.</param>
		/// <param name="exception">The error.</param>
		/// <param name="message">A message to include in the trace line.</param>
		public void TraceMethodError(string moduleName, string methodName, Exception exception, string message) 
		{
			if (_isTracingEnabled)
			{
				string errorMessage = "";
				if (message != null && message.Length > 0)
				{
					errorMessage = message + ", Error: ";
				}
				errorMessage +=
					"type: " + exception.GetType().ToString() +
					" description: " + exception.Message +
					" source: " + exception.Source +
					" location: " + exception.StackTrace;
				TraceMessage(moduleName, methodName, "error", errorMessage);
			}
		}

		/// <summary>
		/// Traces a message.
		/// </summary>
		/// <param name="moduleName">The name of the current module.</param>
		/// <param name="methodName">The name of the current method.</param>
		/// <param name="methodStage">The current stage in the method.</param>
		public void TraceMessage(string moduleName, string methodName, string methodStage)
		{
			TraceMessage(moduleName, methodName, methodStage, null);
		}

		/// <summary>
		/// Traces a message.
		/// </summary>
		/// <param name="moduleName">The name of the current module.</param>
		/// <param name="methodName">The name of the current method.</param>
		/// <param name="methodStage">The current stage in the method.</param>
		/// <param name="message">A message to include in the trace line.</param>
		public void TraceMessage(string moduleName, string methodName, string methodStage, string message) 
		{
			if (_isTracingEnabled)
			{
				TimeSpan duration = DateTime.Now - _startTime;
				_logger.Debug(_traceId + "," + Convert.ToString((long)duration.TotalMilliseconds) + " ms," + moduleName + "," + methodName + "," + methodStage + (message != null && message.Length > 0 ? "," + message : ""));
			}
		}

		/// <summary>
		/// Traces an xml request.
		/// </summary>
		/// <param name="request">The xml request.</param>
		/// <remarks>
		/// The xml is saved to a file with the name 
		/// <i>YYYYMMDD</i>_<i>HHMMSS</i>_<i>threadId</i>_<i>temporaryFileName</i>_<i>loggerName</i>_request.xml. 
		/// The location of the file is the same as the general log file.
		/// </remarks>
		public void TraceRequest(string request) 
		{
			TraceXML(request, "request");
		}

		/// <summary>
		/// Traces an xml response.
		/// </summary>
		/// <param name="response">The xml request.</param>
		/// <remarks>
		/// The xml is saved to a file with the name 
		/// <i>YYYYMMDD</i>_<i>HHMMSS</i>_<i>threadId</i>_<i>temporaryFileName</i>_<i>loggerName</i>_response.xml. 
		/// The location of the file is the same as the general log file.
		/// </remarks>
		public void TraceResponse(string response) 
		{
			TraceXML(response, "response");
		}

		/// <summary>
		/// Traces an xml string.
		/// </summary>
		/// <param name="xmlText">The xml string.</param>
		/// <param name="suffix">A suffix to use in the trace file name.</param>
		/// <remarks>
		/// The xml is saved to a file with the name 
		/// <i>YYYYMMDD</i>_<i>HHMMSS</i>_<i>threadId</i>_<i>temporaryFileName</i>_<i>loggerName</i>_<i>suffix</i>.xml. 
		/// The location of the file is the same as the general log file.
		/// </remarks>
		public void TraceXML(string xmlText, string suffix) 
		{
			if (_isXmlTracingEnabled)
			{
				XmlDocument xmlDocument = new XmlDocument();

				try
				{
					xmlDocument.LoadXml(xmlText);
				}
				catch (XmlException exception)
				{
					XmlElement xmlElement = xmlDocument.CreateElement("XML_ERROR_TRACE");
					xmlElement.SetAttribute("ERRORCODE", exception.GetType().ToString());
					xmlElement.SetAttribute("SOURCE", exception.Source);
					xmlElement.SetAttribute("LINE", exception.LineNumber.ToString());
					xmlElement.SetAttribute("LINEPOS", exception.LinePosition.ToString());
					xmlElement.SetAttribute("REASON", exception.Message);
					xmlElement.SetAttribute("LOCATION", exception.StackTrace);
					xmlDocument.AppendChild(xmlElement);
				}
				finally
				{
					xmlDocument.Save(_traceFolder + "\\" + _traceId + "_" + _loggerName + "_" + suffix + ".xml");
				}
			}
		}

		/// <summary>
		/// Adds the unique trace identifier to an xml error response.
		/// </summary>
		/// <param name="errorResponse">The xml error response.</param>
		/// <remarks>
		/// The RESPONSE/ERROR/DESCRIPTION in the <paramref name="errorResponse"/> is prefixed with 
		/// "[trace id: <i>traceId</i>] " where <i>traceId</i> is the unique identifier for this 
		/// TraceAssist object. 
		/// </remarks>
		public void TraceIdentErrorResponse(ref string errorResponse)
		{
			if (_isTracingEnabled)
			{
				XmlDocument xmlDocument = new XmlDocument();
				try
				{
					xmlDocument.LoadXml(errorResponse);
					XmlNode xmlNode = xmlDocument.SelectSingleNode("RESPONSE/ERROR/DESCRIPTION");
					if (xmlNode != null)
					{
						if (!xmlNode.InnerText.StartsWith("[trace id: ", StringComparison.OrdinalIgnoreCase))
						{
							xmlNode.InnerText = "[trace id: " + _traceId + "] " + xmlNode.InnerText;
							errorResponse = xmlDocument.OuterXml;
						}
					}
				}
				catch (XmlException)
				{
				}
			}
		}
	}
}

