/*
--------------------------------------------------------------------------------------------
Workfile:			TaskAutomationConfigSectionHandler.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Sets up the Task Monitors based on the settings found in 
					the configuration file.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
PSC		13/12/2005	MAR457 Use omLogging wrapper
--------------------------------------------------------------------------------------------
*/
using System;
using System.Configuration;
using System.Xml;
using System.Collections;
using System.Globalization;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// Sets up the Task Monitors based on the settings found in the configuration file
	/// </summary>
	internal class TaskAutomationConfigSectionHandler : IConfigurationSectionHandler
	{
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		
		private string _loggingContext = "TAS";	
	
		#region Properties

		protected string LoggingContext
		{
			set
			{
				_loggingContext = value;
			}
		}

		#endregion

		#region Public Methods

		/// <summary>
		/// Parses the XML in the configuration file to set the properties on the 
		/// create TaskMonitor object
		/// </summary>
		/// <param name="parent">The configuration settings in a corresponding parent configuration section</param>
		/// <param name="configContext">Not used</param>
		/// <param name="section">The XmlNode that contains the configuration information from the configuration file</param>
		/// <returns></returns>
		public object Create (object parent, object configContext, XmlNode section)
		{
			ArrayList monitorList = new ArrayList();
			
			XmlNodeList taskMonitors = section.SelectNodes("taskMonitor");

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Configuration file contains " + taskMonitors.Count.ToString() + " taskMonitor entries");
			}

			foreach (XmlElement taskMonitor in taskMonitors)
			{
				string targetQueueName = taskMonitor.GetAttribute("targetQueueName");
				string pollInterval = taskMonitor.GetAttribute("pollInterval");
				string threads = taskMonitor.GetAttribute("threads");
				string tasksPerCycle = taskMonitor.GetAttribute("tasksPerCycle");
				string maxProcessRetries = taskMonitor.GetAttribute("maxProcessRetries");
				string maxApplicationLockRetries = taskMonitor.GetAttribute("maxApplicationLockRetries");
				string processRetryInterval = taskMonitor.GetAttribute("processRetryInterval");
				string applicationLockRetryInterval = taskMonitor.GetAttribute("applicationLockRetryInterval");

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Configuration file contains taskMonitor entry: targetQueueName = " + targetQueueName +
						          " pollInterval = " + pollInterval + 
								  " threads = " + threads +
						          " tasksPerCycle = " + tasksPerCycle + 
						          " maxProcessRetries = " + maxProcessRetries +
						          " maxApplicationLockRetries = " + maxApplicationLockRetries +
						          " processRetryInterval = " + processRetryInterval +
						          " applicationLockRetryInterval = " + applicationLockRetryInterval); 
				}

				try
				{
					TaskMonitor currentMonitor = CreateTaskMonitor(targetQueueName);
				
					try
					{
						currentMonitor.PollInterval = int.Parse(pollInterval);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default poll interval will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The pollInterval attribute for " + targetQueueName + " is either not specified or not a valid integer. The default poll interval will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The pollInterval attribute for " + targetQueueName + " is too large. The default poll interval will be used");
						}
					}

					try
					{
						currentMonitor.Threads = int.Parse(threads);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default number of threads will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The threads attribute for " + targetQueueName + " is either not specified or not a valid integer. The default number of threads will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The threads attribute for " + targetQueueName + " is too large. The default number of threads will be used");
						}
					}

					try
					{
						currentMonitor.TasksPerCycle = int.Parse(tasksPerCycle);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default Tasks Per Cycle will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The tasksPerCycle attribute for " + targetQueueName + " is either not specified or not a valid integer. The default tasks per cycle will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The tasksPerCycle attribute for " + targetQueueName + " is too large. The default tasks per cycle will be used");
						}
					}

					try
					{
						currentMonitor.MaximumProcessRetries = int.Parse(maxProcessRetries);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default poll interval will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The maxProcessRetries attribute for " + targetQueueName + " is either not specified or not a valid integer. The default maximum process retries will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The maxProcessRetries attribute for " + targetQueueName + " is too large. The default maximum process retries will be used");
						}
					}
					try
					{
						currentMonitor.MaximumApplicationLockRetries = int.Parse(maxApplicationLockRetries);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default poll interval will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The maxApplicationLockRetries attribute for " + targetQueueName + " is either not specified or not a valid integer. The default maximum application lock retries will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The maxApplicationLockRetries attribute for " + targetQueueName + " is too large. The default maximum application lock retries will be used");
						}
					}
					try
					{
						currentMonitor.ProcessRetryInterval = int.Parse(processRetryInterval);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default poll interval will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The processRetryInterval attribute for " + targetQueueName + " is either not specified or not a valid integer. The default process retry interval will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The processRetryInterval attribute for " + targetQueueName + " is too large. The default process retry interval will be used");
						}
					}
					try
					{
						currentMonitor.ApplicationLockRetryInterval = int.Parse(applicationLockRetryInterval);
					}
					catch (ArgumentNullException)
					{
						// Catch this so the default poll interval will be used if the attribute is not set
					}
					catch (FormatException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The applicationLockRetryInterval attribute for " + targetQueueName + " is either not specified or not a valid integer. The default application lock retry interval will be used");
						}
					}
					catch (OverflowException)
					{
						if (_logger.IsWarnEnabled)
						{
							_logger.Warn("The applicationLockRetryInterval attribute for " + targetQueueName + " is too large. The default application lock retry interval will be used");
						}
					}

					monitorList.Add(currentMonitor);			
				}
				catch (ArgumentException e)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("Unable to create Task Monitor for target queue " + targetQueueName, e);
					}				
				}
			}

			return monitorList;
		}

		#endregion

		#region Protected Methods

		/// <summary>
		/// Creates the correct type of task monitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		/// <returns>The task monitor for this type of processing</returns>
		virtual protected TaskMonitor CreateTaskMonitor(string targetQueueName)
		{
			return new TaskMonitor(targetQueueName);
		}

		#endregion
	}
}
