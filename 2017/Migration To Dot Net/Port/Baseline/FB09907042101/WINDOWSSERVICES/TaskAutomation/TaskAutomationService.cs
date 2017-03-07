/*
--------------------------------------------------------------------------------------------
Workfile:			TaskAutomationService.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Windows Service for processing incomplete automation tasks.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
PSC		08/12/2005	MAR782 Use System.Threading.ThreadPool
PSC		13/12/2005	MAR457 Use omLogging wrapper
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using Vertex.Fsd.Omiga.Core;
using System.Threading;
using System.Configuration;
using System.Reflection;
using System.Data.SqlClient;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457


namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{

	public class TaskAutomationService : System.ServiceProcess.ServiceBase
	{
		private ArrayList _taskMonitors;
		private string _loggingContext = "TAS";

		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		#region Properties

		protected string LoggingContext
		{
			set
			{
				_loggingContext = value;
			}
		}

		#endregion

		#region Constructors
	
		/// <summary>
		/// Initialises a new TaskAutomationService
		/// </summary>
		public TaskAutomationService()
		{
			// This call is required by the Windows.Forms Component Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitComponent call
		}

		#endregion

		#region Component Designer generated code

		// The main entry point for the process
		[MTAThread]
		static void Main()
		{
			System.ServiceProcess.ServiceBase[] ServicesToRun;
	
			// More than one user Service may run within the same process. To add
			// another service to this process, change the following line to
			// create a second service object. For example,
			//
			//   ServicesToRun = new System.ServiceProcess.ServiceBase[] {new Service1(), new MySecondUserService()};
			//
			ServicesToRun = new System.ServiceProcess.ServiceBase[] { new TaskAutomationService(), new TaskAutomationRecoveryService() };

			System.ServiceProcess.ServiceBase.Run(ServicesToRun);
		}

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// OmigaTaskAutomation
			// 
			this.ServiceName = "OmigaTaskAutomation";

		}
		
		#endregion

		#region Protected Methods

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// Called when the service starts
		/// </summary>
		/// <param name="args">Data passed by the start command</param>
		protected override void OnStart(string[] args)
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if(_logger.IsInfoEnabled)
			{
				_logger.Info("Service starting");
			}

			// PSC 08/12/2005 MAR782
			System.Threading.ThreadPool.QueueUserWorkItem(new WaitCallback(Initialise));

			if(_logger.IsInfoEnabled)
			{
				_logger.Info("Service started successfully");
			}
		}

		/// <summary>
		/// Called when the service stops
		/// </summary>
		protected override void OnStop()
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;
			
			if(_logger.IsInfoEnabled)
			{
				_logger.Info("Service stopping");
			}

			StopMonitors();	
	
			if(_logger.IsInfoEnabled)
			{
				_logger.Info("Service stopped successfully");
			}
		}

		/// <summary>
		/// Gets an array of TaskMonitors configured by the configuration file
		/// </summary>
		/// <returns></returns>
		protected virtual ArrayList GetTaskMonitorConfiguration()
		{
			return (ArrayList)ConfigurationSettings.GetConfig("taskMonitors");
		}

		#endregion

		#region Private Methods

		/// <summary>
		/// Initialises the individual task monitors
		/// </summary>
		/// <param name="state"></param>
		private void Initialise(object state)
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			// InitialisePerformanceCounters();
			InitialiseMonitors();
		}

		/// <summary>
		/// Shuts down the individual task monitors
		/// </summary>
		private void StopMonitors()
		{
			if (_taskMonitors != null)
			{
				if (_logger.IsInfoEnabled)
				{
					_logger.Info("Stopping monitors");
				}

				ArrayList monitorStopThreads = new ArrayList();
			
				// Tell each monitor to stop monitoring
				foreach(TaskMonitor thisTaskMonitor in _taskMonitors)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Stopping monitor for " + thisTaskMonitor.TargetQueueName);
					}
					Thread monitorStop = new Thread(new ThreadStart(thisTaskMonitor.Stop));
					monitorStop.Start();
					monitorStopThreads.Add(monitorStop);
				}

				// Wait for all monitors to stop
				foreach(Thread monitorStop in monitorStopThreads)
				{
					monitorStop.Join();
				}

				if (_logger.IsInfoEnabled)
				{
					_logger.Info("Monitors stopped");
				}

				// Remove performance counters
				if (PerformanceCounterCategory.Exists("Omiga Task Automation"))
				{
					PerformanceCounterCategory.Delete("Omiga Task Automation");
				}
			}
		}

		private void InitialisePerformanceCounters()
		{
			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Creating performance counters");
			}

			// Install performance counters
			try
			{
				if (!PerformanceCounterCategory.Exists("Omiga Task Automation"))
				{
					CounterCreationDataCollection counterData = new CounterCreationDataCollection();
					CounterCreationData thisCounter = new CounterCreationData();
				
					thisCounter.CounterName = "Tasks Processed";
					thisCounter.CounterHelp = "The number of tasks that have been processed since the service started";
					thisCounter.CounterType = PerformanceCounterType.NumberOfItems64;
					counterData.Add(thisCounter);

					thisCounter = new CounterCreationData();
					thisCounter.CounterName = "Tasks Successfully Processed";
					thisCounter.CounterHelp = "The number of tasks that have been successfully put onto their target queue since the service started";
					thisCounter.CounterType = PerformanceCounterType.NumberOfItems64;
					counterData.Add(thisCounter);

					thisCounter = new CounterCreationData();
					thisCounter.CounterName = "Tasks Failed";
					thisCounter.CounterHelp = "The number of tasks that have failed to be put onto their target queue since the service started";
					thisCounter.CounterType = PerformanceCounterType.NumberOfItems64;
					counterData.Add(thisCounter);
					PerformanceCounterCategory.Create("Omiga Task Automation", "Performance counters for the Omiga Task Automation service ", counterData);
				}

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Performance counters created");
				}
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error("An error occurred trying to install the performance counter category", e); 
				}
			}
		}

		private void InitialiseMonitors()
		{
			string connectionString = string.Empty;
			string omigaUserId = string.Empty;
			string omigaUnitId = string.Empty;
			int omigaAuthorityLevel = 0;
			bool continueProcessing = true;
			
			if (_logger.IsInfoEnabled)
			{
				_logger.Info("Starting monitors");
			}

			try
			{
				connectionString = GetConnectionString();
			}
			catch (Exception e)
			{
				if (_logger.IsFatalEnabled)
				{
					_logger.Fatal("An error occurred reading the database connection settings.", e); 
				}

				EventLog.WriteEntry(ServiceName, "An error occurred reading the database connection settings. \n" + e.Message, EventLogEntryType.Error);
				continueProcessing = false;
			}

			if (continueProcessing)
			{
				try
				{
					GetOmigaSystemUserDetails(connectionString, out omigaUserId, out omigaUnitId, out omigaAuthorityLevel);
				}
				catch (Exception e)
				{
					if (_logger.IsFatalEnabled)
					{
						_logger.Fatal("An error occurred reading the Omiga System User details.", e); 
					}

					EventLog.WriteEntry(ServiceName, "An error occurred reading the Omiga System User details. \n" + e.Message, EventLogEntryType.Error);
					continueProcessing = false;
				}
			}
				
			if (continueProcessing)
			{
				try
				{
					_taskMonitors = GetTaskMonitorConfiguration();
				
					if (_taskMonitors == null)
					{
						throw new ConfigurationException("Unable to read taskMonitor entries in the configuration file");
					}
				}
				catch (ConfigurationException e)
				{
					if (_logger.IsFatalEnabled)
					{
						_logger.Fatal("An error occurred reading the configuration settings.", e); 
					}

					EventLog.WriteEntry(ServiceName, "An error occurred reading the configuration settings. \n" + e.Message, EventLogEntryType.Error);
					continueProcessing = false;
				}
			}

			// Error reading configuration or getting global parameters
			if (!continueProcessing)
			{
				// Stop the service
				ServiceController thisService = new ServiceController("OmigaTaskAutomation");
				thisService.Stop();
				return;
			}


			foreach(TaskMonitor thisMonitor in _taskMonitors)
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Starting monitor for target queue " + thisMonitor.TargetQueueName);
				}
		
				try
				{
					thisMonitor.ConnectionString = connectionString;
					thisMonitor.OmigaUserId = omigaUserId;
					thisMonitor.OmigaUnitId = omigaUnitId;
					thisMonitor.OmigaAuthorityLevel = omigaAuthorityLevel;
					thisMonitor.Start();
				}
				catch (Exception e)
				{
					if (_logger.IsErrorEnabled)
					{
						_logger.Error("An error occurred trying to start the monitor for target queue " + thisMonitor.TargetQueueName, e); 
					}
				}
			}

			if (_logger.IsInfoEnabled)
			{
				_logger.Info("Monitors started");
			}
		}

		private string GetConnectionString()
		{
			string databaseServer = string.Empty;
			string databaseName = string.Empty;
			string userId = string.Empty;
			string password = string.Empty;
			string connectionString = string.Empty;

			Microsoft.Win32.RegistryKey connectionKey;

			// Try app.config for database connection
			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Checking app.config for database connection details");
			}
			
			databaseServer = ConfigurationSettings.AppSettings["Server"];
			databaseName = ConfigurationSettings.AppSettings["Database"];
		
			if (databaseServer != null && databaseServer.Length > 0 && 
				databaseName != null && databaseName.Length > 0)
			{
				userId = ConfigurationSettings.AppSettings["User ID"];
				password = ConfigurationSettings.AppSettings["Password"];
			}

			// Try registry for Task Automation Service database connection
			if ((databaseServer == null || databaseServer.Length == 0) && 
				(databaseName == null || databaseName.Length == 0))
			{
				try
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Checking Task Automation Service registry entry for database connection details");
					}

					connectionKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Marlborough Stirling\\Task Automation Service\\Connection", false);
					
					if (connectionKey != null)
					{
					databaseServer = (string)connectionKey.GetValue("Server");
					databaseName = (string)connectionKey.GetValue("Database Name");

						if (databaseServer != null && databaseServer.Length > 0 && 
							databaseName != null && databaseName.Length > 0)
						{
							try
							{
								userId = (string)connectionKey.GetValue("User ID");
								password = (string)connectionKey.GetValue("Password");
							}
							catch (Exception e)
							{
								if (_logger.IsDebugEnabled)
								{
									if (userId.Length == 0)
									{
										_logger.Debug("Unable to get database connection user id and/or password from Task Automation Service registry entry", e); 
									}
								}
							}
						}
					}
				}
				catch (Exception e)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Unable to get database connection details from Task Automation Service registry entry", e); 
					}
				}
			}

			// Try registry for Omiga database connection
			if ((databaseServer == null || databaseServer.Length == 0) && 
				(databaseName == null || databaseName.Length == 0))
			{
				try
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Checking Omiga4 registry entry for database connection details");
					}

					connectionKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("Software\\Omiga4\\Database Connection", false);
					
					if (connectionKey != null)
					{
						databaseServer = (string)connectionKey.GetValue("Server");
						databaseName = (string)connectionKey.GetValue("Database Name");
						
						if (databaseServer.Length > 0 && databaseName.Length > 0)
						{
							try
							{
								userId = (string)connectionKey.GetValue("User ID");
								password = (string)connectionKey.GetValue("Password");
							}
							catch (Exception e)
							{
								if (_logger.IsDebugEnabled)
								{
									if (userId.Length == 0)
									{
										_logger.Debug("Unable to get database connection user id and/or password from Omiga4 registry entry", e); 
									}
								}
							}
						}
					}
				}
				catch (Exception e)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Unable to get database connection details from Omiga4 registry entry", e); 
					}
				}
			}

			if (databaseServer != null && databaseServer.Length > 0 &
				databaseName != null & databaseName.Length > 0)
			{
	
				connectionString = "Server=" + databaseServer +
					";Database=" + databaseName;

				if (userId != null && userId.Length > 0)
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Using SQL Server security"); 
					}
					
					connectionString+= ";User Id=" + userId + 
						";Password=" + password + ";Trusted_Connection=False";
				}
				else
				{
					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Using integrated security"); 
					}

					connectionString+= ";Trusted_Connection=True";
				}
			}
			else
			{
				throw new ConfigurationException("Unable to determine the database connection details");
			}

			return connectionString;
		}

		private void GetOmigaSystemUserDetails(string connectionString, 
											   out string systemUserId,
											   out string systemUnitId,
											   out int systemAuthorityLevel)
		{
			systemUserId = string.Empty;
			systemUnitId = string.Empty;
			systemAuthorityLevel = -1;

			SqlConnection globalParamConnection = new SqlConnection(connectionString);
			SqlCommand globalParamCommand = new SqlCommand("USP_GETGLOBALPARAMETER", globalParamConnection);
			globalParamCommand.CommandType = CommandType.StoredProcedure;
			SqlParameter globalParamName = globalParamCommand.Parameters.Add("@p_Name", SqlDbType.NVarChar, 30);
			globalParamName.Value = "DefaultUserId";

			globalParamConnection.Open();
			SqlDataReader globalParamReader = globalParamCommand.ExecuteReader();

			try
			{
				if (globalParamReader.Read())
				{
					systemUserId = (string)globalParamReader["STRING"];
				}

				globalParamReader.Close();
				globalParamName.Value = "DefaultUnitId";
				
				globalParamReader = globalParamCommand.ExecuteReader();
				
				if (globalParamReader.Read())
				{
					systemUnitId = (string)globalParamReader["STRING"];
				}
				
				globalParamReader.Close();

				globalParamName.Value = "DefaultAuthorityLevel";
				
				globalParamReader = globalParamCommand.ExecuteReader();
				
				if (globalParamReader.Read())
				{
					systemAuthorityLevel = int.Parse(globalParamReader["AMOUNT"].ToString());
				}
			}
			catch (Exception e)
			{
				throw e;
			}
			finally
			{
				globalParamReader.Close();
				globalParamConnection.Close();
			}
		}

		#endregion
	}
}
