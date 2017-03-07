/*
--------------------------------------------------------------------------------------------
Workfile:			TaskMonitor.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Processes incomplete task automation tasks.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
PSC		08/12/2005	MAR782 Change to use Core.ThreadPool rather than STAThreadPool
PSC		13/12/2005	MAR457 Use omLogging wrapper
PSC		19/12/2005	MAR840 Amend Stop to do nothing if the monitor has already stopped because
                    of an unexpected error
PSC		24/01/2006	MAR1112 Amend Start so that polling begins after the monitoring thread has
					started successfully
PSC		16/02/2006	MAR1290	Amend so that SQL Errors do not stop the monitoring of tasks
MHeys	10/04/2006	MAR1603	Terminate COM+ objects cleanly
MHeys	12/04/2006	MAR1603	Build issues fixed
PSC		08/06/2006	MAR1859 Add Application Priority
--------------------------------------------------------------------------------------------
*/
using System;
using System.Threading;
using System.Collections;
using System.Diagnostics;
using System.Xml;
using System.Data.SqlClient;
using System.Data;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457
using Vertex.Fsd.Omiga.Core;
using omMQ;

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// Polls the database for specific tasks to be processed. For each task found, 
	/// an MQL message is generated and put onto the queue specified in the task 
	/// definition. This message is then processed by the Omiga Task Manager component  
	/// </summary>
	internal class TaskMonitor
	{
		private readonly string _targetQueueName;
		private readonly string _messageSource = string.Empty;
		private string _omigaUserId = string.Empty;
		private string _omigaUnitId = string.Empty;
		private string _loggingContext = "TAS";
		private string _connectionString = string.Empty;

		private int _pollInterval = 10000;
		private int _noOfThreads = 1;
		private int _maxProcessRetries = 0;
		private int _maxApplicationLockRetries = 0;
		private int _processRetryInterval = 0;
		private int _applicationLockRetryInterval = 0;
		private int _omigaAuthorityLevel = 0;
		private int _tasksPerCycle = 0;

		private object _continueProcessing = true;
		private object _processingTasks = false;

		private AutoResetEvent _searchForTasks;
		private AutoResetEvent _threadPoolClosed;

		private Timer _pollTimer;

		private Thread _monitorThread;
		private Core.ThreadPool _taskProcessingThreads;			// PSC 08/12/2005 MAR782
		private PerformanceCounter _tasksProcessed = null;
		private PerformanceCounter _successfulTasks = null;
		private PerformanceCounter _failedTasks = null;
	
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		#region Properties

		/// <summary>
		/// Property for the Target Queue Name
		/// </summary>
		public string TargetQueueName
		{
			get
			{
				return _targetQueueName;
			}
		}

		/// <summary>
		/// Property for the Poll Interval
		/// </summary>
		public int PollInterval
		{
			set
			{
				_pollInterval = value;
			}
		}

		/// <summary>
		/// Property for the Threads
		/// </summary>
		public int Threads
		{
			set
			{
				_noOfThreads = value;
			}
		}

		/// <summary>
		/// Property for the Tasks Per Cycle
		/// </summary>
		public int TasksPerCycle
		{
			set
			{
				_tasksPerCycle = value;
			}
			get
			{
				return _tasksPerCycle;
			}
		}

		/// <summary>
		/// Property for the Process Retries
		/// </summary>
		public int MaximumProcessRetries
		{
			set
			{
				_maxProcessRetries = value;
			}
		}

		/// <summary>
		/// Property for the Application Lock Retries
		/// </summary>
		public int MaximumApplicationLockRetries
		{
			set
			{
				_maxApplicationLockRetries = value;
			}
		}

		/// <summary>
		/// Property for the Application Lock Retry Interval
		/// </summary>
		public int ApplicationLockRetryInterval
		{
			set
			{
				_applicationLockRetryInterval = value;
			}
		}

		/// <summary>
		/// Property for the Process Retry Interval
		/// </summary>
		public int ProcessRetryInterval
		{
			set
			{
				_processRetryInterval  = value;
			}
		}

		/// <summary>
		/// Property for the Connection String
		/// </summary>
		public string ConnectionString
		{
			set
			{
				_connectionString = value;
			}
			get
			{
				return _connectionString; 
			}
		}


		/// <summary>
		/// Property for the Omiga User Id
		/// </summary>
		public string OmigaUserId
		{
			set
			{
				_omigaUserId = value;
			}
		}

		/// <summary>
		/// Property for the Omiga Unit Id
		/// </summary>
		public string OmigaUnitId
		{
			set
			{
				_omigaUnitId = value;
			}
		}

		/// <summary>
		/// Property for the Omiga Authority Level
		/// </summary>
		public int OmigaAuthorityLevel
		{
			set
			{
				_omigaAuthorityLevel = value;
			}
		}

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
		/// Initialises a new instance of TaskMonitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		public TaskMonitor(string targetQueueName) : this(targetQueueName, string.Empty)
		{
		}

		/// <summary>
		/// Initialises a new instance of TaskMonitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		/// <param name="messageSource">The value that is to be put into the MESSAGESOURCE attribute when processing a task</param>
		public TaskMonitor(string targetQueueName, string messageSource)
		{
			if (targetQueueName == null || targetQueueName.Length == 0)
			{
				throw new ArgumentException("targetQueueName cannot be null or have a length of zero");
			}
			_targetQueueName = targetQueueName;
			_messageSource = messageSource;
			_monitorThread = new Thread(new ThreadStart(MonitorTasks));
			_searchForTasks = new AutoResetEvent(true);
			_threadPoolClosed = new AutoResetEvent(false);			
		}

		#endregion

		#region Public Methods

		/// <summary>
		/// Starts monitoring for tasks whose destination is the Target Queue Name
		/// </summary>
		public void Start()
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_targetQueueName.Length == 0)
			{
				throw new InvalidOperationException("TargetQueueName cannot be empty");
			}

			if (_connectionString.Length == 0)
			{
				throw new InvalidOperationException("ConnectionString cannot be empty");
			}

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(_targetQueueName + " starting: Poll Interval = " + _pollInterval.ToString() +
							  "ms Maximum Threads = " + _noOfThreads.ToString());
			}
			// Create counter instances
			try
			{
//				_tasksProcessed = new PerformanceCounter("Omiga Task Automation", "Tasks Processed", _targetQueueName, false);
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error("Could not create the " + _targetQueueName + " instance for the Tasks Processed performance counter", e);
				}
			}

			try
			{
//				_successfulTasks = new PerformanceCounter("Omiga Task Automation", "Tasks Successfully Processed", _targetQueueName, false);
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error("Could not create the " + _targetQueueName + " instance for the Tasks Successfully Processed performance counter", e);
				}
			}

			try
			{
//				_failedTasks = new PerformanceCounter("Omiga Task Automation", "Tasks Failed", _targetQueueName, false);
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error("Could not create the " + _targetQueueName + " instance for the Tasks Failed performance counter", e);
				}
			}
		
			// PSC 08/12/2005 MAR782 - Start
			_taskProcessingThreads = new Core.ThreadPool(_targetQueueName, _noOfThreads, ApartmentState.STA);

			// Register for thread pool events
			_taskProcessingThreads.OnWorkItemStart += new Core.ThreadPool.WorkItemStartEventHandler(OnWorkItemStart);
			_taskProcessingThreads.OnWorkItemEnd += new Core.ThreadPool.WorkItemEndEventHandler(OnWorkItemEnd);
			_taskProcessingThreads.OnWorkItemError += new Core.ThreadPool.WorkItemErrorEventHandler(OnWorkItemError);
			_taskProcessingThreads.OnWorkThreadStart += new Core.ThreadPool.WorkThreadStartEventHandler(OnWorkThreadStart);
			_taskProcessingThreads.OnWorkThreadEnd += new Core.ThreadPool.WorkThreadEndEventHandler(OnWorkThreadEnd);
			_taskProcessingThreads.OnPoolIdle += new Core.ThreadPool.PoolIdleEventHandler(OnPoolIdle);
			// PSC 08/12/2005 MAR782 - End
										
			// Start the main monitoring thread
			_monitorThread.IsBackground = true;
			_monitorThread.Name = _targetQueueName + " Monitor";
			_monitorThread.Start();

			// PSC 24/01/2006 MAR1112
			_pollTimer = new Timer(new TimerCallback(OnTimer), null, _pollInterval, _pollInterval);

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(_targetQueueName + " started");
			}
		}

		/// <summary>
		/// Stops monitoring for tasks whose destination is the Target Queue Name
		/// </summary>
		public void Stop()
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(_targetQueueName + " stopping");
			}

			// PSC 19/12/2005 MAR840 - Start
			// Only do this processing if we have not already closed down unexpectedly.
			// If have closed down unexpectedly the timer will already have been stopped
			// and the thread pool will already have been shut down
			if ((bool)_continueProcessing)
			{
				Interlocked.Exchange(ref _continueProcessing, false);

				_pollTimer.Change(Timeout.Infinite, Timeout.Infinite);
				_searchForTasks.Set();

				_threadPoolClosed.WaitOne();
			}
			else
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(_targetQueueName + " has already stopped monitoring");
				}
			}
			// PSC 19/12/2005 MAR840 - End

			// Clear up performance counters
			if (_tasksProcessed != null)
			{
				_tasksProcessed.Close();
			}

			if (_successfulTasks != null)
			{
				_successfulTasks.Close();
			}

			if (_failedTasks != null)
			{
				_failedTasks.Close();
			}

			// Unregister for thread pool events
			// PSC 08/12/2005 MAR782 - Start
			_taskProcessingThreads.OnWorkItemStart -= new Core.ThreadPool.WorkItemStartEventHandler(OnWorkItemStart);
			_taskProcessingThreads.OnWorkItemEnd -= new Core.ThreadPool.WorkItemEndEventHandler(OnWorkItemEnd);
			_taskProcessingThreads.OnWorkItemError -= new Core.ThreadPool.WorkItemErrorEventHandler(OnWorkItemError);
			_taskProcessingThreads.OnWorkThreadStart -= new Core.ThreadPool.WorkThreadStartEventHandler(OnWorkThreadStart);
			_taskProcessingThreads.OnWorkThreadEnd -= new Core.ThreadPool.WorkThreadEndEventHandler(OnWorkThreadEnd);
			_taskProcessingThreads.OnPoolIdle -= new Core.ThreadPool.PoolIdleEventHandler(OnPoolIdle);
			// PSC 08/12/2005 MAR782 - End

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug(_targetQueueName + " stopped");
			}
		}

		#endregion

		#region Private Methods

		/// <summary>
		/// Constantly monitors and processes the tasks until Stop is called
		/// </summary>
		private void MonitorTasks()
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Started to monitor tasks");
			}

			try
			{
				_searchForTasks.WaitOne();

				while ((bool)_continueProcessing)
				{
					ProcessTaskList();

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Waiting for poll event");
					}

					_searchForTasks.WaitOne();
				}
			}
			catch (Exception e)
			{
				if (_logger.IsErrorEnabled)
				{
					_logger.Error("The task monitor exited unexpectedly", e);
				}
				
				// PSC 08/12/2005 MAR782
				// Stop the timer
				_pollTimer.Change(Timeout.Infinite, Timeout.Infinite);
				// PSC 19/12/2005 MAR840
				// Set _continueProcessing to false so that it can be checked in the
				// Stop method to determine what clean up is required when the monitor
				// is told to stop
				Interlocked.Exchange(ref _continueProcessing, false);
			}
			finally
			{
				_taskProcessingThreads.CloseDown();
				_threadPoolClosed.Set();

				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Stopped monitoring tasks");
				}
			}	
		}

		/// <summary>
		/// Processes the appropriate tasks
		/// </summary>
		private void ProcessTaskList()
		{
			Interlocked.Exchange(ref _processingTasks, true);

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Processing task list");
			}

			// PSC 16/02/2006 MAR1290 - Start
			DataSet taskList = null;

			try
			{
				taskList = GetTasksToProcess();
			}
			catch (SqlException exp)
			{
				if (_logger.IsWarnEnabled)
				{
					_logger.Warn("An error occurred trying to retrieve tasks to process", exp);
				}
			}

			if (taskList != null)
			{

				foreach (DataRow taskToProcess in taskList.Tables["TaskList"].Rows)
				{
					// PSC 08/06/2006 MAR1859
					CaseTask currentTask = new CaseTask((byte[])taskToProcess["CASEACTIVITYGUID"],
						(string)taskToProcess["STAGEID"],
						Convert.ToInt32(taskToProcess["CASESTAGESEQUENCENO"].ToString()),
						(string)taskToProcess["TASKID"],
						Convert.ToInt32(taskToProcess["TASKINSTANCE"].ToString()),
						(string)taskToProcess["APPLICATIONNUMBER"],
						Convert.ToInt32(taskToProcess["APPLICATIONFACTFINDNUMBER"].ToString()),
						Convert.ToInt32(taskToProcess["APPLICATIONPRIORITY"].ToString()),
						(string)taskToProcess["ACTIVITYID"],
						Convert.ToInt32(taskToProcess["ACTIVITYINSTANCE"].ToString()),
						(string)taskToProcess["SOURCEAPPLICATION"]);

					if (_logger.IsDebugEnabled)
					{
						_logger.Debug("Putting task:- " + currentTask.ToString() + " into thread pool for processing");
					}

					// PSC 08/12/2005 MAR782
					_taskProcessingThreads.QueueUserWorkItem(new WaitCallback(ProcessIndividualTask), currentTask);	
				}
			}
			// PSC 16/02/2006 MAR1290 - End

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Processed task list");
			}
			
			Interlocked.Exchange(ref _processingTasks, false);
		}

		/// <summary>
		/// Processes an individual task
		/// </summary>
		/// <param name="state">The case task to process</param>
		private void ProcessIndividualTask(object state)
		{
			CaseTask currentTask = state as CaseTask;
			
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (currentTask == null)
			{
				throw new ArgumentException("Object passed in is not a CaseTask");
			}

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Processing task: " + currentTask.ToString());
			}

			XmlDocument requestDocument = new XmlDocument();
			XmlElement rootElement = requestDocument.CreateElement("REQUEST");
			rootElement.SetAttribute("OPERATION", "SendToQueue");
			requestDocument.AppendChild(rootElement);
			
			// Set up message element
			XmlElement messageElement = requestDocument.CreateElement("MESSAGEQUEUE");
			messageElement.SetAttribute("QUEUENAME", _targetQueueName);
			messageElement.SetAttribute("PROGID", "omTM.omTMBO");
			rootElement.AppendChild(messageElement);

			// Set up data element
			XmlElement dataElement = requestDocument.CreateElement("XML");
			messageElement.AppendChild(dataElement);

			XmlElement taskManagerRequest = requestDocument.CreateElement("REQUEST");
			taskManagerRequest.SetAttribute("USERAUTHORITYLEVEL", _omigaAuthorityLevel.ToString());
			taskManagerRequest.SetAttribute("MACHINEID", System.Environment.MachineName);
			taskManagerRequest.SetAttribute("USERID", _omigaUserId );
			taskManagerRequest.SetAttribute("UNITID", _omigaUnitId);

			taskManagerRequest.SetAttribute("OPERATION", "ProcessTASTask");
			dataElement.AppendChild (taskManagerRequest);
			
			XmlElement caseTaskElement = requestDocument.CreateElement("CASETASK");
			caseTaskElement.SetAttribute("CASEACTIVITYGUID", currentTask.CaseActivityGuidAsString);
			caseTaskElement.SetAttribute("STAGEID", currentTask.StageId);
			caseTaskElement.SetAttribute("CASESTAGESEQUENCENO", currentTask.CaseStageSequenceNo.ToString());
			caseTaskElement.SetAttribute("TASKID", currentTask.TaskId);
			caseTaskElement.SetAttribute("TASKINSTANCE", currentTask.TaskInstance.ToString());
			caseTaskElement.SetAttribute("ACTIVITYID", currentTask.ActivityId);
			caseTaskElement.SetAttribute("ACTIVITYINSTANCE", currentTask.ActivityInstance.ToString());
			caseTaskElement.SetAttribute("SOURCEAPPLICATION", currentTask.SourceApplication);
			caseTaskElement.SetAttribute("MAXPROCESSRETRIES", _maxProcessRetries.ToString());
			caseTaskElement.SetAttribute("PROCESSRETRYINTERVAL", _processRetryInterval.ToString());
			caseTaskElement.SetAttribute("MAXAPPLICATIONLOCKRETRIES", _maxApplicationLockRetries.ToString());
			caseTaskElement.SetAttribute("APPLICATIONLOCKRETRYINTERVAL", _applicationLockRetryInterval.ToString());
			caseTaskElement.SetAttribute("TARGETQUEUENAME", _targetQueueName);
			caseTaskElement.SetAttribute("CASEID", currentTask.ApplicationNumber);

			if (_messageSource.Length > 0)
			{
				caseTaskElement.SetAttribute("MESSAGESOURCE", _messageSource);
			}

			taskManagerRequest.AppendChild(caseTaskElement);

			XmlElement applicationElement = requestDocument.CreateElement("APPLICATION");
			applicationElement.SetAttribute("APPLICATIONNUMBER", currentTask.ApplicationNumber);
			applicationElement.SetAttribute("APPLICATIONFACTFINDNUMBER", currentTask.ApplicationFactFindNumber.ToString());
			// PSC 08/06/2006 MAR1859
			applicationElement.SetAttribute("APPLICATIONPRIORITY", currentTask.ApplicationPriority.ToString());
			taskManagerRequest.AppendChild(applicationElement);

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("SendToQueue Request: " + requestDocument.OuterXml);
			}

			omMQBOClass omigaMessageQueue = new omMQBOClass();
			// MAR1603 M Heys 10/04/2006 start
			Omiga4Response sendResponse;
			try
			{
				//Omiga4Response sendResponse = new Omiga4Response(omigaMessageQueue.omRequest(requestDocument.OuterXml));
				sendResponse = new Omiga4Response(omigaMessageQueue.omRequest(requestDocument.OuterXml));
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(omigaMessageQueue);
			}
			// MAR1603 M Heys 10/04/2006 end

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("SendToQueue Response: " + sendResponse.Xml);
			}

			if (!sendResponse.isSuccess())
			{
				throw new ApplicationException("Error sending data to message queue. " + sendResponse.Xml);
			}
		}

		#endregion

		#region Event Handlers

		/// <summary>
		/// Event handler for the timer to set the poll event 
		/// </summary>
		/// <param name="state"></param>
		private void OnTimer(object state)
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (!(bool)_processingTasks)
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(_targetQueueName + " poll interval reached. Setting event");
				}
				_searchForTasks.Set();
			}
			else
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug(_targetQueueName + " poll interval reached. Still processing tasks");
				}
			}
		}

		/// <summary>
		/// Event handler for the OnPoolIdle event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnPoolIdle(object sender, EventArgs e)
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Threadpool is idle");
			}
		}

		/// <summary>
		/// Event handler for the OnWorkItemStart event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnWorkItemStart(object sender, WorkItemEventArgs e)
		{
			CaseTask currentTask = e.Parameter as CaseTask;
			
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (currentTask != null)
			{
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("About to process task " + currentTask.ToString());
				}
			}
		}

		/// <summary>
		/// Event handler for the OnWorkItemEnd event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnWorkItemEnd(object sender, WorkItemEventArgs e)
		{
			CaseTask currentTask = e.Parameter as CaseTask;
			
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (currentTask != null)
			{
				if (_tasksProcessed != null)
				{
					_tasksProcessed.Increment();
				}

				if (_successfulTasks != null)
				{
					_successfulTasks.Increment();
				}
			
			
				if (_logger.IsDebugEnabled)
				{
					_logger.Debug("Processed task " + currentTask.ToString());
				}
			}
		}

		/// <summary>
		/// Event handler for the OnWorkThreadStart event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnWorkThreadStart(object sender, WorkThreadEventArgs e)
		{
			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Thread starting");
			}
		}

		/// <summary>
		/// Event handler for the OnWorkThreadEnd event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnWorkThreadEnd(object sender, WorkThreadEventArgs e)
		{
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (_logger.IsDebugEnabled)
			{
				_logger.Debug("Thread has ended");
			}
		}

		/// <summary>
		/// Event handler for the OnWorkItemError event of the STAThreadPool
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnWorkItemError(object sender, WorkItemErrorEventArgs e)
		{
			CaseTask currentTask = e.Parameter as CaseTask;

			// PSC 13/12/2005 MAR457
			omLogging.ThreadContext.Properties["LoggingContext"] = _loggingContext;

			if (currentTask != null)
			{
				if (_tasksProcessed != null)
				{
					_tasksProcessed.Increment();
				}

				if (_failedTasks != null)
				{
					_failedTasks.Increment();
				}				
			}

			if (_logger.IsErrorEnabled)
			{
				string errorMessage = "Error = " + e.ProcessingException.Message +
							          " Source = " + e.ProcessingException.Source +
							          " Stack Trace = " + e.ProcessingException.StackTrace;

				if (currentTask != null)
				{
					errorMessage = "An error occurred processing the task " + currentTask.ToString() + " " + errorMessage;
				}
				_logger.Error(errorMessage);
			}
		}

		#endregion

		#region Protected Methods

		/// <summary>
		/// Finds a list of tasks taht are to be processed for the current monitor cycle
		/// </summary>
		/// <returns>A dataset containing the tasks to be processed</returns>
		protected virtual System.Data.DataSet GetTasksToProcess()
		{
			SqlDataAdapter taskListAdapter = new SqlDataAdapter("USP_TASLOCKTASKS", _connectionString);
			taskListAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
			
			SqlParameter taskListParameter = taskListAdapter.SelectCommand.Parameters.Add("@TargetQueueName", SqlDbType.NVarChar, 20);
			taskListParameter.Value = _targetQueueName;
			
			taskListParameter = taskListAdapter.SelectCommand.Parameters.Add("@NoOfTasksToProcess", SqlDbType.Int);
			taskListParameter.Value = _tasksPerCycle;

			DataSet taskList = new DataSet();

			taskListAdapter.Fill(taskList, "TaskList");

			return taskList;
		}

		#endregion
	}
}

