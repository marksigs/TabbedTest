/*
--------------------------------------------------------------------------------------------
Workfile:			RecoveryTaskmonitor.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Processes failed task automation tasks.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
--------------------------------------------------------------------------------------------
*/
using System;
using System.Data.SqlClient;
using System.Data;

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// Polls the database for failed task automation tasks to be processed. 
	/// For each task found, an MQL message is generated and put onto the queue 
	/// specified in the task definition. 
	/// This message is then processed by the Omiga Task Manager component  
	/// </summary>
	internal class RecoveryTaskMonitor : TaskMonitor
	{

		#region Constructors

		/// <summary>
		/// Initialises a new instance of RecoveryTaskMonitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		public RecoveryTaskMonitor(string targetQueueName) : this(targetQueueName, string.Empty)
		{
		}

		/// <summary>
		/// Initialises a new instance of RecoveryTaskMonitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		/// <param name="messageSource">The value that is to be put into the MESSAGESOURCE attribute when processing a task</param>
		public RecoveryTaskMonitor(string targetQueueName, string messageSource) : base(targetQueueName, messageSource)
		{
			LoggingContext = "TAR";
		}

		#endregion

		#region Protected Methods

		/// <summary>
		/// Finds a list of tasks taht are to be processed for the current monitor cycle
		/// </summary>
		/// <returns>A dataset containing the tasks to be processed</returns>
		protected override System.Data.DataSet GetTasksToProcess()
		{
			SqlDataAdapter taskListAdapter = new SqlDataAdapter("USP_TASLOCKFAILEDTASKS", ConnectionString);
			taskListAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
			
			SqlParameter taskListParameter = taskListAdapter.SelectCommand.Parameters.Add("@TargetQueueName", SqlDbType.NVarChar, 20);
			taskListParameter.Value = TargetQueueName;
			
			taskListParameter = taskListAdapter.SelectCommand.Parameters.Add("@NoOfTasksToProcess", SqlDbType.Int);
			taskListParameter.Value = TasksPerCycle;

			DataSet taskList = new DataSet();

			taskListAdapter.Fill(taskList, "TaskList");

			return taskList;	
		}

		#endregion
  
	}
}
