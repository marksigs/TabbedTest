/*
--------------------------------------------------------------------------------------------
Workfile:			TaskAutomationRecoveryConfigSectionHandler.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Sets up the Recovery Task Monitors based on the settings found in 
					the configuration file.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
--------------------------------------------------------------------------------------------
*/
using System;
using System.Configuration;
using System.Xml;
using System.Collections;
using System.Globalization;

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// Sets up the Recovery Task Monitors based on the settings found in 
	/// the configuration file
	/// </summary>
	internal class TaskAutomationRecoveryConfigSectionHandler : TaskAutomationConfigSectionHandler
	{
		#region Constructors

		/// <summary>
		/// Inititalises a new instance of TaskAutomationRecoveryConfigSectionHandler
		/// </summary>
		public TaskAutomationRecoveryConfigSectionHandler()
		{
			LoggingContext = "TAR";
		}

		#endregion

		#region Protected Methods

		/// <summary>
		/// Creates the correct type of task monitor
		/// </summary>
		/// <param name="targetQueueName">The target queue name for tasks that should be monitored</param>
		/// <returns>The task monitor for this type of processing</returns>
		protected override TaskMonitor CreateTaskMonitor(string targetQueueName)
		{
			return new RecoveryTaskMonitor(targetQueueName, "TAR");
		}

		#endregion
	}
}
