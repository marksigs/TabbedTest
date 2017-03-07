/*
--------------------------------------------------------------------------------------------
Workfile:			TaskAutomationRecoveryService.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Windows Service for processing failed automation tasks.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
PSC		13/12/2005	MAR457 Use omLogging wrapper
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Configuration;
using omLogging = Vertex.Fsd.Omiga.omLogging;  // PSC 13/12/2005 MAR457

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	public class TaskAutomationRecoveryService : Vertex.Fsd.Omiga.Windows.Service.TaskAutomation.TaskAutomationService
	{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		// PSC 13/12/2005 MAR457
		private static readonly omLogging.Logger _logger = omLogging.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		#region Constructors

		/// <summary>
		/// Initialises a new instance of TaskAutomationRecoveryService
		/// </summary>
		public TaskAutomationRecoveryService()
		{
			// This call is required by the Windows.Forms Component Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitComponent call
		}

		#endregion

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// TaskAutomationRecovery
			// 
			this.ServiceName = "OmigaTaskAutomationRecovery";

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
			LoggingContext = "TAR";
			base.OnStart (args);
		}

		/// <summary>
		/// Gets an array of TaskMonitors configured by the configuration file
		/// </summary>
		/// <returns></returns>
		protected override ArrayList GetTaskMonitorConfiguration()
		{
			return (ArrayList)ConfigurationSettings.GetConfig("recoveryTaskMonitors");
		}

		#endregion
	}
}
