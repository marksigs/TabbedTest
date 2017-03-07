/*
--------------------------------------------------------------------------------------------
Workfile:			TaskAutomationInstaller.cs
Copyright:			Copyright ©2005 Marlborough Stirling

Description:		Installer for the Task Autometion Services.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
PSC		03/10/2005	MAR32 Created
PSC		31/10/2005	MAR117 Correct logging and add comments
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;

namespace Vertex.Fsd.Omiga.Windows.Service.TaskAutomation
{
	/// <summary>
	/// Summary description for ProjectInstaller.
	/// </summary>
	[RunInstaller(true)]
	public class TaskAutomationInstaller : System.Configuration.Install.Installer
	{
		private System.ServiceProcess.ServiceProcessInstaller ProcessInstaller;
		private System.ServiceProcess.ServiceInstaller TaskAutomationServiceInstaller;
		private System.ServiceProcess.ServiceInstaller TaskAutomationRecoveryServiceInstaller;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TaskAutomationInstaller()
		{
			// This call is required by the Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		public override void Install(IDictionary stateSaver)
		{
			base.Install (stateSaver);
			Microsoft.Win32.RegistryKey serviceKey;
			try
			{
				serviceKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("System\\CurrentControlSet\\Services\\" + this.TaskAutomationServiceInstaller.ServiceName, true);
				serviceKey.SetValue("Description", "Processes any Omiga tasks marked for automatic processing that have reached their due date and time"); 
			
				serviceKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey ("System\\CurrentControlSet\\Services\\" + this.TaskAutomationRecoveryServiceInstaller.ServiceName, true);
				serviceKey.SetValue("Description", "Processes any Omiga tasks marked for automatic processing which have a status of failed"); 
			}
			catch (Exception e)
			{
				Console.WriteLine("An exception was thrown during service installation:\n" + e.ToString()); 
			}
		}

		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.ProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
			this.TaskAutomationServiceInstaller = new System.ServiceProcess.ServiceInstaller();
			this.TaskAutomationRecoveryServiceInstaller = new System.ServiceProcess.ServiceInstaller();
			// 
			// ProcessInstaller
			// 
			this.ProcessInstaller.Password = null;
			this.ProcessInstaller.Username = null;
			// 
			// TaskAutomationServiceInstaller
			// 
			this.TaskAutomationServiceInstaller.DisplayName = "Omiga Task Automation";
			this.TaskAutomationServiceInstaller.ServiceName = "OmigaTaskAutomation";
			// 
			// TaskAutomationRecoveryServiceInstaller
			// 
			this.TaskAutomationRecoveryServiceInstaller.DisplayName = "Omiga Task Automation Recovery";
			this.TaskAutomationRecoveryServiceInstaller.ServiceName = "OmigaTaskAutomationRecovery";
			// 
			// TaskAutomationInstaller
			// 
			this.Installers.AddRange(new System.Configuration.Install.Installer[] {
																					  this.ProcessInstaller,
																					  this.TaskAutomationServiceInstaller,
																					  this.TaskAutomationRecoveryServiceInstaller});

		}
		#endregion
	}
}
