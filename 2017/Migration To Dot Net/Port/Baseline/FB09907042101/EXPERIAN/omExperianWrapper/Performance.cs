/*
--------------------------------------------------------------------------------------------
Workfile:			Performance.cs
Copyright:			Copyright © Vertex Financial Services Ltd 2005

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
LDM		22/05/2006  EP597 Put in Adrians performance monitoring, for load testing. 
--------------------------------------------------------------------------------------------
*/

using System;
using System.Configuration;
using System.Diagnostics;
using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omExperianWrapper
{
	/// <summary>
	/// Summary description for Performance.
	/// </summary>
	public class Performance
	{
		private static bool _enabled = true;
		private static string _enablePerformanceCounters;

		public static bool Enabled 
		{ 
			get 
			{ 
				return true;
				if (_enablePerformanceCounters == null)
				{
					_enablePerformanceCounters = ConfigurationSettings.AppSettings["EnablePerformanceCounters"];
					if (_enablePerformanceCounters != null)
					{
						_enabled = _enablePerformanceCounters.ToLower() == "true";
					}
				}

				return _enabled; 
			} 
		}

		protected const string pcCategoryName = "omExperianWrapper";

		public static PerformanceCounterEx ExperianActive = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian active requests", 
			PerformanceCounterType.NumberOfItems32, 
			"The number of Experian requests being processed. This indicates the load on the interface.", 
			Enabled,
			false, 
			0);

		public static PerformanceCounterEx ExperianRequestAverageTime = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian request average time", 
			PerformanceCounterType.AverageTimer32, 
			"The average time, in seconds, taken to process a successful Experian request. This indicates the performance of the interface.", 
			Enabled,
			false, 
			0);
		public static PerformanceCounterEx ExperianRequestAverageTimeBase = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian request average time base", 
			PerformanceCounterType.AverageBase, 
			null, 
			Enabled,
			false, 
			0);

		public static PerformanceCounterEx ExperianSuccessful = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian successful", 
			PerformanceCounterType.NumberOfItems32, 
			"The number of successful Experian requests, since the process started.", 
			Enabled,
			false, 
			0);
		public static PerformanceCounterEx ExperianUnsuccessful = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian unsuccessful", 
			PerformanceCounterType.NumberOfItems32, 
			"The number of unsuccessful Experian requests, since the process started.", 
			Enabled,
			false, 
			0);

		public static PerformanceCounterEx ExperianSuccessfulPercentage = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian successful %", 
			PerformanceCounterType.RawFraction, 
			"The percentage of successful Experian requests , since the process started. 100% indicates complete success.", 
			Enabled,
			false, 
			0);
		public static PerformanceCounterEx ExperianSuccessfulPercentageBase = new PerformanceCounterEx(
			pcCategoryName,
			null,
			"Experian successful % base", 
			PerformanceCounterType.RawBase, 
			null, 
			Enabled,
			false, 
			0);

		public static void Install()
		{
			try
			{
				Uninstall();

				CounterCreationDataCollection CCDC = new CounterCreationDataCollection();
	        
				ExperianActive.Install(CCDC);
				ExperianRequestAverageTime.Install(CCDC);
				ExperianRequestAverageTimeBase.Install(CCDC);
				ExperianSuccessful.Install(CCDC);
				ExperianUnsuccessful.Install(CCDC);
				ExperianSuccessfulPercentage.Install(CCDC);
				ExperianSuccessfulPercentageBase.Install(CCDC);

				PerformanceCounterCategory.Create(pcCategoryName, "The " + pcCategoryName + " object includes counters for monitoring the Omiga to Experian interface.", CCDC);
			}
			catch (OmigaException)
			{
				throw;
			}
			catch (Exception ex)
			{
				throw new OmigaException("Error installing " + pcCategoryName + " counter category.", ex);
			}
		}

		public static void Uninstall()
		{
			try
			{
				if (PerformanceCounterCategory.Exists(pcCategoryName))
				{
					PerformanceCounterCategory.Delete(pcCategoryName);
				}
			}
			catch (Exception ex)
			{
				throw new OmigaException("Error uninstalling " + pcCategoryName + " counter category.", ex);
			}
		}
	}
}
