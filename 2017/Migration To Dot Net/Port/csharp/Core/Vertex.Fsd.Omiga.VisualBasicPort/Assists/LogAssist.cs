/*
--------------------------------------------------------------------------------------------
Workfile:			LogAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Component profiling.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
IK		28/08/1999	Created
MCS		29/09/1999	Extend functionality
RF		08/11/1999	Added omAU, omOrg; removed omSS
MS		23/06/2000	Added omAQ,omBC,omCust,omAIP,omCE,omTP removed omCR
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from LogAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Diagnostics;
using System.Reflection;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Constants that define the type of object being profiled.
	/// </summary>
	public enum OBJECT_TYPE 
	{
		/// <summary>
		/// Profile all objects for all components.
		/// </summary>
		otNone,
		/// <summary>
		/// Profile the data object for the current component.
		/// </summary>
		otDO,
		/// <summary>
		/// Profile the business object for the current component.
		/// </summary>
		otBO,
		/// <summary>
		/// Profile all objects for the current component.
		/// </summary>
		otAll,
	}

	/// <summary>
	/// Provides profiling of Omiga components.
	/// </summary>
	/// <remarks>
	/// <para>
	/// Profiling events are written to the Windows event log.
	/// </para>
	/// <para>
	/// Examples:
	/// <code>
	/// // Ensure profiling is switched on.
	/// LogAssist.SetProfilingOn();
	///
	/// // Example 1
	/// // Create a new LogAssist object for the current assembly.
	/// LogAssist logAssist1 = new LogAssist(System.Reflection.Assembly.GetExecutingAssembly());
	/// // Switch on profiling for the current assembly.
	/// logAssist1.SetComponentProfilingOn();
	/// // Record start time and log it.
	/// logAssist1.StartTimer("Test1", true);
	/// Thread.Sleep(5000);
	/// // Record elapsed time and log it.
	/// logAssist1.StopTimer();
	///
	/// // Example 2
	/// // This uses the IDisposable interface implemented by LogAssist.
	/// // When the LogAssist object is disposed it will automatically log the elapsed time.
	/// using (LogAssist logAssist2 = new LogAssist(System.Reflection.Assembly.GetExecutingAssembly(), "Test2"))
	/// {
	///	 Thread.Sleep(5000);
	/// }
	/// </code>
	/// </para>
	/// </remarks>
	public sealed class LogAssist : IDisposable
	{
		private DateTime _startTime;
		private DateTime _stopTime;
		private string _timerName = "";
		OBJECT_TYPE _objectType = OBJECT_TYPE.otNone;
		private AppInstance _appInstance;
		private bool _timing;
		private bool _disposed;

		/// <summary>
		/// Initializes a new instance of the <see cref="LogAssist"/> class but does not start the timer.
		/// </summary>
		/// <param name="callingAssembly">The assembly which is calling the <see cref="LogAssist"/> class.</param>
		public LogAssist(Assembly callingAssembly)
		{
			_appInstance = new AppInstance(callingAssembly);
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="LogAssist"/> class and starts the timer.
		/// </summary>
		/// <param name="callingAssembly">The assembly which is calling the <see cref="LogAssist"/> class.</param>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <remarks>
		/// The timer start event is not logged. All types of object in the component are profiled.
		/// </remarks>
		public LogAssist(Assembly callingAssembly, string timerName) 
			: this(callingAssembly, timerName, false, OBJECT_TYPE.otAll)
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="LogAssist"/> class and starts the timer.
		/// </summary>
		/// <param name="callingAssembly">The assembly which is calling the <see cref="LogAssist"/> class.</param>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <param name="logStart">Indicates whether the timer start event should be logged.</param>
		/// <remarks>
		/// All types of object in the component are profiled.
		/// </remarks>
		public LogAssist(Assembly callingAssembly, string timerName, bool logStart)
			: this(callingAssembly, timerName, logStart, OBJECT_TYPE.otAll)
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="LogAssist"/> class and starts the timer.
		/// </summary>
		/// <param name="callingAssembly">The assembly which is calling the <see cref="LogAssist"/> class.</param>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <param name="logStart">Indicates whether the timer start event should be logged.</param>
		/// <param name="objectType">Indicates the type of object in the component that should be profiled.</param>
		public LogAssist(Assembly callingAssembly, string timerName, bool logStart, OBJECT_TYPE objectType) 
			: this(callingAssembly)
		{
			StartTimer(timerName, logStart, objectType);
		}

		~LogAssist()
		{
			Dispose(false);
		}

		/// <summary>
		/// Starts the timer.
		/// </summary>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <remarks>
		/// The timer start event is not logged. All types of object in the component are profiled.
		/// </remarks>
		public void StartTimer(string timerName)
		{
			StartTimer(timerName, false, OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Starts the timer.
		/// </summary>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <param name="logStart">Indicates whether the timer start event should be logged.</param>
		/// <remarks>
		/// All types of object in the component are profiled.
		/// </remarks>
		public void StartTimer(string timerName, bool logStart)
		{
			StartTimer(timerName, logStart, OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Starts the timer.
		/// </summary>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <param name="logStart">Indicates whether the timer start event should be logged.</param>
		/// <param name="objectType">Indicates the types of object in the component that should be profiled.</param>
		/// <remarks>
		/// If already timing, then this method has no effect.
		/// </remarks>
		public void StartTimer(string timerName, bool logStart, OBJECT_TYPE objectType)
		{
			if (!_timing)
			{
				_timing = true;
				if (ProfilingOn() && (objectType == OBJECT_TYPE.otNone || IsComponentProfilingOn(_appInstance.Title, objectType)))
				{
					_timerName = timerName;
					_objectType = objectType;
					_startTime = DateTime.Now;
					if (logStart)
					{
						App.LogEvent(_timerName + " started", EventLogEntryType.Information);
					}
				}
			}
		}

		/// <summary>
		/// Stops the timer.
		/// </summary>
		public void StopTimer()
		{
			StopTimer(_objectType);
		}

		/// <summary>
		/// Stops the timer.
		/// </summary>
		/// <param name="objectType">Indicates the types of object in the component that should be profiled.</param>
		/// <remarks>
		/// If not already timing, then this method has no effect.
		/// </remarks>
		public void StopTimer(OBJECT_TYPE objectType)
		{
			if (_timing)
			{
				_timing = false;
				if (ProfilingOn() && (objectType == OBJECT_TYPE.otNone || IsComponentProfilingOn(_appInstance.Title, objectType)))
				{
					_stopTime = DateTime.Now;
					TimeSpan elapsed = _stopTime - _startTime;
					App.LogEvent(_timerName + " elapsed: " + elapsed.TotalSeconds.ToString() + " secs", EventLogEntryType.Information);
				}
			}
		}

		/// <summary>
		/// Writes a timer event to the Windows event log.
		/// </summary>
		/// <param name="timerName">The name of the timer that will appear in the log.</param>
		/// <param name="message">The message that will appear in the log.</param>
		public static void LogEvent(string timerName, string message)
		{
			App.LogEvent("[" + timerName + "] " + message, EventLogEntryType.Information);
		}

		/// <summary>
		/// Enables profiling of all objects in the component.
		/// </summary>
		public void SetComponentProfilingOn()
		{
			SetComponentProfilingOn(OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Enables profiling of certain types of objects in the component.
		/// </summary>
		/// <param name="objectType">Indicates the types of object in the component that should be profiled.</param>
		public void SetComponentProfilingOn(OBJECT_TYPE objectType)
		{
			SetComponentProfilingOn(_appInstance.Title, objectType);
		}

		/// <summary>
		/// Disables profiling of all objects in the component.
		/// </summary>
		public void SetComponentProfilingOff()
		{
			SetComponentProfilingOff(OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Disables profiling of certain types of objects in the component.
		/// </summary>
		/// <param name="objectType">Indicates the types of object in the component that should not be profiled.</param>
		public void SetComponentProfilingOff(OBJECT_TYPE objectType)
		{
			SetComponentProfilingOff(_appInstance.Title, objectType);
		}

		/// <summary>
		/// Enables profiling.
		/// </summary>
		/// <remarks>
		/// This is a global switch that controls whether profiling is enabled. 
		/// </remarks>
		public static void SetProfilingOn()
		{
			SetProfilingFlag(true);
		}

		/// <summary>
		/// Disables profiling.
		/// </summary>
		/// <remarks>
		/// This is a global switch that controls whether profiling is disabled. 
		/// </remarks>
		public static void SetProfilingOff()
		{
			SetProfilingFlag(false);
		}

		/// <summary>
		/// Enables profiling for a specified component.
		/// </summary>
		/// <param name="componentName">The name of the component.</param>
		public static void SetComponentProfilingOn(string componentName)
		{
			SetComponentProfilingOn(componentName, OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Enables profiling for specified types of object in a specified component.
		/// </summary>
		/// <param name="componentName">The name of the component.</param>
		/// <param name="objectType">Indicates the types of object in the component that should be profiled.</param>
		public static void SetComponentProfilingOn(string componentName, OBJECT_TYPE objectType)
		{
			SetComponentProfilingFlag(true, componentName, objectType);
		}

		/// <summary>
		/// Disables profiling for a specified component.
		/// </summary>
		/// <param name="componentName">The name of the component.</param>
		public static void SetComponentProfilingOff(string componentName)
		{
			SetComponentProfilingOff(componentName, OBJECT_TYPE.otAll);
		}

		/// <summary>
		/// Disables profiling for specified types of object in a specified component.
		/// </summary>
		/// <param name="componentName">The name of the component.</param>
		/// <param name="objectType">Indicates the types of object in the component that should not be profiled.</param>
		public static void SetComponentProfilingOff(string componentName, OBJECT_TYPE objectType)
		{
			SetComponentProfilingFlag(false, componentName, objectType);
		}

		private static bool IsComponentProfilingOn(string componentName, OBJECT_TYPE objectType) 
		{
			bool isComponentProfilingOn = false;

			object obj = GlobalProperty.GetSharedPropertyValue(GetPropertyName(componentName, objectType));
			if (obj is bool)
			{
				isComponentProfilingOn = (bool)obj;
			}

			return isComponentProfilingOn;
		}

		private static string GetPropertyName(string componentName, OBJECT_TYPE objectType)
		{
			return 
				objectType == OBJECT_TYPE.otAll ? 
				componentName + ".Profiling" : 
				componentName + "." + objectType.ToString() + ".Profiling";
		}

		private static void SetComponentProfilingFlag(bool flag, string componentName, OBJECT_TYPE objectType) 
		{
			GlobalProperty.SetSharedPropertyValue(GetPropertyName(componentName, objectType), flag);
		}

		private static bool ProfilingOn() 
		{
			bool isProfilingOn = false;

			object obj = GlobalProperty.GetSharedPropertyValue("Profiling");
			if (obj is bool)
			{
				isProfilingOn = (bool)obj;
			}

			return isProfilingOn;
		}

		private static void SetProfilingFlag(bool flag) 
		{
			GlobalProperty.SetSharedPropertyValue("Profiling", flag);
		}

		
		#region Obsolete.

		// Missing XML comment for publicly visible type or member.
		#pragma warning disable 1591

		[Obsolete("Use StartTimer instead")]
		public void StartTimerEx(string timerName)
		{
			StartTimer(timerName, false, OBJECT_TYPE.otDO);
		}

		[Obsolete("Use StartTimer instead")]
		public void StartTimerEx(string timerName, bool logStart, OBJECT_TYPE objectType)
		{
			StartTimer(timerName, logStart, objectType);
		}

		[Obsolete("Use StopTimer instead")]
		public void StopTimerEx()
		{
			StopTimer(_objectType);
		}

		[Obsolete("Use StopTimer instead")]
		public void StopTimerEx(OBJECT_TYPE objectType)
		{
			StopTimer(objectType);
		}

		// Missing XML comment for publicly visible type or member.
		#pragma warning restore 1591

		#endregion Obsolete.


		#region IDisposable Members

		/// <summary>
		/// Dispose of the object.
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			if (!_disposed)
			{
				if (disposing)
				{
					// Dispose managed resources.
				}

				// Dispose unmanaged resources.
				StopTimer(_objectType);
			}

			_disposed = true;
		}
		#endregion
	}
}
