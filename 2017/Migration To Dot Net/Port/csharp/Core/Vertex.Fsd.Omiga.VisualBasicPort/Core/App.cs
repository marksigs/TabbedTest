/*
--------------------------------------------------------------------------------------------
Workfile:			App.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		VB6 style App type.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		09/05/2007	First version.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// The App class is a replacement for the Visual Basic 6 App object, and uses the same syntax 
	/// when calling its members.
	/// </summary>
	/// <remarks>
	/// The App class is static so that it can be used with the same syntax as the 
	/// Visual Basic 6 App object, e.g., <b>App.Title</b>. Each time one of the static App class 
	/// properties is read it constructs a new AppInstance object for the calling assembly; this 
	/// AppInstance object is then used to get the required information. This has the advantage of 
	/// being able to use the same syntax as in Visual Basic 6 (because the App properties are all 
	/// static) but the disadvantage of having to create a new AppInstance object every time one of 
	/// the App properties is read. If efficiency is important then, in ported code, consider 
	/// replacing the calls to the App class with creating a single new AppInstance object, and using 
	/// it directly.
	/// </remarks>
	/// <seealso cref="AppInstance"/>
	public static class App
	{
		#region Properties
		/// <summary>
		/// The value of the AssemblyDescription attribute for the calling assembly.
		/// </summary>
		public static string Comments
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Comments; }
		}

		/// <summary>
		/// The value of the AssemblyCompany attribute for the calling assembly.
		/// </summary>
		public static string CompanyName
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).CompanyName; }
		}

		/// <summary>
		/// The root name (without the extension) for the calling assembly.
		/// </summary>
		public static string ExeName
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).ExeName; }
		}

		/// <summary>
		/// The value of the AssemblyTitle attribute for the calling assembly.
		/// </summary>
		public static string FileDescription
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).FileDescription; }
		}

		/// <summary>
		/// The value of the AssemblyCopyright attribute for the calling assembly.
		/// </summary>
		public static string LegalCopyright
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).LegalCopyright; }
		}

		/// <summary>
		/// The value of the AssemblyTrademark attribute for the calling assembly.
		/// </summary>
		public static string LegalTrademarks
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).LegalTrademarks; }
		}

		/// <summary>
		/// The major part of the version number, i.e., the first number in the AssemblyVersion attribute, for the calling assembly.
		/// </summary>
		public static int Major
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Major; }
		}

		/// <summary>
		/// The minor part of the version number, i.e., the second number in the AssemblyVersion attribute, for the calling assembly.
		/// </summary>
		public static int Minor
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Minor; }
		}

		/// <summary>
		/// The build number for the calling assembly.
		/// </summary>
		public static int Build
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Build; }
		}

		/// <summary>
		/// The revision part of the version number, i.e., the fourth number in the AssemblyVersion attribute, for the calling assembly.
		/// </summary>
		public static int Revision
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Revision; }
		}

		/// <summary>
		/// The run time directory location of the current assembly for the calling assembly.
		/// </summary>
		public static string Path
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).Path; }
		}

		/// <summary>
		/// The value of the AssemblyProduct attribute for the calling assembly.
		/// </summary>
		public static string ProductName
		{
			get { return new AppInstance(Assembly.GetCallingAssembly()).ProductName; }
		}

		/// <summary>
		/// The value of the AssemblyTitle attribute for the calling assembly.
		/// </summary>
		public static string Title
		{
			get { return new AppInstance(Assembly.GetEntryAssembly()).Title; }
		}
		#endregion Properties

		#region Public methods
		/// <summary>
		/// Logs a string to the Windows event log with a type of <b>Information</b>.
		/// </summary>
		/// <param name="logBuffer">The string to log.</param>
		public static void LogEvent(string logBuffer)
		{
			LogEvent(logBuffer, EventLogEntryType.Information);
		}

		/// <summary>
		/// Logs a string to the Windows event log.
		/// </summary>
		/// <param name="logBuffer">The string to log.</param>
		/// <param name="eventLogEntryType">The type of event to log.</param>
		/// <remarks>
		/// vbLogEventTypeError (1) = EventLogEntryType.Error, 
		/// vbLogEventTypeWarning (2) = EventLogEntryType.Warning, 
		/// vbLogEventTypeInformation (4) = EventLogEntryType.Information.
		/// </remarks>
		public static void LogEvent(string logBuffer, EventLogEntryType eventLogEntryType)
		{
			const string source = "Omiga";

			if (!EventLog.SourceExists(source))
			{
				EventLog.CreateEventSource(source, "Application");
			}

			using (EventLog eventLog = new EventLog())
			{
				eventLog.Source = source;
				eventLog.WriteEntry(logBuffer, eventLogEntryType);
			}
		}
		#endregion Public methods
	}
}
