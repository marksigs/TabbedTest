/*
--------------------------------------------------------------------------------------------
Workfile:			AppInstance.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Replacement for VB6 App object.
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
	/// The AppInstance class is a replacement for the Visual Basic 6 App object, 
	/// with a different syntax when calling its members.
	/// </summary>
	/// <remarks>
	/// The AppInstance class is not static as the AppInstance constructor derives information from 
	/// the current calling assembly, i.e., the assembly that constructed this AppInstance object.
	/// </remarks>
	/// <seealso cref="App"/>
	public class AppInstance
	{
		private Assembly _assembly;
		private string _comments;
		private string _companyName;
		private string _exeName;
		private string _fileDescription;
		private string _legalCopyright;
		private string _legalTrademarks;
		private int _major;
		private int _minor;
		private int _revision;
		private int _build;
		private string _path;
		private string _productName;
		private string _title;

		#region Constructors
		/// <summary>
		/// Constructs a new AppInstance object for the current calling assembly.
		/// </summary>
		public AppInstance() : this(Assembly.GetCallingAssembly())
		{
		}

		/// <summary>
		/// Constructs a new AppInstance object for a specified assembly.
		/// </summary>
		/// <param name="assembly">The assembly for which this AppInstance object returns information.</param>
		public AppInstance(Assembly assembly)
		{
			if (assembly == null)
			{
				throw new ArgumentNullException("assembly");
			}

			_assembly = assembly;
			string location = _assembly.Location;
			FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(location);

			_comments = fileVersionInfo.Comments;
			_companyName = fileVersionInfo.CompanyName;
			_exeName = System.IO.Path.GetFileNameWithoutExtension(location);
			_fileDescription = fileVersionInfo.FileDescription;
			_legalCopyright = fileVersionInfo.LegalCopyright;
			_legalTrademarks = fileVersionInfo.LegalTrademarks;
			_path = System.IO.Path.GetDirectoryName(location);
			Version version = _assembly.GetName().Version;
			_major = version.Major;
			_minor = version.Minor;
			_build = fileVersionInfo.FileBuildPart;
			_revision = version.Revision;
			_productName = fileVersionInfo.ProductName;
			_title = fileVersionInfo.FileDescription;
		}
		#endregion Constructors

		#region Properties
		/// <summary>
		/// The assembly used to initialise this AppInstance object.
		/// </summary>
		public Assembly Assembly
		{
			get { return _assembly; }
		}

		/// <summary>
		/// The value of the AssemblyDescription attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string Comments
		{
			get { return _comments; }
		}

		/// <summary>
		/// The value of the AssemblyCompany attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string CompanyName
		{
			get { return _companyName; }
		}

		/// <summary>
		/// The root name (without the extension) for the assembly to which this AppInstance object refers.
		/// </summary>
		public string ExeName
		{
			get { return _exeName; }
		}

		/// <summary>
		/// The value of the AssemblyTitle attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string FileDescription
		{
			get { return _fileDescription; }
		}

		/// <summary>
		/// The value of the AssemblyCopyright attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string LegalCopyright
		{
			get { return _legalCopyright; }
		}

		/// <summary>
		/// The value of the AssemblyTrademark attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string LegalTrademarks
		{
			get { return _legalTrademarks; }
		}

		/// <summary>
		/// The major part of the version number, i.e., the first number in the AssemblyVersion attribute, for the assembly to which this AppInstance object refers.
		/// </summary>
		public int Major
		{
			get { return _major; }
		}

		/// <summary>
		/// The minor part of the version number, i.e., the second number in the AssemblyVersion attribute, for the assembly to which this AppInstance object refers.
		/// </summary>
		public int Minor
		{
			get { return _minor; }
		}

		/// <summary>
		/// The build number for the assembly to which this AppInstance object refers.
		/// </summary>
		public int Build
		{
			get { return _build; }
		}

		/// <summary>
		/// The revision part of the version number, i.e., the fourth number in the AssemblyVersion attribute, for the assembly to which this AppInstance object refers.
		/// </summary>
		public int Revision
		{
			get { return _revision; }
		}

		/// <summary>
		/// The run time directory location for the assembly to which this AppInstance object refers.
		/// </summary>
		public string Path
		{
			get { return _path;	}
		}

		/// <summary>
		/// The value of the AssemblyProduct attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string ProductName
		{
			get { return _productName; }
		}

		/// <summary>
		/// The value of the AssemblyTitle attribute for the assembly to which this AppInstance object refers.
		/// </summary>
		public string Title
		{
			get	{ return _title; }
		}
		#endregion Properties
	}
}
