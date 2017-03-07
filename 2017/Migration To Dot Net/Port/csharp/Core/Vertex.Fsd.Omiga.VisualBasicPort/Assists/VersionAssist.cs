/*
--------------------------------------------------------------------------------------------
Workfile:			VersionAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Registry access helper object.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		22/06/2007	First .Net version. Ported from VersionAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Gets component version information.
	/// </summary>
	public static class VersionAssist
	{
		// Missing XML comment for publicly visible type or member.
		#pragma warning disable 1591
		[Obsolete("Use GetVersionList method with no parameters instead")]
		public static string GetVersionList(string request)
		{
			return GetVersionList();
		}
		// Missing XML comment for publicly visible type or member.
		#pragma warning restore 1591

		/// <summary>
		/// Gets the version information for all components specified in the OmigaComponents combo.
		/// </summary>
		/// <returns>
		/// The version information is returned in an xml string of the form:
		/// <code>
		/// &lt;RESPONSE&gt;
		///	 &lt;COMPONENTLIST&gt;
		///		 &lt;COMPONENT&gt;
		///			 &lt;DISPLAYNAME&gt;&lt;/DISPLAYNAME&gt;
		///			 &lt;COMPONENTNAME&gt;&lt;/COMPONENTNAME&gt;
		///			 &lt;VERSION&gt;&lt;/VERSION&gt;
		///			 &lt;BUILDDATE&gt;&lt;/BUILDDATE&gt;
		///		 &lt;/COMPONENT&gt;
		///	 &lt;/COMPONENTLIST&gt;
		/// &lt;/RESPONSE&gt;
		/// </code>
		/// </returns>
		public static string GetVersionList() 
		{
			string response = null;

			try 
			{
				XmlDocument xmlOut = new XmlDocument();
				XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
				xmlOut.AppendChild(xmlResponseElement);
				xmlResponseElement.SetAttribute("TYPE", "SUCCESS");
				XmlElement xmlComponentListElement = xmlOut.CreateElement("COMPONENTLIST");
				xmlResponseElement.AppendChild(xmlComponentListElement);

				const string comboName = "OmigaComponents";

				// ------------------------------------------------------------------------------------------
				// get a list of components which should be installed
				// ------------------------------------------------------------------------------------------

				ComboCollection comboGroup = ComboAssist.GetComboGroup(comboName);

				foreach (KeyValuePair<int, List<ComboValue>> pair in comboGroup)
				{
					int valueId = pair.Key;
					List<ComboValue> comboValues = pair.Value;
					if (comboValues != null && comboValues.Count > 0)
					{
						// Only look at the first combo value for each value id.
						ComboValue comboValue = comboValues[0];
						if (comboValue != null)
						{
							string componentDescription = comboValue.ValueName;
							string componentName = comboValue.ValidationType;

							if (componentName.Length == 0)
							{
								throw new InternalErrorException("Component name not found for VALUEID " + valueId.ToString());
							}

							// Assume component is in the same directory as this component.
							string moduleFileName = App.Path + "\\" + componentName;
							int index = componentName.IndexOf('.');
							if (index > 0)
							{
								// File extension in component name on database.
								componentName = componentName.Substring(0, index - 1);
							}
							else
							{
								// No file extension given, so assume component is a DLL.
								moduleFileName = moduleFileName + ".dll";
							}
							string version = null;
							string buildDate = null;
							GetFileVersionDetails(moduleFileName, out version, out buildDate);

							XmlElement xmlComponentElement = xmlOut.CreateElement("COMPONENT");
							xmlComponentListElement.AppendChild(xmlComponentElement);

							XmlElement xmlElement = null;
							xmlElement = xmlOut.CreateElement("DISPLAYNAME");
							xmlElement.InnerText = componentDescription;
							xmlComponentElement.AppendChild(xmlElement);

							xmlElement = xmlOut.CreateElement("COMPONENTNAME");
							xmlElement.InnerText = componentName;
							xmlComponentElement.AppendChild(xmlElement);

							xmlElement = xmlOut.CreateElement("VERSION");
							xmlElement.InnerText = version;
							xmlComponentElement.AppendChild(xmlElement);

							xmlElement = xmlOut.CreateElement("BUILDDATE");
							xmlElement.InnerText = buildDate;
							xmlComponentElement.AppendChild(xmlElement);
						}
					}
				}

				response = xmlOut.OuterXml;
			}
			catch (Exception exception)
			{
				response = new ErrAssistException(exception).CreateErrorResponse();
			}

			return response;
		}

		private static bool GetFileVersionDetails(string fileName, out string version, out string buildDate) 
		{
			bool success = false;
			version = "Missing Version Information";
			buildDate = version;

			FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(fileName);

			if (fileVersionInfo.Comments.Length >= 36)
			{
				if (fileVersionInfo.Comments.StartsWith("Build:", StringComparison.OrdinalIgnoreCase))
				{
					// Comments contain version and date information, i.e., a VB component compiled by the
					// Omiga 4 build process.
					version = fileVersionInfo.Comments.Substring(7, 15);
					if (fileVersionInfo.Comments.Length >= 37)
					{
						buildDate = fileVersionInfo.Comments.Substring(25, 11);
					}
					else
					{
						buildDate = fileVersionInfo.Comments.Substring(25, 10);
					}
					success = true;
				}
			}

			if (!success)
			{
				// An external component - use File version.
				version = fileVersionInfo.FileVersion;

				// Do not use file creation/modified/accessed date as these do not reflect
				// build date.
				// buildDate = GetFileCreationDate(fileName)
				buildDate = "Not applicable";

				success = true;
			}

			return success;
		}
	}
}
