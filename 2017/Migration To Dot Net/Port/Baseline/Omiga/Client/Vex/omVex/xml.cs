// --------------------------------------------------------------------------------------------
// Workfile:			xml.cs
// Copyright:			Copyright ©2006 Vertex Financial Services
//
// Description:		
// --------------------------------------------------------------------------------------------
// History:
// 
// Prog		Date		Description
// PE		20/11/2006	EP2_24 - Created
// PEdney	05/12/2006	EP2_321 - Environment element
// --------------------------------------------------------------------------------------------
using System;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.Reflection;

namespace Vertex.Fsd.Omiga.omVex
{
	public class XmlHelper
	{

		public static void SetNodeValue(XmlDocument Doc, string XPath, string NodeName, string NodeValue, XmlNode SiblingBefore)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());

			try
			{
				XmlNode ParentNode;
				XmlNode ValueNode;
				ParentNode = Doc.SelectSingleNode(XPath);
				
				if(ParentNode != null)
				{
					ValueNode = ParentNode.SelectSingleNode(NodeName);
					if(ValueNode == null)
					{
						ValueNode = CreateNode(Doc, ParentNode, NodeName, SiblingBefore);
					}

					ValueNode.InnerText = NodeValue;
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

		}

		private static XmlNode CreateNode(XmlDocument Doc, XmlNode ParentNode, string NodeName, XmlNode SiblingBefore)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());
			XmlNode NewNode = null;

			try
			{
				NewNode = Doc.CreateElement(NodeName);
				if(SiblingBefore==null)
				{
					ParentNode.PrependChild(NewNode);
				}
				else
				{
					ParentNode.InsertAfter(NewNode, SiblingBefore);
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			return NewNode;
		}

		public static string GetSchema(string FileName)
		{
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());
			Logger.LogStart(MethodName);
			string SchemaPath = "";

			try
			{
				SchemaPath = GetPath("\\OMIGAWEBSERVICES", "\\OMIGAWEBSERVICES\\SCHEMAS\\INTERNAL\\", FileName);

				if(SchemaPath == "")
				{
					SchemaPath = GetPath("\\WEBSERVICES", "\\WEBSERVICES\\SCHEMAS\\INTERNAL\\", FileName);
				}

				if(SchemaPath == "")
				{
					SchemaPath = GetPath("\\DLL", "\\OmigaWebServices\\Schemas\\Internal\\", FileName);
				}

				if(SchemaPath == "")
				{
					Logger.Debug(MethodName + " : Schema not found (" + FileName + ")");
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			Logger.LogFinsh(MethodName);
			return(SchemaPath);			
		}

		public static string GetXML(string FileName)
		{
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			Logger.LogStart(MethodName);
			string XMLPath = "";

			try
			{
				XMLPath = GetPath("\\OMIGAWEBSERVICES", "\\XML\\", FileName);

				if(XMLPath == "")
				{
					XMLPath = GetPath("\\WEBSERVICES", "\\EXTERNALXML\\", FileName);
				}

				if(XMLPath == "")
				{
					XMLPath = GetPath("\\DLL", "\\XML\\", FileName);
				}

				if(XMLPath == "")
				{
					Logger.Debug(MethodName + " : Schema not found (" + FileName + ")");
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			Logger.LogFinsh(MethodName);
			return(XMLPath);			
		}

		public static string GetPath(string Location1, string Replace, string FileName)
		{
			LoggingClass Logger = new LoggingClass(MethodBase.GetCurrentMethod().DeclaringType.ToString());
			string MethodName = MethodInfo.GetCurrentMethod().Name.ToString();
			Logger.LogStart(MethodName);
			int StartPosition = 0;
			string ApplicationRoot = "";
			string SchemaPath = "";
			Uri RootUri;

			try
			{
				ApplicationRoot = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).ToUpper();
				if(!Path.IsPathRooted(ApplicationRoot))
				{
					RootUri = new Uri(ApplicationRoot);
					ApplicationRoot = RootUri.LocalPath;
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			try
			{
				StartPosition = ApplicationRoot.IndexOf(Location1);
				if(StartPosition == -1)
				{
					ApplicationRoot = @"C:\\Program Files\\Marlborough Stirling\\Omiga 4\\DLL";
					StartPosition = ApplicationRoot.IndexOf(Location1);
					if(StartPosition == -1)
					{
						ApplicationRoot = @"D:\\Program Files\\Marlborough Stirling\\Omiga 4\\DLL";
						StartPosition = ApplicationRoot.IndexOf(Location1);
					}
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			try
			{
				if(StartPosition > 0)
				{
					ApplicationRoot = ApplicationRoot.Substring(0, StartPosition);				
					SchemaPath = ApplicationRoot + Replace;

					if(System.IO.Directory.Exists(SchemaPath))
					{
						SchemaPath += FileName;
					}
					else
					{
						SchemaPath = "";
					}
				}
			}
			catch(Exception NewError)
			{
				Logger.LogError(MethodName, NewError.Message);
			}

			Logger.LogFinsh(MethodName);
			return(SchemaPath);
		}

	}
}
