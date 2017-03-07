/* History
 * 
 * Prog		Date		Defect		Description
 * MV		17/10/2005	MAR182		Code Amendments
 * GHun		09/11/2005	MAR182		Changed ApplicationException to OmigaException
 * GHun		30/11/2005	MAR609		Fixed errors
 */
using System;
using System.Xml;
using Vertex.Fsd.Omiga.Core;

namespace Vertex.Fsd.Omiga.omFirstTitle.XML
{
	
	public class XMLManager
	{
		
		public string xmlGetMandatoryNodeText(XmlNode vxmlParentNode, string vstrMatchPattern) 
		{
			XmlNode xmlNode;
			string strValue;
			xmlNode = xmlGetMandatoryNode(vxmlParentNode, vstrMatchPattern);
	        
			
			if (xmlNode != null)
			{
				if (xmlNode.InnerText.Length == 0) 
				{
					throw new ArgumentNullException (xmlNode.Name + " contains null value.");
				}
				else
					strValue = xmlNode.InnerText;
			}
			else
			{
				throw new ArgumentNullException (xmlNode.Name + " does not exits in the ParentNode.");
			}

			return strValue;
		}
		  
		public XmlNode xmlGetMandatoryNode(XmlNode vxmlParentNode, string vstrMatchPattern) 
		{
			
			return GetNode(vxmlParentNode, vstrMatchPattern, true);
		}

	
		public string xmlGetMandatoryAttributeText(XmlNode vobjNode, string vstrAttribName) 
		{
			
			XmlAttribute xmlAttrib;
			xmlAttrib = xmlGetMandatoryAttribute(vobjNode, vstrAttribName);
			if (xmlAttrib == null)
				throw new ArgumentNullException ("Missing Attribute.");
			else
			{
				if (xmlAttrib.InnerXml.Length == 0)
				{
					throw new ArgumentException("not found Match pattern:- " + vstrAttribName);
				}
			}
			return  xmlAttrib.Value.ToString();
		}
		
		    
		public XmlAttribute xmlGetMandatoryAttribute(XmlNode vobjNode, string vstrAttribName) 
		{
			
			XmlAttributeCollection attrColl = vobjNode.Attributes;

			// Change the value for the genre attribute.
			XmlAttribute xmlAttrib = (XmlAttribute)attrColl.GetNamedItem(vstrAttribName);
			
			if (xmlAttrib== null)
				throw new ArgumentNullException ("Missing Attribute.");
			else 
			{
				if (xmlAttrib.Value == null || xmlAttrib.Value == "")
					throw new ArgumentException("not found Match pattern:- " + vstrAttribName);
			}
			
			return xmlAttrib;
		}
		
		public string xmlGetAttributeText(XmlNode xmlnode, string AttribName)
		{
			XmlNode xmlTempNode = xmlGetAttributeNode(xmlnode, AttribName);
		    
			if (xmlTempNode == null) 
				return "";
			else
				return xmlTempNode.Value;
    
		}

		public XmlNode xmlGetAttributeNode(XmlNode xmlnode, string AttributeName)
		{
			 return  xmlnode.Attributes.GetNamedItem(AttributeName);

		}

		public XmlNode CreateXmlNode(string NodeName,string NodeValue)
		{
			XmlDocument xnDoc = new XmlDocument();
			XmlNode xnNode = xnDoc.CreateElement(NodeName);
			xnNode.InnerText = NodeValue;
			return xnNode;
		}

		
		public void xmlSetAttributeValue(XmlNode DestNode, string AttribName, string AttribValue) 
		{
			XmlAttribute xmlAttrib;
			//xmlAttrib = DestNode.ownerDocument.createAttribute(AttribName);
			XmlDocument xnDoc = new XmlDocument();
			xmlAttrib = xnDoc.CreateAttribute(AttribName);
			xmlAttrib.Value = AttribValue;
			DestNode.Attributes.SetNamedItem(xmlAttrib);
			//xmlAttrib.cloneNode(true);
			//xmlAttrib = null;
		}

		public string xmlGetNodeText(XmlNode ParentNode, string MatchPattern) 
		{
			XmlNode xmlNode;
			xmlNode = xmlGetNode(ParentNode, MatchPattern, false);
			if (!(xmlNode == null)) 
			{
				return xmlNode.InnerText;
			}
			else			{
				return string.Empty;
			}
		}

		public XmlNode xmlGetNode(XmlNode ParentNode, string MatchPattern, bool MandatoryNode)
		{																															   
			return GetNode(ParentNode, MatchPattern, MandatoryNode);
		}

		private XmlNode GetNode(XmlNode ParentNode, string MatchPattern, bool MandatoryNode) 
		{
			//const string cstrFunctionName = "GetNode";
			XmlNode xmlNode;
			
			if (ParentNode == null) 
			{
				throw new ArgumentNullException("Missing parent node");
			}
			
			xmlNode = ParentNode.SelectSingleNode(MatchPattern);
			
			if ((MandatoryNode == true) && (xmlNode == null)) 
			{
				throw new ArgumentNullException("Missing Match pattern:- " + MatchPattern);
			}
		
			return xmlNode;
		
		}

		public XmlNode xmlGetRequestNode(XmlNode vxmlElement) 
		{
			const string cstrFunctionName = "xmlGetRequestNode";
			XmlDocument xdDoc = new XmlDocument();
			XmlNode xmlRequestNode;
			XmlNode xnNode;

			if (vxmlElement == null)
			{
				throw new OmigaException("Missing Element: " + cstrFunctionName);
			}

			if (vxmlElement.Name == "REQUEST") 
			{
				xmlRequestNode = vxmlElement.CloneNode(false);
			}
			else 
			{
				xnNode = vxmlElement.SelectSingleNode("//REQUEST");
				if (xnNode == null) 
				{
					throw new OmigaException("Missing primary tag expected REQUEST tag: " + cstrFunctionName);
				}
				xmlRequestNode = xnNode.CloneNode(false);
			}
			
			//xdDoc.DocumentElement.AppendChild(xmlRequestNode);
			
			return xmlRequestNode;
			
		}

	}

}
