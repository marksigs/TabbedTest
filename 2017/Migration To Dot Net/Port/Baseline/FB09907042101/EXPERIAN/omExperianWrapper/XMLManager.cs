using System;
using System.Xml;
 
namespace Vertex.Fsd.Omiga.omExperianWrapper.XML
{
 
	public class XMLManager
	{
  
		public string xmlGetMandatoryNodeText(XmlNode vxmlParentNode,string vstrMatchPattern,Boolean bMandatory) 
		{
			XmlNode xmlNode;
			string strValue;
			xmlNode = xmlGetMandatoryNode(vxmlParentNode, vstrMatchPattern,bMandatory);
         
   
			if (xmlNode != null)
			{
				if (xmlNode.InnerText.Length == 0) 
				{
					throw new ArgumentNullException (xmlNode.Name + "contains null value." );
				}
				else
					strValue = xmlNode.InnerText;
			}
			else
			{
				throw new ArgumentNullException (xmlNode.Name + "does not exits in the ParentNode.");
			}
 
			return strValue;
		}
    
		public XmlNode xmlGetMandatoryNode(XmlNode vxmlParentNode, string vstrMatchPattern ,Boolean bMandatory) 
		{
   
			return GetNode(vxmlParentNode, vstrMatchPattern,bMandatory);
		}
 
 
		public string xmlGetMandatoryAttributeText(XmlNode vobjNode, string vstrAttribName) 
		{
   
			XmlAttribute xmlAttrib;
			xmlAttrib = xmlGetMandatoryAttribute(vobjNode, vstrAttribName );
			if (xmlAttrib == null)
				throw new ArgumentNullException ("Missing Attribute.");
			else
			{
				if (xmlAttrib.InnerText.Length == 0)
				{
					throw new ArgumentException("not found Match pattern:- " + vstrAttribName);
				}
			}
			return  xmlAttrib.InnerText.ToString();
		}
  
      
		public XmlAttribute xmlGetMandatoryAttribute(XmlNode vobjNode, string vstrAttribName) 
		{
   
			XmlAttributeCollection attrColl = vobjNode.Attributes;
 
			// Change the value for the genre attribute
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
 
		public XmlNode xmlGetAttributeNode( XmlNode xmlnode , string AttributeName )
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
			DestNode.Attributes.SetNamedItem(xmlAttrib) ;
			//xmlAttrib.cloneNode(true);
			//xmlAttrib = null;
		}
 
		public string xmlGetNodeText(XmlNode ParentNode, string MatchPattern, Boolean Mandatory) 
		{
			XmlNode xmlNode;
			xmlNode = xmlGetNode(ParentNode, MatchPattern,Mandatory);
			if (!(xmlNode == null)) 
			{
				return xmlNode.InnerText;
			}
			else if (Mandatory == true && xmlNode == null)
			{
				throw new ApplicationException("Missing Node value");
			}
			else if (Mandatory == false && xmlNode == null)
			{
				return string.Empty ;
			}
 
			return string.Empty ;
   
    
		}
 
		public XmlNode xmlGetNode(XmlNode ParentNode, string MatchPattern,Boolean MandatoryNode)
		{                                  
			return GetNode(ParentNode, MatchPattern, MandatoryNode);
		}
 
		private XmlNode GetNode(XmlNode ParentNode, string MatchPattern, Boolean MandatoryNode) 
		{
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
 
	}
 
}
