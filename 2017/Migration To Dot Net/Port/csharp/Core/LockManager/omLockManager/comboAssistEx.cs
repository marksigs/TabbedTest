using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort;


namespace omLockManager
{
    public class comboAssistEx
    {

        // -------------------------------------------------------------------------------
        // BBG Specific History:
        // 
        // Prog   Date        AQR     Description
        // TK     22/11/2004  BBG1821 Performance related fixes
        // -------------------------------------------------------------------------------
        private static XmlDocument   gxmldocCombos = null;
        // ----------------------------------------------------------------------
        // BMIDS Change History
        // Prog       Date            Ref
        // GD         10/07/2002      BMIDS00165 - New Function IsValidationTypeInValidationList
        // BS         13/05/2003      BM0310 Amended test for nothing found in IsValidationTypeInValidationList
        // ----------------------------------------------------------------------
        public static string GetComboText(string vstrGroupName, short vIntValueID) 
        {
            string strPattern = String.Empty;

            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + 
                vstrGroupName + 
                "'  and  @VALUEID='" + 
                Convert.ToString(vIntValueID) + "']";
            if (!(gxmldocCombos.selectSingleNode(strPattern) == null))
            {
            }
            return xmlAssistEx.xmlGetAttributeText(gxmldocCombos.selectSingleNode(strPattern), "VALUENAME", "");
        }

        public static bool IsValidationType(string vstrGroupName, short vIntValueID, string vstrValidationType) 
        {
            bool IsValidationType = false;
            string strPattern = String.Empty;

            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + vstrGroupName + 
                "'  and  @VALUEID='" + Convert.ToString(vIntValueID) + 
                "'  and  @VALIDATIONTYPE='" + vstrValidationType + "']";
            if (!(gxmldocCombos.selectSingleNode(strPattern) == null))
            {
                IsValidationType = true;
            }
            else
            {
                IsValidationType = false;
            }
            return IsValidationType;
        }

        public static void GetValueIdsForValidationType(string vstrGroupName, string vstrValidationType, Microsoft.VisualBasic.Collection vcolValueIds) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            MSXML2.IXMLDOMNodeList xmlNodeList = null;
            string strPattern = String.Empty;

            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + 
                vstrGroupName + 
                "'  and  @VALIDATIONTYPE='" + 
                vstrValidationType + "']";
            xmlNodeList = gxmldocCombos.selectNodes(strPattern);
            foreach(MSXML2.IXMLDOMNode __each1 in xmlNodeList)
            {
                xmlNode = __each1;
                vcolValueIds.Add(xmlAssistEx.xmlGetAttributeText(xmlNode, "VALUEID", ""), null, null, null);
            }
            xmlNodeList = null;
            xmlNode = null;
        }

        public static string GetValidationTypeForValueID(string vstrGroupName, short vIntValueID) 
        {
            string strPattern = String.Empty;


            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + 
                vstrGroupName + 
                "'  and  @VALUEID= " + vIntValueID + "]";
            if (!(gxmldocCombos.selectSingleNode(strPattern) == null))
            {
            }
            return xmlAssistEx.xmlGetAttributeText(gxmldocCombos.selectSingleNode(strPattern), "VALIDATIONTYPE", "");
        }

        private static void LoadComboValidationGroup(string vstrGroupName) 
        {
            XmlElement xmlElem = null;
            XmlNode xmlComboListNode = null;
            XmlDocument xmlRequestDoc = null;
            XmlNode xmlRootNode = null;
            XmlNode xmlRequestNode = null;
            XmlNode xmlSchemaNode = null;
            // Dim xmlComboNode As IXMLDOMNode
            if (gxmldocCombos == null)
            {
                gxmldocCombos = new XmlDocument();
                xmlElem = gxmldocCombos.CreateElement("COMBOVALIDATIONLIST");
                xmlComboListNode = gxmldocCombos.AppendChild(xmlElem);
            }
            if (gxmldocCombos.SelectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" + vstrGroupName + "']") == null)
            {
                // create one FreeThreadedDOMDocument40 to contain schema & request
                xmlRequestDoc = new XmlDocument();
                // create BOGUS root node
                xmlElem = xmlRequestDoc.CreateElement("BOGUS");
                xmlRootNode = xmlRequestDoc.AppendChild(xmlElem);
                // create schema for TM_COMBOS view
                xmlElem = xmlRequestDoc.CreateElement("SCHEMA");
                xmlSchemaNode = xmlRootNode.AppendChild(xmlElem);
                xmlElem = xmlRequestDoc.CreateElement("COMBOVALIDATION");
                xmlElem.SetAttribute("ENTITYTYPE", "LOGICAL");
                xmlElem.SetAttribute("DATASRCE", "TM_COMBOVALIDATION");
                xmlSchemaNode = xmlSchemaNode.AppendChild(xmlElem);
                xmlElem = xmlRequestDoc.CreateElement("GROUPNAME");
                xmlElem.SetAttribute("DATATYPE", "STRING");
                xmlElem.SetAttribute("LENGTH", "30");
                xmlSchemaNode.AppendChild(xmlElem);
                xmlElem = xmlRequestDoc.CreateElement("VALUEID");
                xmlElem.SetAttribute("DATATYPE", "SHORT");
                xmlSchemaNode.AppendChild(xmlElem);
                xmlElem = xmlRequestDoc.CreateElement("VALIDATIONTYPE");
                xmlElem.SetAttribute("DATATYPE", "STRING");
                xmlElem.SetAttribute("LENGTH", "20");
                xmlSchemaNode.AppendChild(xmlElem);
                // create COMBOLIST request
                xmlElem = xmlRequestDoc.CreateElement("COMBOVALIDATIONLIST");
                xmlElem.SetAttribute("GROUPNAME", vstrGroupName);
                xmlRequestNode = xmlRootNode.AppendChild(xmlElem);
                if (xmlComboListNode == null)
                {
                    xmlComboListNode = gxmldocCombos.SelectSingleNode("COMBOVALIDATIONLIST");
                }
                adoAssistEx.adoGetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlComboListNode, String.Empty, String.Empty);
            }

            xmlElem = null;
            xmlComboListNode = null;
            xmlRequestDoc = null;
            xmlRootNode = null;
            xmlRequestNode = null;
            xmlSchemaNode = null;

        }

        private static void LoadComboGroup(string vstrGroupName) 
        {
            XmlElement xmlElem = null;
            XmlNode xmlComboListNode = null;
            XmlDocument xmlRequestDoc = null;
            XmlNode xmlRootNode = null;
            XmlNode xmlRequestNode = null;
            XmlNode xmlSchemaNode = null;
            // Dim xmlComboNode As IXMLDOMNode
            if (gxmldocCombos == null)
            {
                gxmldocCombos = new XmlDocument();
                xmlElem = gxmldocCombos.CreateElement("COMBOLIST");
                xmlComboListNode = gxmldocCombos.appendChild(xmlElem);
            }
            if (gxmldocCombos.selectSingleNode("COMBOLIST/COMBO[@GROUPNAME='" + vstrGroupName + "']") == null)
            {
                // create one FreeThreadedDOMDocument40 to contain schema & request
                xmlRequestDoc = new MSXML2.FreeThreadedDOMDocument40();
                // create BOGUS root node
                xmlElem = xmlRequestDoc.createElement("BOGUS");
                xmlRootNode = xmlRequestDoc.appendChild(xmlElem);
                // create schema for TM_COMBOS view
                xmlElem = xmlRequestDoc.createElement("SCHEMA");
                xmlSchemaNode = xmlRootNode.appendChild(xmlElem);
                xmlElem = xmlRequestDoc.createElement("COMBO");
                xmlElem.setAttribute("ENTITYTYPE", "LOGICAL");
                xmlElem.setAttribute("DATASRCE", "TM_COMBOS");
                xmlSchemaNode = xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlRequestDoc.createElement("GROUPNAME");
                xmlElem.setAttribute("DATATYPE", "STRING");
                xmlElem.setAttribute("LENGTH", "30");
                xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlRequestDoc.createElement("VALUEID");
                xmlElem.setAttribute("DATATYPE", "SHORT");
                xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlRequestDoc.createElement("VALUENAME");
                xmlElem.setAttribute("DATATYPE", "STRING");
                xmlElem.setAttribute("LENGTH", "50");
                xmlSchemaNode.appendChild(xmlElem);
                xmlElem = xmlRequestDoc.createElement("VALIDATIONTYPE");
                xmlElem.setAttribute("DATATYPE", "STRING");
                xmlElem.setAttribute("LENGTH", "20");
                xmlSchemaNode.appendChild(xmlElem);
                // create COMBOLIST request
                xmlElem = xmlRequestDoc.createElement("COMBOLIST");
                xmlElem.setAttribute("GROUPNAME", vstrGroupName);
                xmlRequestNode = xmlRootNode.appendChild(xmlElem);
                if (xmlComboListNode == null)
                {
                    xmlComboListNode = gxmldocCombos.selectSingleNode("COMBOLIST");
                }
                adoAssistEx.adoGetRecordSetAsXML(xmlRequestNode, xmlSchemaNode, xmlComboListNode, String.Empty, String.Empty);
            }

            xmlElem = null;
            xmlComboListNode = null;
            xmlRequestDoc = null;
            xmlRootNode = null;
            xmlRequestNode = null;
            xmlSchemaNode = null;

        }

        public static string GetFirstComboValueId(string strComboGroup, string strValidationType) 
        {
            Microsoft.VisualBasic.Collection colValueId = null;
            colValueId = new Microsoft.VisualBasic.Collection();
            GetValueIdsForValidationType(strComboGroup, strValidationType, colValueId);
            if (colValueId.Count > 0)
            {
            }
            else
            {
                errAssistEx.errThrowError("GetFirstComboValueId", (short)(OMIGAERROR.oeRecordNotFound), );
            }
            return Convert.ToString(colValueId[1 - 1]);
        }

        public static void GetValueIdsForValueName(string vstrGroupName, string vstrValueName, Microsoft.VisualBasic.Collection vcolValueIds) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            MSXML2.IXMLDOMNodeList xmlNodeList = null;
            string strPattern = String.Empty;

            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + 
                vstrGroupName + 
                "'  and  @VALUENAME='" + 
                vstrValueName + "']";
            xmlNodeList = gxmldocCombos.selectNodes(strPattern);
            foreach(MSXML2.IXMLDOMNode __each2 in xmlNodeList)
            {
                xmlNode = __each2;
                vcolValueIds.Add(xmlAssistEx.xmlGetAttributeText(xmlNode, "VALUEID", ""), null, null, null);
            }
            xmlNodeList = null;
            xmlNode = null;
        }

        public static string GetFirstComboTextId(string strComboGroup, string strValueName) 
        {
            string GetFirstComboTextId = String.Empty;
            Microsoft.VisualBasic.Collection colValueId = null;
            colValueId = new Microsoft.VisualBasic.Collection();
            GetValueIdsForValueName(strComboGroup, strValueName, colValueId);
            if (colValueId.Count > 0)
            {
                GetFirstComboTextId = Convert.ToString(colValueId[1 - 1]);
            }
            else
            {
                GetFirstComboTextId = "";
            }
            return GetFirstComboTextId;
        }

        public static bool IsValidationTypeInValidationList(string vstrGroupName, object vstrValidationType, short vIntValueID) 
        {
            string strPattern = String.Empty;
            MSXML2.IXMLDOMNodeList xmlValidationList = null;
            bool blnResult = false;
            // GD added 09/07/02
            LoadComboGroup(vstrGroupName);
            strPattern = 
                "COMBOLIST/COMBO[@GROUPNAME='" + 
                vstrGroupName + 
                "'  and  @VALUEID= " + vIntValueID + "  and  @VALIDATIONTYPE = '" + vstrValidationType + "']";
            xmlValidationList = gxmldocCombos.selectNodes(strPattern);
            // BS BM0310 13/05/03
            // If Not xmlValidationList Is Nothing Then
            if (xmlValidationList.length > 0)
            {
                blnResult = true;
            }
            else
            {
                blnResult = false;
            }
            xmlValidationList = null;
            return blnResult;
        }


    }

}
