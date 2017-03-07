using System;
using MSXML2;
using ADODB;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility.VB6;
namespace omLockManager
{

    // =============================================
    // Enumeration Declaration Section
    // =============================================
    public enum DBPROVIDER {
        omiga4DBPROVIDERUnknown,
        omiga4DBPROVIDEROracle,
        omiga4DBPROVIDERMSOracle,
        // SDS LIVE00009659  22/01/2004 Support for MS and Oracle OLEDB Drivers
        omiga4DBPROVIDERSQLServer,
    }

    public class adoAssistEx
    {

        // -----------------------------------------------------------------------------
        // Prog   Date        Description
        // SR     05/03/2001  New public function 'adoGetOmigaNumberForDatabaseError'
        // AW     20/03/2001  Added 'PARAMMODE' check to adoConvertRecordSetToXML
        // AS     31/05/2001  Added adoGetValidRecordset for SQL Server Port
        // AS     11/06/2001  Fixed compile error with GENERIC_SQL=1
        // LD     12/06/2001  SYS2368 Modify code to get the connection string for Oracle and SQLServer
        // LD     11/06/2001  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
        // LD     15/06/2001  SYS2368 SQL Server Port - Guids to be generated as type adBinary in function getParam
        // LD     19/09/2001  SYS2722 SQL Server Port - Make function adoGuidToString public
        // IK     12/10/2001  SYS2803 Work Arounds for MSXML3 IXMLDOMNodeList bug
        // PSC    17/10/01    SYS2815 Allow integrated security with SQL Server
        // SG     18/06/02    SYS4889 Code error in executeGetRecordSet
        // SDS    22/01/2004   LIVE00009659  Added pieces of code to support both Microsoft and Oracle OLEDB Drivers
        // -------------------------------------------------------------------------------
        // BBG Specific History:
        // 
        // Prog   Date        AQR     Description
        // TK     22/11/2004  BBG1821 Performance related fixes
        // -------------------------------------------------------------------------------
        private static MSXML2.FreeThreadedDOMDocument40 gxmldocSchemas = null;
        private static string gstrDbConnectionString = String.Empty;
        private static short gintDbRetries = 0;
        private static DBPROVIDER genumDbProvider;
        // application name
        private const string gstrAppName = "Omiga4";
        // registry section for database connection info
        private const string gstrREGISTRY_SECTION = "Database Connection";
        // registry keys for database connection info
        private const string gstrPROVIDER_KEY = "Provider";
        private const string gstrSERVER_KEY = "Server";
        private const string gstrDATABASE_KEY = "Database Name";
        private const string gstrUID_KEY = "User ID";
        private const string gstrPASSWORD_KEY = "Password";
        private const string gstrDATA_SOURCE_KEY = "Data Source";
        // registry keys for other database info
        private const string gstrRETRIES_KEY = "Retries";
        public const string gstrSchemaNameId = "_SCHEMA_";
        public const string gstrExtractTypeId = "_EXTRACTTYPE_";
        public const string gstrOutNodeId = "_OUTNODENAME_";
        public const string gstrOrderById = "_ORDERBY_";
        public const string gstrWhereId = "_WHERE_";
        public const string gstrDataTypeId = "DATATYPE";
        public const string gstrDataTypeDerivedId = "DERIVED";
        public const string gstrDataSourceId = "DATASRCE";
        private const string gstrMODULEPREFIX = "adoAssist.";
        private const int glngENTRYNOTFOUND = -2147024894;
        // -------------------------------------------------------------------------------
        // BMIDS History
        // RF     18/11/02    BMIDS00935 Applied Core Change SYS4752 (SYSMCP0734)
        // -------------------------------------------------------------------------------
        public static void adoCreateFromNodeList(MSXML2.IXMLDOMNodeList vxmlRequestNodeList, string vstrSchemaName) 
        {
            const string strFunctionName = "adoCreateFromNodeList";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNode xmlDataNode = null;
            short intIndex = 0;
            // WARNING: On Error GOTO adoCreateFromNodeListExit is not supported
            try 
            {

                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                // IK AQR SYS2803
                // fix for MSXML bug
                for(intIndex = (short)1;
                    intIndex <= vxmlRequestNodeList.length;
                    intIndex = Convert.ToInt16(intIndex + 1))
                {
                    adoCreate(vxmlRequestNodeList.item[intIndex - 1], xmlSchemaNode);
                }
                // For Each xmlDataNode In vxmlRequestNodeList
                // adoCreate xmlDataNode, xmlSchemaNode
                // Next
                // WARNING: adoCreateFromNodeListExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                xmlDataNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoCreateFromNode(MSXML2.IXMLDOMNode vxmlRequestNode, string vstrSchemaName) 
        {
            const string strFunctionName = "adoCreateFromNode";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            // WARNING: On Error GOTO adoCreateFromNodeExit is not supported
            try 
            {

                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                adoCreate(vxmlRequestNode, xmlSchemaNode);
                // WARNING: adoCreateFromNodeExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoUpdateFromNodeList(MSXML2.IXMLDOMNodeList vxmlRequestNodeList, string vstrSchemaName) 
        {
            const string strFunctionName = "adoUpdateFromNodeList";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            short intIndex = 0;
            // WARNING: On Error GOTO adoUpdateFromNodeListExit is not supported
            try 
            {

                // Dim xmlDataNode As IXMLDOMNode
                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                // IK AQR SYS2803
                // fix for MSXML bug
                for(intIndex = (short)1;
                    intIndex <= vxmlRequestNodeList.length;
                    intIndex = Convert.ToInt16(intIndex + 1))
                {
                    adoUpdate(vxmlRequestNodeList.item[intIndex - 1], xmlSchemaNode, false);
                }
                // For Each xmlDataNode In vxmlRequestNodeList
                // adoUpdate xmlDataNode, xmlSchemaNode
                // Next
                // WARNING: adoUpdateFromNodeListExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                // Set xmlDataNode = Nothing
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoUpdateFromNode(MSXML2.IXMLDOMNode vxmlRequestNode, string vstrSchemaName) 
        {
            const string strFunctionName = "adoUpdateFromNode";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            // WARNING: On Error GOTO adoUpdateFromNodeExit is not supported
            try 
            {

                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                adoUpdate(vxmlRequestNode, xmlSchemaNode, false);
                // WARNING: adoUpdateFromNodeExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoUpdateMultipleInstancesFromNode(MSXML2.IXMLDOMNode vxmlRequestNode, string vstrSchemaName) 
        {
            const string strFunctionName = "adoUpdateMultipleInstancesFromNode";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            // WARNING: On Error GOTO adoUpdateMultipleInstancesFromNodeExit is not supported
            try 
            {

                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                adoUpdate(vxmlRequestNode, xmlSchemaNode, true);
                // WARNING: adoUpdateMultipleInstancesFromNodeExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoGetAsXML(MSXML2.IXMLDOMNode vxmlRequestNode, MSXML2.IXMLDOMNode vxmlResponseNode, string vstrSchemaName, string vstrFilter, string vstrOrderBy) 
        {
            const string strFunctionName = "adoGetAsXML";
            string strOrderBy = String.Empty;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            // WARNING: On Error GOTO adoGetAsXMLExit is not supported
            try 
            {
                strOrderBy = vstrOrderBy;
                if (xmlAssistEx.xmlAttributeValueExists(vxmlRequestNode, gstrOrderById) == true)
                {
                    if (Strings.Len(strOrderBy) > 0)
                    {
                        strOrderBy = strOrderBy + ",";
                    }
                    strOrderBy = strOrderBy + xmlAssistEx.xmlGetAttributeText(vxmlRequestNode, gstrOrderById, "");
                }
                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                adoCloneBaseEntityRefs(xmlSchemaNode);
                if (xmlAssistEx.xmlGetAttributeText(xmlSchemaNode, "ENTITYTYPE", "") == "PROCEDURE")
                {
                    adoGetStoredProcAsXML(vxmlRequestNode, xmlSchemaNode, vxmlResponseNode);
                }
                else
                {
                    adoGetRecordSetAsXML(
                        vxmlRequestNode, xmlSchemaNode, vxmlResponseNode, vstrFilter, strOrderBy);
                }

                // WARNING: adoGetAsXMLExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoCreate(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode) 
        {
            const string strFunctionName = "adoCreate";
            MSXML2.IXMLDOMNode xmlSchemaFieldNode = null;
            MSXML2.IXMLDOMAttribute xmlDataAttrib = null;
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            ADODB.Command cmd = null;
            ADODB.Parameter param = null;
            string strSQL = String.Empty;
            string strDataType = String.Empty;
            string strDataSource = String.Empty;
            string strColumns = String.Empty;
            string strValues = String.Empty;
            short intRetryCount = 0;
            short intRetryMax = 0;
            int lngRecordsAffected = 0;
            bool blnDbCmdOk = false;
            // WARNING: On Error GOTO adoCreateExit is not supported
            try 
            {

                strDataSource = xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "DATASRCE", "");
                if (Strings.Len(strDataSource) == 0)
                {
                    strDataSource = vxmlSchemaNode.nodeName;
                }
                // get any generated keys & default values
                foreach(MSXML2.IXMLDOMNode __each1 in vxmlSchemaNode.childNodes)
                {
                    xmlSchemaFieldNode = __each1;
                    if (vxmlDataNode.attributes.getNamedItem(xmlSchemaFieldNode.nodeName) == null)
                    {
                        if (!(xmlSchemaFieldNode.attributes.getNamedItem("DEFAULT") == null))
                        {
                            xmlAttrib = 
                                vxmlDataNode.ownerDocument.createAttribute(xmlSchemaFieldNode.nodeName);
                            xmlAttrib.text = 
                                xmlSchemaFieldNode.attributes.getNamedItem("DEFAULT").text;
                            vxmlDataNode.attributes.setNamedItem(xmlAttrib);
                        }
                        if (!(xmlSchemaFieldNode.attributes.getNamedItem("KEYTYPE") == null))
                        {
                            if ((xmlSchemaFieldNode.attributes.getNamedItem("KEYTYPE").text == "PRIMARY") || (xmlSchemaFieldNode.attributes.getNamedItem("KEYTYPE").text == "SECONDARY"))
                            {
                                if (!(xmlSchemaFieldNode.attributes.getNamedItem("KEYSOURCE") == null))
                                {
                                    if (xmlSchemaFieldNode.attributes.getNamedItem("KEYSOURCE").text == "GENERATED")
                                    {
                                        GetGeneratedKey(vxmlDataNode, xmlSchemaFieldNode);
                                    }
                                }
                            }
                        }
                    }
                }
                intRetryMax = adoGetDbRetries();
                cmd = new ADODB.Command();
                foreach(MSXML2.IXMLDOMAttribute __each2 in vxmlDataNode.attributes)
                {
                    xmlDataAttrib = __each2;

                    xmlSchemaFieldNode = vxmlSchemaNode.selectSingleNode(xmlDataAttrib.name);
                    if (!(xmlSchemaFieldNode == null))
                    {
                        if (Strings.Len(strColumns) > 0)
                        {
                            strColumns = strColumns + ",";
                        }
                        strColumns = strColumns + xmlDataAttrib.name;
                        if (Strings.Len(strValues) > 0)
                        {
                            strValues = strValues + ",";
                        }
                        strValues = strValues + "?";
                        param = getParam(xmlSchemaFieldNode, xmlDataAttrib, true);
                        cmd.Parameters.Append(param);
                    }
                }
                strSQL = "INSERT INTO " + strDataSource + " (" + strColumns + ") VALUES (" + strValues + ")";
                System.Diagnostics.Debug.WriteLine("adoCreate strSQL: " + strSQL);
                if (Strings.Len(strValues) > 0)
                {
                    cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                    cmd.CommandText = strSQL;
                    lngRecordsAffected = executeSQLCommand(cmd);
                }
                // If Len(strValues) > 0 Then
                // 
                // cmd.CommandType = adCmdText
                // cmd.CommandText = strSQL
                // 
                // On Error Resume Next
                // 
                // Do While blnDbCmdOk = False
                // 
                // lngRecordsAffected = executeSQLCommand(cmd)
                // 
                // If Err.Number = 0 Then
                // 
                // blnDbCmdOk = True
                // 
                // Else
                // 
                // If IsRetryError() = False Then
                // 
                // On Error GoTo adoCreateExit
                // 
                // Else
                // 
                // intRetryCount = intRetryCount + 1
                // 
                // Dim nTimer1 As Single, _
                // '                        nTimer2 As Single
                // 
                // ' DEBUG
                // App.LogEvent "adoCreate - retry [" & intRetryCount & "]: " & strSQL
                // 
                // If (intRetryCount = intRetryMax) Then
                // 
                // On Error GoTo adoCreateExit
                // 
                // End If
                // 
                // End If
                // 
                // End If
                // 
                // Loop
                // 
                // End If
                cmd = null;
                xmlSchemaFieldNode = null;
                xmlDataAttrib = null;
                xmlAttrib = null;
                param = null;

                System.Diagnostics.Debug.WriteLine("adoCreate records created: " + lngRecordsAffected);
                // WARNING: adoCreateExit: is not supported 
            }
            catch(Exception exc)
            {
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoUpdate(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, bool blnNonUniqueInstanceAllowed) 
        {
            const string strFunctionName = "adoUpdate";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNodeList xmlSchemaNodeList = null;
            MSXML2.IXMLDOMAttribute xmlSchemaAttrib = null;
            MSXML2.IXMLDOMAttribute xmlDataAttrib = null;
            ADODB.Command cmd = null;
            ADODB.Parameter param = null;
            string strSQL = String.Empty;
            string strSQLSet = String.Empty;
            string strSQLWhere = String.Empty;
            string strDataType = String.Empty;
            string strDataSource = String.Empty;
            short intLoop = 0;
            short intSetCount = 0;
            short intKeys = 0;
            short intKeyCount = 0;
            short intRetryCount = 0;
            short intRetryMax = 0;
            int lngRecordsAffected = 0;
            bool blnDbCmdOk = false;
            bool blnDoSet = false;
            // WARNING: On Error GOTO adoUpdateExit is not supported
            try 
            {

                strDataSource = xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "DATASRCE", "");
                if (Strings.Len(strDataSource) == 0)
                {
                    strDataSource = vxmlSchemaNode.nodeName;
                }
                intRetryMax = adoGetDbRetries();
                cmd = new ADODB.Command();
                strSQL = "UPDATE " + strDataSource;
                xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@DATATYPE]");
                foreach(MSXML2.IXMLDOMNode __each3 in xmlSchemaNodeList)
                {
                    xmlSchemaNode = __each3;
                    blnDoSet = true;
                    xmlSchemaAttrib = xmlSchemaNode.attributes.getNamedItem("KEYTYPE");
                    if (!(xmlSchemaAttrib == null))
                    {

                        if ((Convert.ToString(xmlSchemaAttrib.value) == "PRIMARY") || (Convert.ToString(xmlSchemaAttrib.value) == "SECONDARY"))
                        {
                            blnDoSet = false;
                        }
                    }
                    if (blnDoSet)
                    {
                        xmlDataAttrib = 
                            vxmlDataNode.attributes.getNamedItem(xmlSchemaNode.nodeName);
                        if (!(xmlDataAttrib == null))
                        {
                            if (Strings.Len(strSQLSet) != 0)
                            {
                                strSQLSet = strSQLSet + ", ";
                            }
                            strSQLSet = strSQLSet + xmlSchemaNode.nodeName + "=?";
                            param = getParam(xmlSchemaNode, xmlDataAttrib, false);
                            cmd.Parameters.Append(param);
                            intSetCount = (short)(intSetCount + 1);
                        }
                    }
                }
                xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']");
                intKeys = (short)(xmlSchemaNodeList.length);
                foreach(MSXML2.IXMLDOMNode __each4 in xmlSchemaNodeList)
                {
                    xmlSchemaNode = __each4;

                    xmlDataAttrib = 
                        vxmlDataNode.attributes.getNamedItem(xmlSchemaNode.nodeName);
                    if (!(xmlDataAttrib == null))
                    {
                        if (Convert.ToString(xmlDataAttrib.value) != "")
                        {
                            if (Strings.Len(strSQLWhere) != 0)
                            {
                                strSQLWhere = strSQLWhere + " AND ";
                            }
                            strSQLWhere = strSQLWhere + xmlSchemaNode.nodeName + "=?";
                            param = getParam(xmlSchemaNode, xmlDataAttrib, false);
                            cmd.Parameters.Append(param);
                            intKeyCount = (short)(intKeyCount + 1);
                        }
                    }
                }
                strSQL = strSQL + " SET " + strSQLSet;
                strSQL = strSQL + " WHERE " + strSQLWhere;
                System.Diagnostics.Debug.WriteLine("adoUpdate strSQL: " + strSQL);
                if ((blnNonUniqueInstanceAllowed == false) && (intKeys != intKeyCount))
                {
                    errAssistEx.errThrowError("adoUpdate", (short)(OMIGAERROR.oeXMLMissingAttribute), );
                }
                if (intSetCount > 0)
                {

                    cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                    cmd.CommandText = strSQL;
                    lngRecordsAffected = executeSQLCommand(cmd);
                }
                // On Error Resume Next
                // 
                // Do While blnDbCmdOk = False
                // 
                // cmd.CommandType = adCmdText
                // cmd.CommandText = strSQL
                // 
                // lngRecordsAffected = executeSQLCommand(cmd)
                // 
                // If Err.Number = 0 Then
                // 
                // blnDbCmdOk = True
                // 
                // Else
                // 
                // If IsRetryError() = False Then
                // 
                // On Error GoTo adoUpdateExit
                // 
                // Else
                // 
                // intRetryCount = intRetryCount + 1
                // 
                // ' DEBUG
                // App.LogEvent "adoUpdate - retry [" & intRetryCount & "]: " & strSQL
                // 
                // If (intRetryCount = intRetryMax) Then
                // 
                // On Error GoTo adoUpdateExit
                // 
                // End If
                // 
                // End If
                // 
                // End If
                // 
                // Loop
                // 
                // Set cmd = Nothing
                System.Diagnostics.Debug.WriteLine("adoUpdate records updated: " + lngRecordsAffected);
                // WARNING: adoUpdateExit: is not supported 
            }
            catch(Exception exc)
            {
                cmd = null;
                xmlSchemaNode = null;
                xmlSchemaNodeList = null;
                xmlSchemaAttrib = null;
                xmlDataAttrib = null;
                param = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoGetStoredProcAsXML(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMNode vxmlResponseNode) 
        {
            const string strFunctionName = "adoGetRecordSetAsXML";
            MSXML2.IXMLDOMNodeList xmlSchemaKeyNodeList = null;
            MSXML2.IXMLDOMNode xmlSchemaKeyNode = null;
            ADODB.Command cmd = null;
            ADODB.Recordset rst = null;
            string strSQLParamClause = String.Empty;
            // WARNING: On Error GOTO adoGetStoredProcAsXMLExit is not supported
            try 
            {
                cmd = new ADODB.Command();
                xmlSchemaKeyNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']");
                foreach(MSXML2.IXMLDOMNode __each5 in xmlSchemaKeyNodeList)
                {
                    xmlSchemaKeyNode = __each5;
                    adoAddParam(
                        xmlSchemaKeyNode, xmlAssistEx.xmlGetAttributeNode(vxmlDataNode, xmlSchemaKeyNode.nodeName), 
                        cmd);
                    if (Strings.Len(strSQLParamClause) == 0)
                    {
                        strSQLParamClause = "?";
                    }
                    else
                    {
                        strSQLParamClause = strSQLParamClause + ",?";
                    }

                }
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                cmd.CommandText = 
                    "{CALL " + 
                    vxmlSchemaNode.attributes.getNamedItem("DATASRCE").text + 
                    "(" + strSQLParamClause + ")}";
                rst = executeGetRecordSet(cmd);
                if (!(rst == null))
                {

                    adoConvertRecordSetToXML(
                        vxmlSchemaNode, 
                        vxmlResponseNode, 
                        rst, IsComboLookUpRequired(vxmlDataNode));
                }
                // WARNING: adoGetStoredProcAsXMLExit: is not supported 
            }
            catch(Exception exc)
            {
                rst = null;
                cmd = null;
                xmlSchemaKeyNodeList = null;
                xmlSchemaKeyNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoGetRecordSetAsXML(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMNode vxmlResponseNode, string vstrFilter, string vstrOrderBy) 
        {
            const string strFunctionName = "adoGetRecordSetAsXML";
            ADODB.Recordset rst = null;
            // WARNING: On Error GOTO adoGetRecordSetAsXMLExit is not supported
            try 
            {
                rst = adoGetRecordSet(vxmlDataNode, vxmlSchemaNode, vstrFilter, vstrOrderBy);
                if (!(rst == null))
                {

                    adoConvertRecordSetToXML(
                        vxmlSchemaNode, 
                        vxmlResponseNode, 
                        rst, IsComboLookUpRequired(vxmlDataNode));
                }
                // WARNING: adoGetRecordSetAsXMLExit: is not supported 
            }
            catch(Exception exc)
            {
                rst = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoCloneBaseEntityRefs(MSXML2.IXMLDOMNode vxmlSchemaNode) 
        {
            MSXML2.IXMLDOMNode xmlSchemaElement = null;
            MSXML2.IXMLDOMNode xmlBaseSchemaEntityNode = null;
            MSXML2.IXMLDOMNode xmlBaseSchemaNode = null;
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            string strEntityRef = String.Empty;

            foreach(MSXML2.IXMLDOMNode __each6 in vxmlSchemaNode.childNodes)
            {
                xmlSchemaElement = __each6;
                if ((xmlSchemaElement.attributes.getNamedItem("DATASRCE") == null) && (!(xmlSchemaElement.attributes.getNamedItem("ENTITYREF") == null)))
                {
                    strEntityRef = xmlSchemaElement.attributes.getNamedItem("ENTITYREF").text;
                    if (xmlBaseSchemaEntityNode == null)
                    {
                        xmlBaseSchemaEntityNode = adoGetSchema(strEntityRef);
                    }
                    else
                    {
                        if (xmlBaseSchemaEntityNode.nodeName != strEntityRef)
                        {
                            xmlBaseSchemaEntityNode = adoGetSchema(strEntityRef);
                        }
                    }
                    if (!((xmlBaseSchemaEntityNode == null)))
                    {
                        xmlBaseSchemaNode = 
                            xmlBaseSchemaEntityNode.selectSingleNode(xmlSchemaElement.nodeName);
                        if (!(xmlBaseSchemaNode == null))
                        {
                            foreach(MSXML2.IXMLDOMAttribute __each7 in xmlBaseSchemaNode.attributes)
                            {
                                xmlAttrib = __each7;
                                if ((xmlAttrib.name).Substring(0, 3) != "KEY")
                                {
                                    if (xmlSchemaElement.attributes.getNamedItem(xmlAttrib.name) == null)
                                    {
                                        xmlSchemaElement.attributes.setNamedItem(
                                            xmlAttrib.cloneNode(true));
                                    }
                                }
                            }
                            xmlSchemaElement.attributes.removeNamedItem("ENTITYREF");
                            // vxmlSchemaNode.replaceChild _
                            // '                        xmlBaseSchemaNode.cloneNode(False), _
                            // '                        xmlSchemaElement
                        }
                    }
                }
            }
            xmlSchemaElement = null;
            xmlBaseSchemaEntityNode = null;
            xmlBaseSchemaNode = null;
            xmlAttrib = null;
        }

        public static ADODB.Recordset adoGetRecordSet(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, string vstrFilter, string vstrOrderBy) 
        {
            const string strFunctionName = "adoGetRecordSet";
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMAttribute xmlDataAttrib = null;
            ADODB.Command cmd = null;
            ADODB.Parameter param = null;
            string strSQL = String.Empty;
            string strSQLWhere = String.Empty;
            string strPattern = String.Empty;
            string strDataSource = String.Empty;
            short intLoop = 0;
            short intRetryCount = 0;
            short intRetryMax = 0;
            bool blnDbCmdOk = false;
            string strSQLNoLock = String.Empty;
            // WARNING: On Error GOTO adoGetRecordSetExit is not supported
            try 
            {


                // RF BMIDS00935 (SYS4752)
                strDataSource = xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "DATASRCE", "");
                if (Strings.Len(strDataSource) == 0)
                {
                    strDataSource = vxmlSchemaNode.nodeName;
                }
                // RF BMIDS00935 Start
                // SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
                // This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
                if ((genumDbProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer) && (xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK", "") == "TRUE"))
                {
                    strSQLNoLock = " WITH (NOLOCK)";
                }
                // RF BMIDS00935 End
                intRetryMax = adoGetDbRetries();
                cmd = new ADODB.Command();
                if (!(vxmlDataNode == null))
                {

                    foreach(MSXML2.IXMLDOMAttribute __each8 in vxmlDataNode.attributes)
                    {
                        xmlDataAttrib = __each8;
                        strPattern = xmlDataAttrib.name + "[@DATATYPE]";
                        xmlSchemaNode = vxmlSchemaNode.selectSingleNode(strPattern);
                        if (!(xmlSchemaNode == null))
                        {
                            if (Strings.Len(strSQLWhere) != 0)
                            {
                                strSQLWhere = strSQLWhere + " AND ";
                            }
                            strSQLWhere = strSQLWhere + xmlSchemaNode.nodeName + "=?";
                            param = getParam(xmlSchemaNode, xmlDataAttrib, false);
                            cmd.Parameters.Append(param);
                        }
                    }
                }
                // used by comboAssist
                if (Strings.Len(vstrFilter) > 0 && (vstrFilter.Substring(0, 6)).ToUpper() == "SELECT")
                {

                    strSQL = vstrFilter;
                }
                else
                {

                    if (Strings.Len(vstrFilter) > 0)
                    {
                        if (Strings.Len(strSQLWhere) != 0)
                        {
                            strSQLWhere = strSQLWhere + " AND ";
                        }
                        strSQLWhere = strSQLWhere + vstrFilter;
                    }
                    if (cmd.Parameters.Count != 0)
                    {
                        // RF BMIDS00935 Start
                        // SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
                        strSQL = "SELECT * FROM " + strDataSource + strSQLNoLock + " WHERE (" + strSQLWhere + ")";
                        // RF BMIDS00935 End
                    }
                    else
                    {
                        // RF BMIDS00935 Start
                        // SYS4752 - allow schema to specify the SQL-Server (NOLOCK) hint.
                        strSQL = "SELECT * FROM " + strDataSource + strSQLNoLock;
                        // RF BMIDS00935 End
                        if (Strings.Len(strSQLWhere) != 0)
                        {
                            strSQL = strSQL + " WHERE (" + strSQLWhere + ")";
                            // JLD SYS1774
                        }
                    }
                    if (Strings.Len(vstrOrderBy) > 0)
                    {
                        strSQL = strSQL + " ORDER BY " + vstrOrderBy;
                    }
                }
                System.Diagnostics.Debug.WriteLine("adoGetRecordSet strSQL: " + strSQL);
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                cmd.CommandText = strSQL;
                // retries never work !!!
                // On Error Resume Next
                // 
                // Do While blnDbCmdOk = False
                // 
                // Set adoGetRecordSet = executeGetRecordSet(cmd)
                // 
                // If Err.Number = 0 Then
                // 
                // blnDbCmdOk = True
                // 
                // Else
                // 
                // If IsRetryError() = False Then
                // 
                // On Error GoTo adoGetRecordSetExit
                // 
                // Else
                // 
                // intRetryCount = intRetryCount + 1
                // 
                // ' DEBUG
                // App.LogEvent "adoGetRecordSet - retry [" & intRetryCount & "]: " & strSQL
                // 
                // If (intRetryCount = intRetryMax) Then
                // 
                // On Error GoTo adoGetRecordSetExit
                // 
                // End If
                // 
                // End If
                // 
                // End If
                // 
                // Loop
                if (adoGetRecordSet == null)
                {
                    System.Diagnostics.Debug.WriteLine("adoGetRecordSet records retrieved: 0");
                }
                else
                {
                    // ISSUE: Method or data member not found: 'print'
                    System.Diagnostics.Debug.print("adoGetRecordSet records retrieved: ", "", adoGetRecordSet().RecordCount);
                }
                // WARNING: adoGetRecordSetExit: is not supported 
            }
            catch(Exception exc)
            {
                cmd = null;
                xmlSchemaNode = null;
                xmlDataAttrib = null;
                param = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
            return executeGetRecordSet(cmd);
        }

        public static void adoConvertRecordSetToXML(MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMNode vxmlResponseNode, ADODB.Recordset vrst, bool vblnDoComboLookUp) 
        {
            MSXML2.IXMLDOMElement xmlElem = null;
            MSXML2.IXMLDOMNode xmlThisSchemaNode = null;
            ADODB.Field fld = null;
            string strNodeName = String.Empty;

            if (!(vxmlSchemaNode.attributes.getNamedItem("OUTNAME") == null))
            {
                strNodeName = vxmlSchemaNode.attributes.getNamedItem("OUTNAME").text;
            }
            else
            {
                strNodeName = vxmlSchemaNode.nodeName;
            }
            while(!(vrst.EOF))
            {
                if (vxmlResponseNode.nodeName == strNodeName)
                {
                    xmlElem = vxmlResponseNode;
                }
                else
                {
                    xmlElem = vxmlResponseNode.ownerDocument.createElement(strNodeName);
                }
                foreach(MSXML2.IXMLDOMNode __each9 in vxmlSchemaNode.childNodes)
                {
                    xmlThisSchemaNode = __each9;
                    fld = null;
                    if (xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "ENTITYTYPE", "") == "PROCEDURE")
                    {
                        if (xmlAssistEx.xmlGetAttributeText(xmlThisSchemaNode, "PARAMMODE", "") != "IN")
                        {
                            fld = vrst.Fields[xmlThisSchemaNode.nodeName];
                        }
                    }
                    else
                    {
                        fld = vrst.Fields[xmlThisSchemaNode.nodeName];
                    }
                    if (!(fld == null))
                    {
                        if (!(Convert.IsDBNull(fld.Value)))
                        {
                            if (!(Convert.IsDBNull(fld.Value)))
                            {
                                FieldToXml(fld, xmlElem, xmlThisSchemaNode, vblnDoComboLookUp);
                            }
                        }
                    }
                }
                if (xmlElem.attributes.length > 0)
                {
                    vxmlResponseNode.appendChild(xmlElem);
                }

                vrst.MoveNext();
            } 
            vrst.Close();
            fld = null;
            xmlElem = null;
            xmlThisSchemaNode = null;
        }

        public static ADODB.Recordset executeGetRecordSet(ADODB.Command cmd) 
        {
            const string strFunctionName = "executeGetRecordSet";
            ADODB.Connection conn = null;
            ADODB.Recordset rst = null;
            // WARNING: On Error GOTO executeGetRecordSetExit is not supported
            try 
            {

                conn = new ADODB.Connection();
                rst = new ADODB.Recordset();
                conn.ConnectionString = adoGetDbConnectString();
                conn.Open("", "", "", -1);
                cmd.ActiveConnection = conn.ConnectionString;
                rst.CursorLocation = (ADODB.CursorLocationEnum)ADODB.CursorLocationEnum.adUseClient;
                rst.CursorType = (ADODB.CursorTypeEnum)ADODB.CursorTypeEnum.adOpenStatic;
                rst.LockType = (ADODB.LockTypeEnum)ADODB.LockTypeEnum.adLockReadOnly;
                rst.Open(cmd, null, (ADODB.CursorTypeEnum)(-1), (ADODB.LockTypeEnum)(-1), -1);
                // SG 18/06/02 SYS4889
                // adoGetValidRecordset rst
                // disconnect RecordSet
                rst.ActiveConnection = null;
                conn.Close();
                if (!(rst.EOF))
                {
                    rst.MoveFirst();
                }
                // WARNING: executeGetRecordSetExit: is not supported 
            }
            catch(Exception exc)
            {
                rst = null;
                cmd = null;
                conn = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
            return rst;
        }

        private static int executeSQLCommand(ADODB.Command cmd) 
        {
            const string strFunctionName = "executeSQLCommand";
            int lngRecordsAffected = 0;
            ADODB.Connection conn = null;
            // WARNING: On Error GOTO executeSQLCommandExit is not supported
            try 
            {


                conn = new ADODB.Connection();
                conn.ConnectionString = adoGetDbConnectString();
                conn.Open("", "", "", -1);
                cmd.ActiveConnection = conn.ConnectionString;
                cmd.Execute(ref lngRecordsAffected, null, ADODB.ExecuteOptionEnum.adExecuteNoRecords);
                conn.Close();
                cmd = null;
                conn = null;
                // WARNING: executeSQLCommandExit: is not supported 
            }
            catch(Exception exc)
            {
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
            return lngRecordsAffected;
        }

        public static void adoAddParam(MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMAttribute vxmlAttrib, ADODB.Command vcmd) 
        {
            ADODB.Parameter param = null;
            param = getParam(vxmlSchemaNode, vxmlAttrib, false);
            vcmd.Parameters.Append(param);
            param = null;
        }

        private static ADODB.Parameter getParam(MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMAttribute vxmlAttrib, bool blnIsCreate) 
        {
            ADODB.Parameter param = null;
            string strDataType = String.Empty;
            string strDataValue = String.Empty;
            param = new ADODB.Parameter();

            strDataType = vxmlSchemaNode.attributes.getNamedItem("DATATYPE").text;
            switch (strDataType)
            {
                case "STRING":
                    param.Type = (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adBSTR;
                    param.Size = Strings.Len(vxmlAttrib.text);
                    break;
                case "GUID":
                    param.Type = (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adBinary;
                    param.Size = 16;
                    break;
                case "SHORT":
                case "BOOLEAN":
                case "COMBO":
                case "LONG":
                    param.Type = (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adInteger;
                    break;
                case "DOUBLE":
                case "CURRENCY":
                    param.Type = (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adDouble;
                    break;
                case "DATE":
                case "DATETIME":
                    param.Type = (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adDBTimeStamp;
                    break; 
            }
            param.Direction = (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput;
            param.Value = System.DBNull.Value;
            if (!(vxmlAttrib == null))
            {
                if (Strings.Len(vxmlAttrib.text) > 0)
                {
                    if (strDataType != "GUID")
                    {
                        param.Value = vxmlAttrib.text;
                    }
                    else
                    {
                        param.Value = GuidStringToByteArray(vxmlAttrib.text);
                    }
                }
            }
            param = null;
            return param;
        }

        private static void GetGeneratedKey(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode) 
        {
            MSXML2.IXMLDOMAttribute xmlAttrib = null;
            string strDataType = String.Empty;
            string strGuid = String.Empty;
            strDataType = xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "DATATYPE", "");
            switch (strDataType)
            {
                case "GUID":

                    xmlAttrib = 
                        vxmlDataNode.ownerDocument.createAttribute(vxmlSchemaNode.nodeName);
                    xmlAttrib.value = guidAssistEx.CreateGUID();
                    vxmlDataNode.attributes.setNamedItem(xmlAttrib);
                    xmlAttrib = null;
                    break;
                case "SHORT":

                    xmlAttrib = 
                        vxmlDataNode.ownerDocument.createAttribute(vxmlSchemaNode.nodeName);
                    xmlAttrib.value = GetNextKeySequence(vxmlDataNode, vxmlSchemaNode, vxmlSchemaNode.parentNode);
                    vxmlDataNode.attributes.setNamedItem(xmlAttrib);
                    xmlAttrib = null;
                    break; 
            }
        }

        private static short GetNextKeySequence(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, MSXML2.IXMLDOMNode vxmlSchemaParentNode) 
        {
            MSXML2.IXMLDOMNode xmlNode = null;
            ADODB.Command cmd = null;
            ADODB.Recordset rst = null;
            ADODB.Parameter param = null;
            string strSQL = String.Empty;
            string strSQLWhere = String.Empty;
            string strPattern = String.Empty;
            string strDataSource = String.Empty;
            string strSequenceFieldName = String.Empty;
            short intThisSequence = 0;
            bool blnDbCmdOk = false;
            string strSQLNoLock = String.Empty;
            // WARNING: On Error GOTO GetNextKeySequenceExit is not supported
            try 
            {


                // RF BMIDS00935 (SYS4572)
                strDataSource = xmlAssistEx.xmlGetAttributeText(vxmlSchemaParentNode, "DATASRCE", "");
                if (Strings.Len(strDataSource) == 0)
                {
                    strDataSource = vxmlSchemaParentNode.nodeName;
                }
                strSequenceFieldName = vxmlSchemaNode.nodeName;
                // RF BMIDS00935 Start
                // SYS4752 - If SQL Server and SQLNOLOCK is specified in the schema then set the NOLOCK SQL hint.
                // This will stop SQL-Server issuing shared locks on the table/page/row/key/etc.
                if ((genumDbProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer) && (xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "SQLNOLOCK", "") == "TRUE"))
                {
                    strSQLNoLock = " WITH (NOLOCK)";
                }
                // RF BMIDS00935 End
                cmd = new ADODB.Command();
                foreach(MSXML2.IXMLDOMNode __each10 in vxmlSchemaParentNode.childNodes)
                {
                    xmlNode = __each10;

                    if (!(xmlNode.attributes.getNamedItem("KEYTYPE") == null))
                    {

                        if (xmlNode.attributes.getNamedItem("KEYSOURCE") == null)
                        {

                            if (!(vxmlDataNode.attributes.getNamedItem(xmlNode.nodeName) == null))
                            {
                                if (Strings.Len(strSQLWhere) != 0)
                                {
                                    strSQLWhere = strSQLWhere + " AND ";
                                }
                                strSQLWhere = strSQLWhere + xmlNode.nodeName + "=?";
                                param = getParam(xmlNode, 
                                    vxmlDataNode.attributes.getNamedItem(xmlNode.nodeName), false);
                                cmd.Parameters.Append(param);
                            }
                        }
                    }
                }
                // RF BMIDS00935 Start
                // SYS4752 - Allow schema to specify the SQL-Server (NOLOCK) hint.
                strSQL = "SELECT MAX(" + strSequenceFieldName + ") FROM " + strDataSource + strSQLNoLock + " WHERE (" + strSQLWhere + ")";
                // RF BMIDS00935 End
                System.Diagnostics.Debug.WriteLine("adoAssist.GetNextKeySequence strSQL: " + strSQL);
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                cmd.CommandText = strSQL;
                rst = executeGetRecordSet(cmd);
                if (!(rst == null))
                {
                    rst.MoveFirst();
                    if (Convert.IsDBNull(rst.Fields[0 - 1].Value) == false)
                    {
                        intThisSequence = Convert.ToInt16(rst.Fields[0].Value);
                    }
                    rst.Close();
                }
                Information.Err().Clear();
                // WARNING: GetNextKeySequenceExit: is not supported 
            }
            catch(Exception exc)
            {
                rst = null;
                cmd = null;
                xmlNode = null;
                param = null;
                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
                }
            }
            return intThisSequence + 1;
        }

        public static string adoGuidToString(byte[] bytArray) 
        {
            short i = 0;
            string strGuid = String.Empty;
            strGuid = "";
            for(i = (short)0;
                i <= 15.0;
                i = Convert.ToInt16(i + 1))
            {
                strGuid = strGuid + (Conversion.Hex(bytArray[i] + 0x100)).Substring((Conversion.Hex(bytArray[i] + 0x100)).Length - 2);
            }
            return strGuid;
        }

        private static string adoDateToString(object vvarDate) 
        {

            return (Conversion.Str(DateAndTime.DatePart("d", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Substring((Conversion.Str(DateAndTime.DatePart("d", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Length - 2) + "/" + (Conversion.Str(DateAndTime.DatePart("m", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Substring((Conversion.Str(DateAndTime.DatePart("m", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Length - 2) + "/" + (Conversion.Str(DateAndTime.DatePart("yyyy", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System))).Substring((Conversion.Str(DateAndTime.DatePart("yyyy", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System))).Length - 4);
        }

        private static string adoDateTimeToString(object vvarDate) 
        {

            return adoDateToString(vvarDate) + " " + 
                (Conversion.Str(DateAndTime.DatePart("h", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Substring((Conversion.Str(DateAndTime.DatePart("h", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Length - 2) + ":" + 
                (Conversion.Str(DateAndTime.DatePart("n", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Substring((Conversion.Str(DateAndTime.DatePart("n", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Length - 2) + ":" + 
                (Conversion.Str(DateAndTime.DatePart("s", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Substring((Conversion.Str(DateAndTime.DatePart("s", Convert.ToDateTime(vvarDate), Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.System) + 100)).Length - 2);
        }

        private static void FieldToXml(ADODB.Field vfld, MSXML2.IXMLDOMElement vxmlOutElem, MSXML2.IXMLDOMNode vxmlSchemaNode, bool vblnDoComboLookUp) 
        {
            string strElementValue = String.Empty;
            string strComboGroup = String.Empty;
            string strComboValue = String.Empty;
            string strFormatMask = String.Empty;
            string strFormatted = String.Empty;
            // Dim xmlAttrib As IXMLDOMAttribute

            switch (vxmlSchemaNode.attributes.getNamedItem("DATATYPE").text)
            {
                case "GUID":
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, adoGuidToString(Convert.ToByte(vfld.Value)));
                    break;
                case "CURRENCY":

                    strFormatted = Support.Format(vfld.Value, "0.0000000", (FirstDayOfWeek)1, (FirstWeekOfYear)1);
                    strFormatted = strFormatted.Substring(0, Strings.InStr(1, strFormatted, ".", CompareMethod.Binary) + 2);
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, 
                        strFormatted);
                    break;
                case "DOUBLE":

                    strFormatMask = "";
                    if (!(vxmlSchemaNode.attributes.getNamedItem("FORMATMASK") == null))
                    {
                        strFormatMask = vxmlSchemaNode.attributes.getNamedItem("FORMATMASK").text;
                    }
                    if (Strings.Len(strFormatMask) != 0)
                    {

                        vxmlOutElem.setAttribute(
                            vxmlSchemaNode.nodeName + "_RAW", 
                            vfld.Value);
                        vxmlOutElem.setAttribute(
                            vxmlSchemaNode.nodeName, 
                            Support.Format(vfld.Value, strFormatMask, (FirstDayOfWeek)1, (FirstWeekOfYear)1));
                    }
                    else
                    {

                        vxmlOutElem.setAttribute(
                            vxmlSchemaNode.nodeName, 
                            vfld.Value);
                    }
                    break;
                case "COMBO":
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, 
                        vfld.Value);
                    if (vblnDoComboLookUp == true)
                    {
                        if (!(vxmlSchemaNode.attributes.getNamedItem("COMBOGROUP") == null))
                        {
                            strComboGroup = vxmlSchemaNode.attributes.getNamedItem("COMBOGROUP").text;
                            strComboValue = comboAssistEx.GetComboText(strComboGroup, Convert.ToInt16(vfld.Value));
                            vxmlOutElem.setAttribute(
                                vxmlSchemaNode.nodeName + "_TEXT", strComboValue);
                        }
                    }
                    break;
                case "DATE":
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, adoDateToString(vfld.Value));
                    break;
                case "DATETIME":
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, adoDateTimeToString(vfld.Value));
                    break;
                default:
                    vxmlOutElem.setAttribute(
                        vxmlSchemaNode.nodeName, 
                        vfld.Value);
                    break; 
            }
        }

        public static int adoDeleteFromNode(MSXML2.IXMLDOMNode vxmlRequestNode, string vstrSchemaName, bool blnOnlySingleRecord) 
        {
            const string strFunctionName = "adoDeleteFromNode";
            int lngRecordsAffected = 0;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            // WARNING: On Error GOTO adoDeleteFromNodeExit is not supported
            try 
            {

                xmlSchemaNode = adoGetSchema(vstrSchemaName);
                lngRecordsAffected = adoDelete(vxmlRequestNode, xmlSchemaNode, blnOnlySingleRecord);
                // WARNING: adoDeleteFromNodeExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSchemaNode = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
            return lngRecordsAffected;
        }

        public static int adoDelete(MSXML2.IXMLDOMNode vxmlDataNode, MSXML2.IXMLDOMNode vxmlSchemaNode, bool blnOnlySingleRecord) 
        {
            const string strFunctionName = "adoDelete";
            string strDataSource = String.Empty;
            string strSQL = String.Empty;
            string strSQLWhere = String.Empty;
            MSXML2.IXMLDOMNode xmlSchemaNode = null;
            MSXML2.IXMLDOMNodeList xmlSchemaNodeList = null;
            MSXML2.IXMLDOMAttribute xmlDataAttrib = null;
            ADODB.Command cmd = null;
            ADODB.Parameter param = null;
            int lngRecordsAffected = 0;
            bool blnDbCmdOk = false;
            short intRetryMax = 0;
            short intRetryCount = 0;
            // WARNING: On Error GOTO adoDeleteExit is not supported
            try 
            {

                cmd = new ADODB.Command();
                // Get the table name from the schema
                strDataSource = xmlAssistEx.xmlGetAttributeText(vxmlSchemaNode, "DATASRCE", "");
                if (Strings.Len(strDataSource.Trim()) == 0)
                {
                    errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLMissingAttribute), "DATASRCE");
                }
                // If specified that only a single record may be deleted, ensure that
                // the whole primary key has been provided. Otherwise raise an error.
                if (blnOnlySingleRecord)
                {
                    xmlSchemaNodeList = vxmlSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']");
                    foreach(MSXML2.IXMLDOMNode __each11 in xmlSchemaNodeList)
                    {
                        xmlSchemaNode = __each11;
                        xmlDataAttrib = vxmlDataNode.attributes.getNamedItem(xmlSchemaNode.nodeName);
                        if (!(xmlDataAttrib == null))
                        {
                            if (Convert.ToString(xmlDataAttrib.value) == "")
                            {
                                errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLInvalidAttributeValue), xmlSchemaNode.nodeName);
                            }
                        }
                        else
                        {
                            errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeXMLMissingAttribute), xmlSchemaNode.nodeName);
                        }
                    }
                }
                // Get each condition in the WHERE clause
                foreach(MSXML2.IXMLDOMNode __each12 in vxmlSchemaNode.childNodes)
                {
                    xmlSchemaNode = __each12;
                    xmlDataAttrib = vxmlDataNode.attributes.getNamedItem(xmlSchemaNode.nodeName);
                    if (!(xmlDataAttrib == null))
                    {
                        if (Convert.ToString(xmlDataAttrib.value) != "")
                        {
                            if (Strings.Len(strSQLWhere) != 0)
                            {
                                strSQLWhere = strSQLWhere + " AND ";
                            }
                            strSQLWhere = strSQLWhere + xmlSchemaNode.nodeName + "=?";
                            param = getParam(xmlSchemaNode, xmlDataAttrib, false);
                            cmd.Parameters.Append(param);
                        }
                    }
                }
                // Run the SQL command
                if (Strings.Len(strSQLWhere) > 0)
                {
                    strSQL = "DELETE FROM " + strDataSource + " WHERE " + strSQLWhere;
                    System.Diagnostics.Debug.WriteLine("adoDelete strSQL: " + strSQL);
                    cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdText;
                    cmd.CommandText = strSQL;
                    // intRetryMax = adoGetDbRetries
                    // On Error Resume Next
                    // Do While blnDbCmdOk = False
                    lngRecordsAffected = executeSQLCommand(cmd);
                    // If Err.Number = 0 Then
                    // blnDbCmdOk = True
                    // Else
                    // If IsRetryError() = False Then
                    // On Error GoTo adoDeleteExit
                    // Else
                    // intRetryCount = intRetryCount + 1
                    // ' DEBUG
                    // App.LogEvent "adoDelete - retry [" & intRetryCount & "]: " & strSQL
                    // If (intRetryCount = intRetryMax) Then
                    // On Error GoTo adoDeleteExit
                    // End If
                    // End If
                    // End If
                    // Loop
                }
                System.Diagnostics.Debug.WriteLine("adoDelete records deleted: " + lngRecordsAffected);
                // WARNING: adoDeleteExit: is not supported 
            }
            catch(Exception exc)
            {
                cmd = null;
                param = null;
                xmlSchemaNode = null;
                xmlDataAttrib = null;
                param = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
            return lngRecordsAffected;
        }

        public static void adoPopulateChildKeys(MSXML2.IXMLDOMNode vxmlSrceNode, MSXML2.IXMLDOMNode vxmlDestNode) 
        {
            const string strFunctionName = "adoPopulateChildKeys";
            MSXML2.IXMLDOMNode xmlSrceSchemaNode = null;
            MSXML2.IXMLDOMNode xmlDestSchemaNode = null;
            MSXML2.IXMLDOMNodeList xmlPrimaryKeyNodeList = null;
            // WARNING: On Error GOTO adoPopulateChildKeysExit is not supported
            try 
            {
                xmlSrceSchemaNode = adoGetSchema(vxmlSrceNode.nodeName);
                if (xmlSrceSchemaNode == null)
                {
                    return ;
                }
                xmlDestSchemaNode = adoGetSchema(vxmlDestNode.nodeName);
                if (xmlDestSchemaNode == null)
                {
                    xmlSrceSchemaNode = null;
                    return ;
                }
                xmlPrimaryKeyNodeList = 
                    xmlDestSchemaNode.selectNodes("*[@KEYTYPE='PRIMARY']");
                foreach(MSXML2.IXMLDOMNode __each13 in xmlPrimaryKeyNodeList)
                {
                    xmlSrceSchemaNode = __each13;
                    xmlAssistEx.xmlCopyAttribute(vxmlSrceNode, vxmlDestNode, xmlSrceSchemaNode.nodeName);
                }
                // WARNING: adoPopulateChildKeysExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlSrceSchemaNode = null;
                xmlDestSchemaNode = null;
                xmlPrimaryKeyNodeList = null;
                errAssistEx.errCheckError(strFunctionName, String.Empty, null);
            }
        }

        public static void adoBuildDbConnectionString() 
        {
            object objWshShell = null;
            string strConnection = String.Empty;
            string strProvider = String.Empty;
            string strRegSection = String.Empty;
            string strRetries = String.Empty;
            string strUserId = String.Empty;
            string strPassword = String.Empty;
            int lngErrNo = 0;
            string strSource = String.Empty;
            string strDescription = String.Empty;
            // WARNING: On Error GOTO BuildDbConnectionStringVbErr is not supported
            objWshShell = new System.__ComObject();
            strRegSection = "HKLM\\SOFTWARE\\" + gstrAppName + "\\" + App.Title + "\\" + gstrREGISTRY_SECTION + "\\";
            // WARNING: On Error Resume Next is not supported

            // ISSUE: Method or data member not found: 'RegRead'
            strProvider = Convert.ToString(objWshShell.RegRead(strRegSection + gstrPROVIDER_KEY));
            // WARNING: On Error GOTO BuildDbConnectionStringVbErr is not supported

            if (Strings.Len(strProvider) == 0)
            {
                strRegSection = "HKLM\\SOFTWARE\\" + gstrAppName + "\\" + gstrREGISTRY_SECTION + "\\";
                // ISSUE: Method or data member not found: 'RegRead'
                strProvider = Convert.ToString(objWshShell.RegRead(strRegSection + gstrPROVIDER_KEY));
            }
            genumDbProvider = DBPROVIDER.omiga4DBPROVIDERUnknown;
            // If strProvider = "MSDAORA" Then
            // SDS LIVE00009659  22/01/2004 STARTS
            // Support for both MS and Oracle OLEDB Drivers
            if ((strProvider == "MSDAORA" || strProvider == "ORAOLEDB.ORACLE"))
            {
                switch (strProvider)
                {
                    case "MSDAORA":
                        genumDbProvider = DBPROVIDER.omiga4DBPROVIDERMSOracle;
                        break;
                    case "ORAOLEDB.ORACLE":
                        genumDbProvider = DBPROVIDER.omiga4DBPROVIDEROracle;
                        break; 
                }
                // genumDbProvider = omiga4DBPROVIDEROracle
                // ISSUE: Method or data member not found: 'RegRead'
                // ISSUE: Method or data member not found: 'RegRead'
                // ISSUE: Method or data member not found: 'RegRead'
                strConnection = 
                    "Provider=" + strProvider + ";Data Source=" + objWshShell.RegRead(strRegSection + gstrDATA_SOURCE_KEY) + ";" + 
                    "User ID=" + objWshShell.RegRead(strRegSection + gstrUID_KEY) + ";" + 
                    "Password=" + objWshShell.RegRead(strRegSection + gstrPASSWORD_KEY) + ";";
                // SDS LIVE00009659  22/01/2004 ENDS
            }
            else if( strProvider == "SQLOLEDB")
            {
                genumDbProvider = DBPROVIDER.omiga4DBPROVIDERSQLServer;
                // PSC 17/10/01 SYS2815 - Start
                strUserId = "";
                strPassword = "";
                // WARNING: On Error Resume Next is not supported

                // ISSUE: Method or data member not found: 'RegRead'
                strUserId = Convert.ToString(objWshShell.RegRead(strRegSection + gstrUID_KEY));
                lngErrNo = Information.Err().Number;
                strSource = Information.Err().Source;
                strDescription = Information.Err().Description;

                // WARNING: On Error GOTO BuildDbConnectionStringVbErr is not supported
                if (Information.Err().Number != glngENTRYNOTFOUND && Information.Err().Number != 0)
                {
                    Information.Err().Raise(lngErrNo, strSource, strDescription, null, null);
                }
                // WARNING: On Error Resume Next is not supported

                // ISSUE: Method or data member not found: 'RegRead'
                strPassword = Convert.ToString(objWshShell.RegRead(strRegSection + gstrPASSWORD_KEY));
                lngErrNo = Information.Err().Number;
                strSource = Information.Err().Source;
                strDescription = Information.Err().Description;

                // WARNING: On Error GOTO BuildDbConnectionStringVbErr is not supported
                if (Information.Err().Number != glngENTRYNOTFOUND && Information.Err().Number != 0)
                {
                    Information.Err().Raise(lngErrNo, strSource, strDescription, null, null);
                }
                // ISSUE: Method or data member not found: 'RegRead'
                // ISSUE: Method or data member not found: 'RegRead'
                strConnection = 
                    "Provider=SQLOLEDB;Server=" + objWshShell.RegRead(strRegSection + gstrSERVER_KEY) + ";" + 
                    "database=" + objWshShell.RegRead(strRegSection + gstrDATABASE_KEY) + ";";
                // If User Id is present use SQL Server Authentication else
                // use integrated security
                if (Strings.Len(strUserId) > 0)
                {
                    strConnection = strConnection + "UID=" + strUserId + ";" + 
                        "pwd=" + strPassword + ";";
                }
                else
                {
                    strConnection = strConnection + "Integrated Security=SSPI;Persist Security Info=False";
                }
                // PSC 17/10/01 SYS2815 - End
            }

            gstrDbConnectionString = strConnection;
            // ISSUE: Method or data member not found: 'RegRead'
            strRetries = Convert.ToString(objWshShell.RegRead(strRegSection + gstrRETRIES_KEY));
            if (Strings.Len(strRetries) > 0)
            {
                gintDbRetries = (short)(Convert.ToInt32(strRetries));
            }
            objWshShell = null;
            System.Diagnostics.Debug.WriteLine(strConnection);
            return ;
        BuildDbConnectionStringVbErr:
            objWshShell = null;
            Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
        }

        public static string adoGetDbConnectString() 
        {
            return gstrDbConnectionString;
        }

        public static short adoGetDbRetries() 
        {
            return gintDbRetries;
        }

        public static DBPROVIDER adoGetDbProvider() 
        {
            return genumDbProvider;
        }

        public static void adoLoadSchema() 
        {
            string strFileSpec = String.Empty;

            // pick up XML map from "...\Omiga 4\XML" directory
            // Only do the subsitution once to change DLL -> XML
            strFileSpec = System.Reflection.Assembly.GetExecutingAssembly().Location + "\\" + App.Title + ".xml";
            strFileSpec = Strings.Replace(strFileSpec, "DLL", "XML", 1, 1, ((Microsoft.VisualBasic.CompareMethod)(CompareMethod.Text)));
            gxmldocSchemas = new MSXML2.FreeThreadedDOMDocument40();
            gxmldocSchemas.async = false;
            gxmldocSchemas.setProperty("NewParser", true);
            gxmldocSchemas.validateOnParse = false;
            gxmldocSchemas.load(strFileSpec);
        }

        public static MSXML2.IXMLDOMNode adoGetSchema(string vstrSchemaName) 
        {
            string strPattern = String.Empty;

            strPattern = "//" + vstrSchemaName + "[@DATASRCE]";
            return gxmldocSchemas.selectSingleNode(strPattern);
        }

        private static int GetErrorNumber(string strErrDesc) 
        {
            const string cstrFunctionName = "GetErrorNumber";
            string strErr = String.Empty;
            // WARNING: On Error GOTO GetErrorNumberVbErr is not supported
            try 
            {
                switch (genumDbProvider)
                {
                    // SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
                    case DBPROVIDER.omiga4DBPROVIDEROracle:
                    case DBPROVIDER.omiga4DBPROVIDERMSOracle:
                        // oracle errors have format "ORA-nnnnn"
                        if (Strings.Len(strErrDesc) > 10)
                        {
                            strErr = strErrDesc.Substring(5 - 1, 5);
                        }
                        else
                        {
                            // "Cannot interpret database error"
                            errAssistEx.errThrowError(cstrFunctionName, (short)108, );
                        }
                        break;
                    case DBPROVIDER.omiga4DBPROVIDERSQLServer:
                        // SQL Server
                        errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeNotImplemented), );

                        break;
                    case DBPROVIDER.omiga4DBPROVIDERUnknown:
                        errAssistEx.errThrowError(cstrFunctionName, (short)(OMIGAERROR.oeNotImplemented), );
                        break; 
                }
                return strErr;
                // WARNING: GetErrorNumberVbErr: is not supported 
            }
            catch(Exception exc)
            {
                Information.Err().Raise(Information.Err().Number, cstrFunctionName, Information.Err().Description, null, null);
            }
        }

        public static int adoGetOmigaNumberForDatabaseError(string strErrDesc) 
        {
            const string cstrFunctioName = "adoGetOmigaNumberForDatabaseError";
            int lngErrNo = 0;
            int lngOmigaErrorNo = 0;
            // -----------------------------------------------------------------------------------
            // Description : Find the omiga equivalent number for a database error. This is used
            // to trap specific errors. Add to the list below, if you want trap a
            // a new error
            // Pass        : strErrDesc : Description of the Error Message (from database )
            // ------------------------------------------------------------------------------------
            // WARNING: On Error GOTO GetOmigaNumberForDatabaseErrorVbErr is not supported
            try 
            {
                lngErrNo = GetErrorNumber(strErrDesc);
                // 'SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
                if (genumDbProvider == DBPROVIDER.omiga4DBPROVIDEROracle || genumDbProvider == DBPROVIDER.omiga4DBPROVIDERMSOracle)
                {
                    switch (lngErrNo)
                    {
                        case 1:
                            lngOmigaErrorNo = OMIGAERROR.oeDuplicateKey;
                            break;
                        case 2292:
                            lngOmigaErrorNo = OMIGAERROR.oeChildRecordsFound;
                            break;
                        default:
                            lngOmigaErrorNo = OMIGAERROR.oeUnspecifiedError;
                            break; 
                    }
                }
                else
                {
                    errAssistEx.errThrowError(cstrFunctioName, (short)(OMIGAERROR.oeNotImplemented), "Error trapping for this database engine");
                }
                return lngOmigaErrorNo;
                // WARNING: GetOmigaNumberForDatabaseErrorVbErr: is not supported 
            }
            catch(Exception exc)
            {
                Information.Err().Raise(Information.Err().Number, cstrFunctioName, Information.Err().Description, null, null);
            }
        }

        public static bool adoGetValidRecordset(ref ADODB.Recordset vrstRecordSet, short nRecordSet) 
        {
            const string cstrFunctionName = "adoGetValidRecordset";
            bool bSuccess = false;
            // WARNING: On Error GOTO adoGetValidRecordsetErr is not supported
            try 
            {
                bSuccess = false;
                if (genumDbProvider == DBPROVIDER.omiga4DBPROVIDERSQLServer)
                {
                    while(!(vrstRecordSet == null))
                    {
                        if (!(vrstRecordSet.State == ADODB.ObjectStateEnum.adStateClosed))
                        {
                            if (!((vrstRecordSet.BOF && vrstRecordSet.EOF)))
                            {
                                if (nRecordSet > 0)
                                {
                                    nRecordSet = (short)(nRecordSet - 1);
                                }
                                bSuccess = true;
                                if (nRecordSet == 0)
                                {
                                    break;
                                }
                            }
                        }
                        vrstRecordSet = vrstRecordSet.NextRecordset();
                    } 
                }
                else if( !(vrstRecordSet == null) && !(vrstRecordSet.State == ADODB.ObjectStateEnum.adStateClosed) && !((vrstRecordSet.BOF && vrstRecordSet.EOF)))
                {
                    // Found an open recordset with records.
                    bSuccess = true;
                }
                return bSuccess;
                // WARNING: adoGetValidRecordsetErr: is not supported 
            }
            catch(Exception exc)
            {
                Information.Err().Raise(Information.Err().Number, cstrFunctionName, Information.Err().Description, null, null);
            }
        }

        private static bool IsComboLookUpRequired(MSXML2.IXMLDOMNode vxmlRequestNode) 
        {
            bool IsComboLookUpRequired = false;
            if (!(vxmlRequestNode.attributes.getNamedItem("_COMBOLOOKUP_") == null))
            {
                IsComboLookUpRequired = xmlAssistEx.xmlGetAttributeAsBoolean(vxmlRequestNode, "_COMBOLOOKUP_", "");
            }
            else
            {
                if (!(vxmlRequestNode.ownerDocument.firstChild == null))
                {
                    IsComboLookUpRequired = xmlAssistEx.xmlGetAttributeAsBoolean(vxmlRequestNode.ownerDocument.firstChild, "_COMBOLOOKUP_", "");
                }
                else
                {
                    IsComboLookUpRequired = false;
                }
            }
            return IsComboLookUpRequired;
        }

        private static string FormatGuid(string strSrcGuid, adoAssistEx.GUIDSTYLE eGuidStyle) 
        {
            string strFunctionName = String.Empty;
            string strTgtGuid = String.Empty;
            short intIndex1 = 0;
            short intIndex2 = 0;
            short intOffset = 0;
            // header ----------------------------------------------------------------------------------
            // description:  Converts a guid string from one format to another.
            // pass:
            // strSrcGuid  The guid string to be converted.
            // eGuidStyle  The style of the return string.
            // guidHyphen converts from "DA6DA163412311D4B5FA00105ABB1680" to
            // "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
            // guidBinary converts from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to
            // "DA6DA163412311D4B5FA00105ABB1680" format.
            // guidLiteral converts from either guidBinary or guidHyphen format to
            // a format that can be used as a literal input parameter for the
            // database. This is the default to maintain backwards compatibility with
            // Omiga 4 Phase 2.
            // return:       The converted guid string.
            // AS            05/03/01    First version
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO FormatGuidVbErr is not supported
            try 
            {
                strFunctionName = "FormatGuid";
                // By default return string unchanged, i.e., if style is guidHyphen but the source string
                // is currently not in binary format, then it will be returned unchanged (perhaps it is
                // already in hyphenated format?).
                strTgtGuid = strSrcGuid;
                if (eGuidStyle == GUIDSTYLE.guidHyphen)
                {
                    // Convert from "DA6DA163412311D4B5FA00105ABB1680" to "{63A16DDA-2341-D411-B5FA-00105ABB1680}"
                    System.Diagnostics.Debug.Assert(Strings.Len(strSrcGuid) == 32);
                    if (Strings.Len(strSrcGuid) == 32)
                    {
                        strTgtGuid = "{";
                        for(intIndex1 = (short)0;
                            intIndex1 <= 15.0;
                            intIndex1 = Convert.ToInt16(intIndex1 + 1))
                        {
                            intIndex2 = ConvertGuidIndex(intIndex1);
                            strTgtGuid = strTgtGuid + strSrcGuid.Substring((intIndex2 * 2) + 1 - 1, 2);
                            if (intIndex1 == 3 || intIndex1 == 5 || intIndex1 == 7 || intIndex1 == 9)
                            {
                                strTgtGuid = strTgtGuid + "-";
                            }
                        }
                        strTgtGuid = strTgtGuid + "}";
                    }
                    else
                    {
                        errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeInvalidParameter), "strSrcGuid length must be 32");
                    }
                }
                else if( eGuidStyle == GUIDSTYLE.guidBinary)
                {
                    // Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
                    System.Diagnostics.Debug.Assert(Strings.Len(strSrcGuid) == 38);
                    if (Strings.Len(strSrcGuid) == 38)
                    {
                        strTgtGuid = "";
                        intOffset = (short)2;
                        for(intIndex1 = (short)0;
                            intIndex1 <= 15.0;
                            intIndex1 = Convert.ToInt16(intIndex1 + 1))
                        {
                            intIndex2 = ConvertGuidIndex(intIndex1);
                            strTgtGuid = strTgtGuid + strSrcGuid.Substring((intIndex2 * 2) + intOffset - 1, 2);
                            if (intIndex1 == 3 || intIndex1 == 5 || intIndex1 == 7 || intIndex1 == 9)
                            {
                                intOffset = (short)(intOffset + 1);
                            }
                        }
                    }
                    else
                    {
                        errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeInvalidParameter), "strSrcGuid length must be 38");
                    }
                }
                else if( eGuidStyle == GUIDSTYLE.guidLiteral)
                {
                    // Convert guid into a format that can be used as a literal input parameter to the database.
                    // This assumes that the database type is raw for Oracle, or binary/varbinary for SQL Server.
                    if (Strings.Len(strSrcGuid) == 38)
                    {
                        // Guid is in hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}",
                        // so convert to binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
                        strSrcGuid = FormatGuid(strSrcGuid, GUIDSTYLE.guidBinary);
                    }
                    System.Diagnostics.Debug.Assert(Strings.Len(strSrcGuid) == 32);
                    if (Strings.Len(strSrcGuid) == 32)
                    {
                        switch (adoGetDbProvider())
                        {
                            case DBPROVIDER.omiga4DBPROVIDERSQLServer:
                                // e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                                strTgtGuid = "0x" + strSrcGuid;
                                break;
                            case DBPROVIDER.omiga4DBPROVIDEROracle:
                            case DBPROVIDER.omiga4DBPROVIDERMSOracle:
                                // SDS LIVE00009659  22/01/2004 Support for both MS and Oracle OLEDB Drivers
                                // e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                                strTgtGuid = "HEXTORAW('" + strSrcGuid + "')";
                                break; 
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.Assert(false);
                    errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeInvalidParameter), "strGuid length must be 32");
                }
                return strTgtGuid;
                // WARNING: FormatGuidVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static byte[] GuidStringToByteArray(string strGuid) 
        {
            string strFunctionName = String.Empty;
            byte[] rbytGuid = new byte[15 + 1];
            short intIndex = 0;
            // header ----------------------------------------------------------------------------------
            // description:      Converts a guid string into a byte array.
            // pass:
            // strGuid         The guid string to be converted.
            // Can be in either binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680",
            // or hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
            // return:           The byte array.
            // AS                05/03/01    First version
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO GuidStringToByteArrayVbErr is not supported
            try 
            {
                strFunctionName = "GuidStringToByteArray";
                if (Strings.Len(strGuid) == 38)
                {
                    // Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
                    strGuid = FormatGuid(strGuid, GUIDSTYLE.guidBinary);
                }
                if (Strings.Len(strGuid) == 32)
                {
                    // Convert from "DA6DA163412311D4B5FA00105ABB1680" to byte array.
                    for(intIndex = (short)0;
                        intIndex <= Information.UBound(rbytGuid, 1);
                        intIndex = Convert.ToInt16(intIndex + 1))
                    {
                        rbytGuid[intIndex] = Convert.ToByte("&H" + strGuid.Substring((intIndex * 2) + 1 - 1, 2));
                    }
                }
                else
                {
                    System.Diagnostics.Debug.Assert(false);
                    errAssistEx.errThrowError(strFunctionName, (short)(OMIGAERROR.oeInvalidParameter), "strGuid length must be 32");
                }
                return rbytGuid;
                // WARNING: GuidStringToByteArrayVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        private static short ConvertGuidIndex(short intIndex1) 
        {
            short intIndex2 = 0;
            // header ----------------------------------------------------------------------------------
            // description:  Helper function for ByteArrayToGuidString and FormatGuid.
            // pass:
            // return:
            // AS            31/01/01    First version
            // ------------------------------------------------------------------------------------------
            switch (intIndex1)
            {
                case 0:
                    intIndex2 = (short)3;
                    break;
                case 1:
                    intIndex2 = (short)2;
                    break;
                case 2:
                    intIndex2 = (short)1;
                    break;
                case 3:
                    intIndex2 = (short)0;
                    break;
                case 4:
                    intIndex2 = (short)5;
                    break;
                case 5:
                    intIndex2 = (short)4;
                    break;
                case 6:
                    intIndex2 = (short)7;
                    break;
                case 7:
                    intIndex2 = (short)6;
                    break;
                default:
                    intIndex2 = intIndex1;
                    break; 
            }
            return intIndex2;
        }


        private enum GUIDSTYLE {
            guidBinary = 0,
            guidHyphen = 1,
            guidLiteral = 2,
        }


    }

}
