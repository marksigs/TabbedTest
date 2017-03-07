using System;
using System.Transactions;
using System.Xml;
using System.Runtime.InteropServices;
using System.EnterpriseServices;
//using COMSVCSLib;
//using MSXML2;
//using omBase;
//using ADODB;
using omOrg;
using omAU;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace omLockManager
{
    //public class LockManagerBO : COMSVCSLib.ObjectControl
    public class LockManagerBO : ServicedComponent
    {

        // Workfile:      LockManagerBO.cls
        // Copyright:     Copyright © 2005 Marlborough Stirling
        // Description:   A business object that co-ordinates the calls to create and
        // release locks for the customer and application.
        // Note: XML is not used internally within this module, but only when interfacing
        // to other components that have an XML based interface.
        // Dependencies:
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog   Date        Description
        // AS     05/04/2005  Created
        // ------------------------------------------------------------------------------------------
        private COMSVCSLib.ObjectContext m_objContext = null;
        private ErrAssistException m_objErrAssistException = null;
        // TODO Replace with real error number.
        private const short gstrLockManagerErrorNumber = 8009;
        public string omRequest(string vstrXmlIn) 
        {
            string omRequest = String.Empty;
            const string cstrFunctionName = "omLockManager.LockManagerBO.omRequest";
            XmlDocument xmlInDoc = null;
            XmlDocument xmlOutDoc = null;
            XmlNode xmlRequestNode = null;
            XmlElement xmlResponseElem = null;
            XmlNode xmlResponseNode = null;
            bool blnSuccess = false;
            string strUserId = String.Empty;
            string strUnitID = String.Empty;
            string strMachineID = String.Empty;
            string strChannelID = String.Empty;
            string strApplicationNumber = String.Empty;
            string strCustomerNumber = String.Empty;
            string strOperation = String.Empty;
            bool blnEnableLocking = false;
            bool blnCreateNew = false;
            string strLocks = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Executes a application/customer lock request.
            // vstrRequest
            // xml Request data stream
            // Format:
            // <REQUEST
            // USERID=""
            // UNITID=""
            // MACHINEID=""
            // CHANNELID=""
            // LOCKING=""
            // APPLICATIONNUMBER=""
            // OPERATION="LOCK|UNLOCK">
            // </REQUEST>
            // WARNING: On Error GOTO omRequestExit is not supported
            try 
            {



                blnSuccess = false;

                xmlInDoc = new XmlDocument();
                xmlInDoc.setProperty("NewParser", true);
                //xmlInDoc.DocumentType.Name = "NewParser", true);
                xmlOutDoc = new XmlDocument();
                xmlOutDoc.setProperty("NewParser", true);

                xmlInDoc.async = false;
                xmlOutDoc.async = false;

                xmlResponseElem = xmlOutDoc.CreateElement("RESPONSE");
                xmlResponseElem.SetAttribute("TYPE", "SUCCESS");
                xmlResponseNode = xmlOutDoc.AppendChild(xmlResponseElem);

                System.Diagnostics.Debug.WriteLine(vstrXmlIn);
                xmlInDoc.LoadXml(vstrXmlIn);

                xmlRequestNode = xmlAssistEx.xmlGetMandatoryNode(xmlInDoc, "REQUEST", NO_MESSAGE_NUMBER);


                strUserId = xmlAssistEx.xmlGetMandatoryAttributeText(xmlRequestNode, "USERID", NO_MESSAGE_NUMBER);
                strUnitID = xmlAssistEx.xmlGetMandatoryAttributeText(xmlRequestNode, "UNITID", NO_MESSAGE_NUMBER);
                strMachineID = xmlAssistEx.xmlGetMandatoryAttributeText(xmlRequestNode, "MACHINEID", NO_MESSAGE_NUMBER);
                strChannelID = xmlAssistEx.xmlGetMandatoryAttributeText(xmlRequestNode, "CHANNELID", NO_MESSAGE_NUMBER);
                strApplicationNumber = xmlAssistEx.xmlGetAttributeText(xmlRequestNode, "APPLICATIONNUMBER", "");
                strCustomerNumber = xmlAssistEx.xmlGetAttributeText(xmlRequestNode, "CUSTOMERNUMBER", "");
                strOperation = xmlAssistEx.xmlGetMandatoryAttributeText(xmlRequestNode, "OPERATION", NO_MESSAGE_NUMBER);

                // For MAPS this is the global parameter "Locking"
                // For POS this is the global parameter "POSLocking"
                // The caller (GUI etc) decides which parameter to use, and passes it in as the "ENABLELOCKING" attribute.
                blnEnableLocking = xmlAssistEx.xmlGetAttributeAsBoolean(xmlRequestNode, "ENABLELOCKING", "1");

                // If True then there must not already be an APPLICATIONLOCK for this application.
                // If False then there must not already be an APPLICATIONLOCK for this application,
                // unless it is for this user.
                blnCreateNew = xmlAssistEx.xmlGetAttributeAsBoolean(xmlRequestNode, "CREATENEW", "1");

                if (ValidateWorkingHours(strUserId, strChannelID))
                {
                    if (blnEnableLocking)
                    {
                        if (strOperation == "LOCK")
                        {
                            if (strApplicationNumber != "")
                            {
                                blnSuccess = LockApplication(strApplicationNumber, strUserId, strUnitID, strMachineID, blnCreateNew, ref strLocks);
                            }
                            else if( strCustomerNumber != "")
                            {
                                blnSuccess = LockCustomer(strCustomerNumber, strUserId, strUnitID, strMachineID, ref strLocks);
                            }
                            else
                            {
                                //errAssistEx.errThrowError(cstrFunctionName, OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                               // m_objErrAssistException.t .errThrowError(cstrFunctionName, OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                                throw new ErrAssistException(Vertex.Fsd.Omiga.VisualBasicPort.OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                            }
                        }
                        else if( strOperation == "UNLOCK")
                        {
                            if (strApplicationNumber != "")
                            {
                                blnSuccess = UnlockApplication(strApplicationNumber);
                            }
                            else if( strCustomerNumber != "")
                            {
                                blnSuccess = UnlockCustomer(strCustomerNumber);
                            }
                            else
                            {
                                //errAssistEx.errThrowError(cstrFunctionName, OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                                throw new ErrAssistException(Vertex.Fsd.Omiga.VisualBasicPort.OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                            }
                        }
                        else
                        {
                           // errAssistEx.errThrowError(cstrFunctionName, OMIGAERROR.oeInvalidParameter, "OPERATION attribute in REQUEST tag must be either \"LOCK\" or \"UNLOCK\"");
                            throw new ErrAssistException(Vertex.Fsd.Omiga.VisualBasicPort.OMIGAERROR.oeInvalidParameter, "APPLICATIONNUMBER or CUSTOMERNUMBER attribute in REQUEST tag must be specified");
                        }
                    }
                    else
                    {
                        blnSuccess = true;
                    }
                }

                //omRequest = xmlOutDoc.xml;
                omRequest = xmlOutDoc.Value;

                // WARNING: omRequestExit: is not supported 
            }
            catch(Exception exc)
            {
                xmlRequestNode = null;
                xmlResponseElem = null;
                xmlResponseNode = null;
                xmlInDoc = null;
                xmlOutDoc = null;

                if (Information.Err().Number == 0)
                {
                    // No error.
                    if (!(m_objContext == null))
                    {
                        m_objContext.SetComplete();
                    }
                }
                else
                {
                    // Error.
                   // Information.Err().Source = cstrFunctionName + ", " + Information.Err().Source;

                    omRequest = CreateErrorResponse(strLocks);

                    if (!(m_objContext == null))
                    {
                        m_objContext.SetAbort();
                    }
                }

                if (strApplicationNumber != "")
                {
                    // Ignore any errors that occur when auditing.
                    // WARNING: On Error Resume Next is not supported
                    // Only audit if locking an application.
                    // The audit record must be created after any error handling
                    // (as calling CreateApplicationAccessAudit will reset the Err object).
                    CreateApplicationAccessAudit(blnSuccess, strUserId, strMachineID, strApplicationNumber);
                }

                System.Diagnostics.Debug.WriteLine(omRequest);

            }
            return omRequest;
        }

        private string CreateErrorResponse(string strLocks) 
        {
            XmlElement xmlResponseElem = null;
            XmlDocument  xmlLocksDoc = null;
            XmlElement xmlErrorElem = null;
            XmlElement xmlLocksElem = null;

            xmlResponseElem = (XmlElement)m_objErrAssistException.CreateErrorResponseNode(false);

            // Must be after calling CreateErrorResponseNode, as "On Error" clears the error object.
            // WARNING: On Error GOTO CreateErrorResponseExit is not supported
            try 
            {

                if (strLocks.Length > 0)
                {

                    xmlLocksDoc = new XmlDocument();
                    //xmlLocksDoc.setProperty("NewParser", true);
                    //xmlLocksDoc.async = false;
                    xmlLocksDoc.LoadXml(strLocks);

                    xmlErrorElem = (XmlElement)xmlResponseElem.SelectSingleNode("ERROR");
                    if (!(xmlErrorElem == null))
                    {
                        xmlLocksElem = (XmlElement)xmlLocksDoc.SelectSingleNode("LOCKS");
                        if (!(xmlLocksElem == null))
                        {
                            xmlErrorElem.AppendChild(xmlLocksElem.CloneNode(true));
                        }
                    }
                }

                // WARNING: CreateErrorResponseExit: is not supported 
            }
            catch(Exception exc)
            {
                if (!(xmlResponseElem == null))
                {
                }

                xmlLocksElem = null;
                xmlErrorElem = null;
                xmlLocksDoc = null;
                xmlResponseElem = null;
            }
            return xmlResponseElem.xml;
        }

        private bool LockApplication(string strApplicationNumber, string strUserId, string strUnitID, string strMachineID, bool blnCreateNew, ref string strLocks) 
        {
            const string cstrFunctionName = "omLockManager.LockManagerBO.LockApplication";
            omBase.ADOAssist objADOAssist = null;
            ADODB.Command cmd = null;
            ADODB.Stream strm = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Places an application lock and related customer locks into the database via a stored-
            // procedure.
            // Based on omApp.ApplicationManagerDO.cls:IApplicationManagerDO_LockCustomersForApplication().
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO LockApplicationExit is not supported
            try 
            {



                objADOAssist = Interaction.IIf(m_objContext == null, new omBase.ADOAssist(), m_objContext.CreateInstance(StdData.gstrBASE_COMPONENT + ".ADOAssist"));

                cmd = new ADODB.Command();
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdStoredProc;
                cmd.CommandText = "usp_LockManagerLockApplication";
                cmd.ActiveConnection = objADOAssist.GetConnStr();
                cmd.Parameters.Append(cmd.CreateParameter("@applicationNumber", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 12, strApplicationNumber));
                cmd.Parameters.Append(cmd.CreateParameter("@userID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 10, strUserId));
                cmd.Parameters.Append(cmd.CreateParameter("@unitID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 10, strUnitID));
                cmd.Parameters.Append(cmd.CreateParameter("@machineID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 30, strMachineID));
                cmd.Parameters.Append(cmd.CreateParameter("@createNew", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adBoolean, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 0, blnCreateNew));

                strm = new ADODB.Stream();
                strm.Open(null, (ADODB.ConnectModeEnum)(0), (ADODB.StreamOpenOptionsEnum)(-1), "", "");
                cmd.Properties["Output Stream"].Value = strm;
                cmd.Properties["XML Root"].Value = "LOCKS";

                cmd.Execute(null, null, ADODB.ExecuteOptionEnum.adExecuteStream);
                strLocks = strm.ReadText();
                if (Strings.InStr(1, strLocks, "<LOCKS></LOCKS>", CompareMethod.Binary) == 0)
                {
                    // Assume strLocks matches "<LOCKS><CUSTOMERLOCK>...</CUSTOMERLOCK></LOCKS>",
                    // i.e., existing customer locks have been returned.
                    errAssistEx.errThrowError(cstrFunctionName, gstrLockManagerErrorNumber, null);
                }
                else
                {
                    strLocks = "";
                }

                // WARNING: LockApplicationExit: is not supported 
            }
            catch(Exception exc)
            {
                if (!(strm == null))
                {
                    if (strm.State == ADODB.ObjectStateEnum.adStateOpen)
                    {
                        strm.Close();
                    }
                }

                strm = null;
                cmd = null;
                objADOAssist = null;

                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                }
                else
                {
                }
            }
            return true;
        }

        private bool UnlockApplication(string strApplicationNumber) 
        {
            const string cstrFunctionName = "omLockManager.LockManagerBO.UnlockApplication";
            omBase.ADOAssist objADOAssist = null;
            ADODB.Command cmd = null;
            // header ----------------------------------------------------------------------------------
            // description:
            // Removes an application lock and related customer locks via a stored-
            // procedure.
            // Based on omApp.ApplicationManagerDO.cls:IApplicationManagerDO_UnlockApplication().
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO UnlockApplicationExit is not supported
            try 
            {



                objADOAssist = Interaction.IIf(m_objContext == null, new omBase.ADOAssist(), m_objContext.CreateInstance(StdData.gstrBASE_COMPONENT + ".ADOAssist"));

                cmd = new ADODB.Command();
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdStoredProc;
                cmd.CommandText = "usp_LockManagerUnlockApplication";
                cmd.ActiveConnection = objADOAssist.GetConnStr();
                cmd.Parameters.Append(cmd.CreateParameter("@applicationNumber", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 12, strApplicationNumber));

                cmd.Execute(null, null, -1);

                // WARNING: UnlockApplicationExit: is not supported 
            }
            catch(Exception exc)
            {
                objADOAssist = null;
                cmd = null;

                //if (Information.Err().Number != 0)
                //{
                //    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                //}
                //else
                //{
                //}
            }
            return true;
        }

        private bool LockCustomer(string strCustomerNumber, string strUserId, string strUnitID, string strMachineID, ref string strLocks) 
        {
            const string cstrFunctionName = "omLockManager.LockManagerBO.LockCustomer";
            omBase.ADOAssist objADOAssist = null;
            ADODB.Command cmd = null;
            ADODB.Stream strm = null;
            // WARNING: On Error GOTO LockCustomerExit is not supported
            try 
            {



                objADOAssist = Interaction.IIf(m_objContext == null, new omBase.ADOAssist(), m_objContext.CreateInstance(StdData.gstrBASE_COMPONENT + ".ADOAssist"));

                cmd = new ADODB.Command();
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdStoredProc;
                cmd.CommandText = "usp_LockManagerLockCustomer";
                cmd.ActiveConnection = objADOAssist.GetConnStr();
                cmd.Parameters.Append(cmd.CreateParameter("@customerNumber", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 12, strCustomerNumber));
                cmd.Parameters.Append(cmd.CreateParameter("@userID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 10, strUserId));
                cmd.Parameters.Append(cmd.CreateParameter("@unitID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 10, strUnitID));
                cmd.Parameters.Append(cmd.CreateParameter("@machineID", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 30, strMachineID));

                strm = new ADODB.Stream();
                strm.Open(null, (ADODB.ConnectModeEnum)(0), (ADODB.StreamOpenOptionsEnum)(-1), "", "");
                cmd.Properties["Output Stream"].Value = strm;
                cmd.Properties["XML Root"].Value = "LOCKS";

                cmd.Execute(null, null, ADODB.ExecuteOptionEnum.adExecuteStream);
                strLocks = strm.ReadText();
                if (Strings.InStr(1, strLocks, "<LOCKS></LOCKS>", CompareMethod.Binary) == 0)
                {
                    // Assume strLocks matches "<LOCKS><CUSTOMERLOCK>...</CUSTOMERLOCK></LOCKS>",
                    // i.e., existing customer locks have been returned.
                    errAssistEx.errThrowError(cstrFunctionName, gstrLockManagerErrorNumber, null);
                }
                else
                {
                    strLocks = "";
                }

                // WARNING: LockCustomerExit: is not supported 
            }
            catch(Exception exc)
            {
                if (!(strm == null))
                {
                    if (strm.State == ADODB.ObjectStateEnum.adStateOpen)
                    {
                        strm.Close();
                    }
                }

                strm = null;
                cmd = null;
                objADOAssist = null;

                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                }
                else
                {
                }
            }
            return true;
        }

        private bool UnlockCustomer(string strCustomerNumber) 
        {
            const string cstrFunctionName = "omLockManager.UnlockManagerBO.LockCustomer";
            omBase.ADOAssist objADOAssist = null;
            ADODB.Command cmd = null;
            // WARNING: On Error GOTO UnlockCustomerExit is not supported
            try 
            {



                objADOAssist = Interaction.IIf(m_objContext == null, new omBase.ADOAssist(), m_objContext.CreateInstance(StdData.gstrBASE_COMPONENT + ".ADOAssist"));

                cmd = new ADODB.Command();
                cmd.CommandType = (ADODB.CommandTypeEnum)ADODB.CommandTypeEnum.adCmdStoredProc;
                cmd.CommandText = "usp_LockManagerUnlockCustomer";
                cmd.ActiveConnection = objADOAssist.GetConnStr();
                cmd.Parameters.Append(cmd.CreateParameter("@customerNumber", (ADODB.DataTypeEnum)ADODB.DataTypeEnum.adVarChar, (ADODB.ParameterDirectionEnum)ADODB.ParameterDirectionEnum.adParamInput, 12, strCustomerNumber));

                cmd.Execute(null, null, -1);

                // WARNING: UnlockCustomerExit: is not supported 
            }
            catch(Exception exc)
            {
                cmd = null;
                objADOAssist = null;

                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                }
                else
                {
                }
            }
            return true;
        }

        private bool ValidateWorkingHours(string strUserId, string strChannelID) 
        {
            const string cstrFunctionName = "omLockManager.LockManagerBO.ValidateWorkingHours";
            omOrg.OrganisationBO objOrganisationBO = null;
            // WARNING: On Error GOTO ValidateWorkingHoursExit is not supported
            try 
            {


                objOrganisationBO = Interaction.IIf(m_objContext == null, new omOrg.OrganisationBO(), m_objContext.CreateInstance(StdData.gstrORGANISATION_COMPONENT + ".OrganisationBO"));

                // This will raise an error if the user is outside working hours.
                objOrganisationBO.ValidateWorkingHours("<REQUEST><OMIGAUSER><USERID>" + strUserId + "</USERID><CHANNELID>" + strChannelID + "</CHANNELID></OMIGAUSER></REQUEST>");

                // WARNING: ValidateWorkingHoursExit: is not supported 
            }
            catch(Exception exc)
            {
                objOrganisationBO = null;

                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                }
                else
                {
                }
            }
            return true;
        }

        private bool CreateApplicationAccessAudit(bool blnSuccess, string strUserId, string strMachineID, string strApplicationNumber) 
        {
            const string cstrFunctionName = "omLockManager.LockManagerBO.CreateApplicationAccessAudit";
            omAU.IAuditTxBO objIAuditTxBO = null;
            omAU.IAuditDO objIAuditDO = null;
            XmlDocument xmlRequest = null;
            string strAuditRecordType = String.Empty;
            string strRequest = String.Empty;
            // WARNING: On Error GOTO CreateApplicationAccessAuditExit is not supported
            try 
            {


                // CORE113 AS 04/05/2005 Do not use NTxBO as this results in lock time outs on
                // the APPLICATIONACCESS table when calling objIAuditNTxBO.CreateAccessAudit.
                // Because we're not using NTxBO (requires new), but TxBO (requires), then the audit
                // records are created in the same transaction as the locks. Thus, if the transaction is
                // aborted (e.g., on error), no audit records will be inserted.

                objIAuditTxBO = Interaction.IIf(m_objContext == null, new omAU.AuditTxBO(), m_objContext.CreateInstance(StdData.gstrAUDIT_COMPONENT + ".AuditTxBO"));
                objIAuditDO = Interaction.IIf(m_objContext == null, new omAU.AuditDO(), m_objContext.CreateInstance(StdData.gstrAUDIT_COMPONENT + ".AuditDO"));

                strAuditRecordType = objIAuditDO.GetApplicationLockValueId();

                xmlRequest = new XmlDocument();
                xmlRequest.setProperty("NewParser", true);

                strRequest = 
                    "<REQUEST>" + 
                    "<ACCESSAUDIT>" + 
                    "<USERID>" + strUserId + "</USERID>" + 
                    "<MACHINEID>" + strMachineID + "</MACHINEID>" + 
                    "<AUDITRECORDTYPE>" + strAuditRecordType + "</AUDITRECORDTYPE>" + 
                    "<SUCCESSINDICATOR>" + Interaction.IIf(blnSuccess, "1", "0") + "</SUCCESSINDICATOR>" + 
                    "<APPLICATIONNUMBER>" + strApplicationNumber + "</APPLICATIONNUMBER>" + 
                    "</ACCESSAUDIT>" + 
                    "</REQUEST>";
                xmlRequest.LoadXml(strRequest);

                errAssistEx.errCheckXMLResponse(objIAuditTxBO.CreateAccessAudit(xmlRequest.DocumentElement).xml, true, null);
                ErrAssistException.CheckXMLResponse(objIAuditTxBO.CreateAccessAudit(xmlRequest.DocumentElement).xml, true, null);
                // Raise an error on unsuccessful response.
                // WARNING: CreateApplicationAccessAuditExit: is not supported 
            }
            catch(Exception exc)
            {
                objIAuditDO = null;
                objIAuditTxBO = null;
                xmlRequest = null;

                if (Information.Err().Number != 0)
                {
                    Information.Err().Raise(Information.Err().Number, cstrFunctionName + ", " + Information.Err().Source, Information.Err().Description, Information.Err().HelpFile, Information.Err().HelpContext);
                }
                else
                {
                }
            }
            return true;
        }

        private void ObjectControl_Activate() 
        {
            object GetObjectContext = null;
            m_objContext = GetObjectContext;
        }

        private bool ObjectControl_CanBePooled() 
        {
            return false;
        }

        private void ObjectControl_Deactivate() 
        {
            m_objContext = null;
        }


    }

}
