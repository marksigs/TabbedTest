using System;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.omAU
{
    [ComVisible(true)]
    [ProgId("omAU.AuditDO")]
    [Guid("88584CB5-EE42-4F9B-9C92-8D87D71ED12F")]
    [Transaction(TransactionOption.Supported)]
    public class AuditDO : ServicedComponent, IAuditDO
    {
		// Workfile:      AuditDO.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // Description:   Data objects class for omAU
        // 
        // Dependencies:  ADOAssist
        // Add any other dependent components
        // 
        // Issues:        Instancing:         MultiUse
        // MTSTransactionMode: UsesTransaction
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog  Date         Description
        // MCS    02/09/99    Created
        // RF     29/09/99    Added CreateAccessAudit and GetNumberOfFailedAttempts
        // RF     19/11/99    Change to search logic in GetNumberOfFailedAttempts
        // MH     05/05/00    SYS0571 Use of Now() without formatting
        // MC     16/05/00    SYS0210 - Synchronise the password change date/time with corresponding
        // access audit record
        // CL     18/10/00    Core00004 Modifications made to conform to coding templates
        // PSC    12/12/00    CORE00004 Change RaiseError to ThrowError
        // APS    27/02/01    SYS1986
        // MV     06/03/01    SYS2001: Commenting in CreateAccessaudit , as CreateAccessaudit is a sub there is no return Data
        // at the end of the proc
        // LD    11/06/01     SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
        // LD    19/06/01     SYS2386 All projects to use guidassist.bas rather than guidassist.cls
        // ------------------------------------------------------------------------------------------
        // BBG Specific History:
        // 
        // Prog   Date        Description
        // TK     30/11/2004  E2EM00002504 - Performance related fixes.
        // ------------------------------------------------------------------------------------------
        // =============================================
        // Variable Declaration Section
        // =============================================
        // =============================================
        // Constant Declaration Section
        // =============================================
        private const string cstrLogon = "L";
        private const string cstrPasswordChange = "C";
        private const string cstrApplicationAccess = "AL";
        private const string cstrApplicationRelease = "AR";

        // ------------------------------------------------------------------------------------------
        // BMIDS Specific History:
        // 
        // Prog   Date        AQR         Description
        // MV     12/08/2002  BMIDS00322  Core Ref AQR: SYS2341 ; Modified GetHighestSequenceNumberForAttempt
        // DB     10/03/2003  BM0383      Commented out code that increments AttemptNumber.
        // ------------------------------------------------------------------------------------------

		#region IAuditDO Members

		int IAuditDO.GetNumberOfFailedAttempts(XmlElement vxmlTableElement)
        {
            try
            {
                    string strCriteria = String.Empty;
                    string strRecTypeTagValue = String.Empty;
                    string strRecTypeFieldName = String.Empty;
                    string strDateFieldName = String.Empty;
                    string strTableName = String.Empty;
                    string strUserIdFieldName = String.Empty;
                    string strUserIdTagValue = String.Empty;
                    object dteLogon = null;
                    object dteChangePassword = null;
                    DateTime dteSearch;
                    bool blnLogonDateFound = false;
                    bool blnChangePasswordDateFound = false;
                    bool blnSearchDateFound = false;
                    // header ----------------------------------------------------------------------------------
                    // description:
                    // Get number of failed attempts for a given audit record type and user id.
                    // Search logic is as follows:
                    // DateL = datetime of most recent successful logon
                    // DateCP  = datetime of most recent change password
                    // if DateL is found
                    // if DateCP is found
                    // if DateL > DateCP
                    // searchdate = DateL
                    // else (DateCP > DateL)
                    // searchdate = DateCP
                    // End If
                    // else (DateCP not found)
                    // searchdate = DateL
                    // End If
                    // else (DateL not found)
                    // if DateCP is found
                    // searchdate = DateCP
                    // Else
                    // searchdate = Null
                    // End If
                    // End If
                    // if searchdate not null
                    // NumFailedAttempts = number of failures since searchdate
                    // Else
                    // NumFailedAttempts = number of failures ever
                    // End If
                    // pass:
                    // vstrXMLRequest
                    // Format:
                    // <REQUEST>
                    // <USERID></USERID>
                    // <AUDITRECORDTYPE></AUDITRECORDTYPE>
                    // </REQUEST>
                    // return:
                    // ------------------------------------------------------------------------------------------
                    
                    
                    Type typADOAssist = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ADOAssist", true);
                    
                    // ------------------------------------------------------------------------------------------
                    // validate parameters
                    // ------------------------------------------------------------------------------------------
                    strRecTypeFieldName = "AUDITRECORDTYPE";
                    strUserIdFieldName = "USERID";
                    strRecTypeTagValue = XmlAssist.GetMandatoryNodeText(
                        vxmlTableElement, ".//" + strRecTypeFieldName);
                    strUserIdTagValue = XmlAssist.GetMandatoryNodeText(
                        vxmlTableElement, ".//" + strUserIdFieldName);
                    if (((IAuditDO)this).IsLogon(strRecTypeTagValue) == false && ((IAuditDO)this).IsPasswordChange(strRecTypeTagValue) == false)
                    {
                        //"Expected Logon or ChangePassword Audit record type");
                        throw new ErrAssistException(OMIGAERROR.MTSNotFound);
                    }
                    blnChangePasswordDateFound = false;
                    blnLogonDateFound = false;
                    blnSearchDateFound = false;
                    strTableName = "ACCESSAUDIT";
                    strDateFieldName = "ACCESSDATETIME";
                    // ------------------------------------------------------------------------------------------
                    // Get datetime of most recent successful change password record
                    // ------------------------------------------------------------------------------------------
                    // Build criteria string of format:
                    // "date = (select max date from accessaudit where
                    // userid = x and rectype = y and successindicator = true)"
                    strCriteria =
                        strDateFieldName + " = (SELECT MAX (" + strDateFieldName + ")" +
                        " FROM " + strTableName +
                        " WHERE " +
                        strUserIdFieldName + " = " + SqlAssist.FormatString(strUserIdTagValue, false) +
                        " AND " + strRecTypeFieldName + " = " + ((IAuditDO)this).GetChangePasswordValueId() +
                        " AND SUCCESSINDICATOR = 1)";
                    
                    bool blnReturnValue = true;
                    object tmpdteChangePassword = dteChangePassword;
                    
                    AdoAssist.GetValueFromTable(strTableName, strCriteria,
                        strDateFieldName, out tmpdteChangePassword, out blnReturnValue);
                    
                    blnChangePasswordDateFound = true;
                    
                    if (((IAuditDO)this).IsPasswordChange(strRecTypeTagValue) == true)
                    {
                        // ------------------------------------------------------------------------------------------
                        // set the search date
                        // ------------------------------------------------------------------------------------------
                        if (blnChangePasswordDateFound == true)
                        {
                            dteSearch = Convert.ToDateTime(dteChangePassword);
                            blnSearchDateFound = true;
                        }
                    }
                    else
                    {
                        // ------------------------------------------------------------------------------------------
                        // Get datetime of most recent successful logon
                        // ------------------------------------------------------------------------------------------
                        strCriteria =
                            strDateFieldName + " = (SELECT MAX (" + strDateFieldName + ")" +
                            " FROM " + strTableName +
                            " WHERE " +
                            strUserIdFieldName + " = " + SqlAssist.FormatString(strUserIdTagValue, false) +
                            " AND " + strRecTypeFieldName + " = " + strRecTypeTagValue +
                            " AND SUCCESSINDICATOR = 1)";
                        
                        bool blnFlag = true;
                        object tmpdteLogon = dteLogon;
                        AdoAssist.GetValueFromTable(strTableName, strCriteria,
                            strDateFieldName, out tmpdteLogon, out blnFlag);
                        blnChangePasswordDateFound = true;
                        
                        // ------------------------------------------------------------------------------------------
                        // set the search date
                        // ------------------------------------------------------------------------------------------
                        if (blnLogonDateFound == true)
                        {
                            if (blnChangePasswordDateFound == true)
                            {
                                if (Convert.ToDateTime(dteLogon) > Convert.ToDateTime(dteChangePassword))
                                {
                                    dteSearch = Convert.ToDateTime(dteLogon);
                                    blnSearchDateFound = true;
                                }
                                else
                                {
                                    dteSearch = Convert.ToDateTime(dteChangePassword);
                                    blnSearchDateFound = true;
                                }
                            }
                            else
                            {
                                dteSearch = Convert.ToDateTime(dteLogon);
                                blnSearchDateFound = true;
                            }
                        }
                        else
                        {
                            if (blnChangePasswordDateFound == true)
                            {
                                dteSearch = Convert.ToDateTime(dteChangePassword);
                                blnSearchDateFound = true;
                            }
                        }
                    }
                    // ------------------------------------------------------------------------------------------
                    // set up the criteria for the count of failed attempts
                    // ------------------------------------------------------------------------------------------
                    if (blnSearchDateFound == true)
                    {
                        // set criteria of format:
                        // "userid = x and rectype = y and SUCCESSINDICATOR = 0 and date > searchdate"
                        dteSearch = Convert.ToDateTime(dteChangePassword);
                        strCriteria =
                            strUserIdFieldName + " = " + SqlAssist.FormatString(strUserIdTagValue, false) +
                            " and " + strRecTypeFieldName + " = " + strRecTypeTagValue +
                            " and SUCCESSINDICATOR = 0" +
                            " and " + strDateFieldName + " > " +
                            SqlAssist.FormatDate(dteSearch, DATETIMEFORMAT.dtfDateTime);
                    }
                    else
                    {
                        // set criteria of format:
                        // "userid = x and rectype = y and SUCCESSINDICATOR = 0"
                        strCriteria =
                            strUserIdFieldName + " = " + SqlAssist.FormatString(strUserIdTagValue, false) +
                            " and " + strRecTypeFieldName + " = " + strRecTypeTagValue +
                            " and SUCCESSINDICATOR = 0";
                    }
                    // ------------------------------------------------------------------------------------------
                    // do the count of failed attempts
                    // ------------------------------------------------------------------------------------------
                   ContextUtility.SetComplete();
                   return AdoAssist.GetNumberOfRecords(strTableName, strCriteria);
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                ContextUtility.SetAbort();
                exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                return 0;
            }
        }

        private int GetHighestSequenceNumberForAttempt(string vstrUserId, string vstrInAuditRecordType) 
        {
            int GetHighestSequenceNumberForAttempt = 0;
            
            SqlCommand cmd = null;
            SqlConnection adoConnection = null;
            
            string strSQL = String.Empty;
            string sSQLNoLock = String.Empty;
            short intRetries = 0;
            short intMaxAttempts = 0;
            short intAttempt = 0;
            bool blnOpenedOk = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // get the currently highest sequence number from table ACCESSAUDIT for this value
            // of vstrUserId and vstrInAuditRecordType
            // pass:
            // vstrUserId
            // vstrInAuditRecordType
            // ------------------------------------------------------------------------------------------
            try 
            {
                using (adoConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
                {

                    if (!(vstrUserId.Length > 0))
                    {
                        throw new ErrAssistException(OMIGAERROR.MTSNotFound, "Invalid USERID");
                    }
                    if (!(vstrInAuditRecordType.Length > 0))
                    {
                        throw new ErrAssistException(OMIGAERROR.MTSNotFound, "Invalid AUDITRECORDTYPE");
                    }

                    // SYS2341 - The following SQL places an exclusive locks which can cause deadlocks during heavy
                    // periods of logging on. It does mean, however, that if two or more people login with the same
                    // userid at the same time, the ATTEMPTNUMBER returned will be the same, rather than increment.
                    if (AdoAssist.GetDbProvider() == DBPROVIDER.omiga4DBPROVIDERSQLServer)
                    {
                        sSQLNoLock = " WITH (NOLOCK)";
                    }

                    #region ADO.net Code

                    strSQL = "SELECT MAX(ATTEMPTNUMBER)" +
                            " FROM ACCESSAUDIT" + sSQLNoLock +
                            " WHERE USERID = @userId AND AUDITRECORDTYPE = @auditRecordType";
                    cmd.CommandText = strSQL;
                    cmd.Connection = adoConnection;

                    // 2. define parameters used in command object
                    SqlParameter paramUserID = new SqlParameter();
                    paramUserID.ParameterName = "userID";
                    paramUserID.Value = vstrUserId;

                    // 3. add new parameter to command object
                    cmd.Parameters.Add(paramUserID);

                    // 2. define parameters used in command object
                    SqlParameter paramAuditRecordType = new SqlParameter();
                    paramAuditRecordType.ParameterName = "auditRecordType";
                    paramAuditRecordType.Value = vstrInAuditRecordType;

                    cmd.Parameters.Add(paramAuditRecordType);

                    intRetries = (short)AdoAssist.GetDbRetries();
                    intMaxAttempts = (short)(1 + intRetries);
                    blnOpenedOk = false;
                    intAttempt = (short)1;
                    while ((blnOpenedOk == false) && (intAttempt <= intMaxAttempts))
                    {
                        adoConnection.Open();
                        if (adoConnection.State == ConnectionState.Open)
                        {
                            blnOpenedOk = true;
                        }
                        intAttempt = (short)(intAttempt + 1);
                    }
                    if (blnOpenedOk == false)
                    {
                        throw new ErrAssistException(OMIGAERROR.UnableToConnect);
                    }

                    using (SqlDataReader sqlreader = cmd.ExecuteReader())
                    {

                        while (sqlreader.Read())
                        {
                            if (Convert.IsDBNull(sqlreader[0].ToString()))
                            {
                                // sequence number not yet set on this field
                                GetHighestSequenceNumberForAttempt = 0;
                            }
                            else
                            {
                                GetHighestSequenceNumberForAttempt = Convert.ToInt32(sqlreader.GetInt32(0));
                            }
                        }
                    }
                }               
                ContextUtility.SetComplete();
                
                return GetHighestSequenceNumberForAttempt;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                
                if (adoConnection.State == ConnectionState.Open  )
                {
                    adoConnection.Close();
                }

                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                else
                {
                    ContextUtility.SetComplete();
                }
            }
            return GetHighestSequenceNumberForAttempt;
            
            #endregion

        }
    
		bool IAuditDO.IsPasswordChange(string vstrAuditRecType) 
        {
            bool blnIsPasswordChange = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // HasValue if access audit type is a password change
            // pass:
            // ------------------------------------------------------------------------------------------

            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					if (comboDO.IsItemInValidation("AccessAuditType", vstrAuditRecType, cstrPasswordChange))
					{
						blnIsPasswordChange = true;
					}
					else
					{
						blnIsPasswordChange = false;
					}
				}
                ContextUtility.SetComplete();
                
                return blnIsPasswordChange;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                if (exception != null)
                {
                    if (exception.IsSystemError())
                    {
                        ContextUtility.SetAbort();
                        exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);        
                    }
                    else
                    {
                        ContextUtility.SetComplete();
                    }
                }
                return blnIsPasswordChange;
            }
        }

        bool IAuditDO.IsApplicationAccess(string vstrAuditRecType) 
        {
            bool blnIsApplicationAccess = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // HasValue if access audit type is an application access
            // pass:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					if (comboDO.IsItemInValidation("AccessAuditType", vstrAuditRecType, cstrApplicationAccess))
					{
						blnIsApplicationAccess = true;
					}
					else
					{
						blnIsApplicationAccess = false;
					}
				}
                ContextUtility.SetComplete();
                
                return blnIsApplicationAccess;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                if (exception  != null)
                {
                    if (exception.IsSystemError())
                    {
                        ContextUtility.SetAbort();
                        exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                    }
                    else
                    {
                        ContextUtility.SetComplete();
                    }
                }

                return blnIsApplicationAccess;
            }
        }

        bool IAuditDO.IsApplicationRelease(string vstrAuditRecType) 
        {
            bool blnIsApplicationRelease = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // HasValue if access audit type is an application release
            // pass:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					if (comboDO.IsItemInValidation("AccessAuditType", vstrAuditRecType, cstrApplicationRelease))
					{
						blnIsApplicationRelease = true;
					}
					else
					{
						blnIsApplicationRelease = false;
					}
				}
                ContextUtility.SetComplete();
                return blnIsApplicationRelease;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                if (exception  != null)
                {
                    if (exception.IsSystemError())
                    {
                        ContextUtility.SetAbort();
                        exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                    }
                    else
                    {
                        ContextUtility.SetComplete();
                    }
                }                
                return blnIsApplicationRelease;
            }
        }

        bool IAuditDO.IsLogon(string vstrAuditRecType) 
        {
            bool blnIsLogon = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // HasValue if access audit type is a system logon or logoff
            // pass:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					if (comboDO.IsItemInValidation("AccessAuditType", vstrAuditRecType, cstrLogon))
					{
						blnIsLogon = true;
					}
					else
					{
						blnIsLogon = false;
					}
				}
                ContextUtility.SetComplete();
                return blnIsLogon;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                if (exception  != null)
                {
                    if (exception.IsSystemError())
                    {
                        ContextUtility.SetAbort();
                        exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                    }
                    else
                    {
                        ContextUtility.SetComplete();
                    }
                }
                return blnIsLogon;
            }
        }

        string IAuditDO.GetApplicationLockValueId() 
        {
            XmlDocument xmlResponseDoc = null;
            string strResponse = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // pass:
            // return:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					strResponse = comboDO.GetComboValueId("AccessAuditType", cstrApplicationAccess);
				}
                xmlResponseDoc = XmlAssist.Load(strResponse);
                ContextUtility.SetComplete();
                
                bool blnRef = false;
                return XmlAssist.GetTagValue(xmlResponseDoc.DocumentElement, "VALUEID", out blnRef, true);
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                if (exception != null)
                {
                    if (exception.IsSystemError())
                    {
                        ContextUtility.SetAbort();
                        exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                    }
                    else
                    {
                        ContextUtility.SetComplete();
                    }
                    return null;
                }
                return null;
            }
        }

        string IAuditDO.GetApplicationReleaseValueId() 
        {
            XmlDocument xmlResponseDoc = null;
            string strResponse = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // pass:
            // return:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					strResponse = comboDO.GetComboValueId("AccessAuditType", cstrApplicationRelease);
				}
                xmlResponseDoc = XmlAssist.Load(strResponse);
                ContextUtility.SetComplete();

                bool blnRef = false;
                return XmlAssist.GetTagValue(xmlResponseDoc.DocumentElement, "VALUEID", out blnRef, true);
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                
                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                else
                {
                    ContextUtility.SetComplete();
                }
                
                return null;
            }
        }

        string IAuditDO.GetChangePasswordValueId() 
        {
            XmlDocument xmlResponseDoc = null;
            string strResponse = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Get combo value id for a change password audit record type
            // pass:
            // return:
            // ------------------------------------------------------------------------------------------

            try 
            {
               //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
               //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
               //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					strResponse = comboDO.GetComboValueId("AccessAuditType", cstrPasswordChange);
				}
                xmlResponseDoc = XmlAssist.Load(strResponse);
                ContextUtility.SetComplete();

                bool blnRef = false;
                return XmlAssist.GetTagValue(xmlResponseDoc.DocumentElement, "VALUEID", out blnRef, true);
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                
                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                else
                {
                    ContextUtility.SetComplete();
                }
                
                return null;
            }
        }

        string IAuditDO.GetLogonValueId() 
        {
            XmlDocument xmlResponseDoc = null;
            string strResponse = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // Get combo value id for a logon audit record type
            // pass:
            // return:
            // ------------------------------------------------------------------------------------------
            try 
            {
                //[DS]- Removed reference to m_objContext. TBD: Instantiation of ComboDO()
                //Type typComboDO = Type.GetTypeFromProgID(StdData.gstrBASE_COMPONENT + ".ComboDO", true);
                //objComboDO = (ComboDO)Activator.CreateInstance(typComboDO);

				using (omBase.ComboDO comboDO = new omBase.ComboDO())
				{
					strResponse = comboDO.GetComboValueId("AccessAuditType", cstrLogon);
				}
                xmlResponseDoc = XmlAssist.Load(strResponse);
                ContextUtility.SetComplete();
                
                bool blnRef = false; ;
                return XmlAssist.GetTagValue(xmlResponseDoc.DocumentElement, "VALUEID", out blnRef, true);
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                
                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                else
                {
                    ContextUtility.SetComplete();
                }
                
                return null;
            }
        }

        void IAuditDO.AddDerivedData(XmlNode vxmlData) 
        {
            // header ----------------------------------------------------------------------------------
            // description:
            // XML elements must be created for any derived values as specified.
            // Add any derived values to XML. E.g. data type 'double' fields will
            // need to be formatted as strings to required precision & rounding.
            // pass:
            // vxmlData          base XML node
            // as:
            // <tablename>
            // <element1>element1 value</element1>
            // <elementn>elementn value</elementn>
            // return:                n/a
            // ------------------------------------------------------------------------------------------

            try 
            {
                ContextUtility.SetComplete();
                return ;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);
                
                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                }
                else
                {
                    ContextUtility.SetComplete();
                }
            }
        }

        void IAuditDO.CreateAccessAudit(XmlElement vxmlTableElement) 
        {
            XmlDocument xmlIn = new XmlDocument();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode xmlTableNode = null;
            XmlElement xmlElem = null;

            string strCreationDate = String.Empty;
            string strUserId = String.Empty;
            string strAuditRecType = String.Empty;
            IomAUClassDef objIomAUClassDef;
            XmlDocument xmlClassDefDoc = null;
            string strONBEHALFOFUSERID = String.Empty;
            string strTableName = String.Empty;
            string strAppNo = String.Empty;
            // header ----------------------------------------------------------------------------------
            // description:
            // pass:
            // vxmlTableElement  xml Request data stream containing data to be persisted.
            // Format:
            // <ACCESSAUDIT>
            // <USERID></USERID>
            // <AUDITRECORDTYPE></AUDITRECORDTYPE>
            // <MACHINEID></MACHINEID>
            // <SUCCESSINDICATOR></SUCCESSINDICATOR>
            // <ONBEHALFOFUSERID></ONBEHALFOFUSERID>
            // <APPLICATIONNUMBER></APPLICATIONNUMBER>
            // <PASSWORDCREATIONDATE>optional</PASSWORDCREATIONDATE>
            // </ACCESSAUDIT>
            // ------------------------------------------------------------------------------------------
            try 
            {
                // initialise some generally used values
                strUserId = XmlAssist.GetMandatoryNodeText(vxmlTableElement, ".//USERID");
                strAuditRecType = XmlAssist.GetMandatoryNodeText(vxmlTableElement, ".//AUDITRECORDTYPE");
                // ------------------------------------------------------------------------------------------
                // create ACCESSAUDIT record
                // ------------------------------------------------------------------------------------------
                xmlDoc = new XmlDocument();
                xmlDoc.AppendChild(vxmlTableElement.CloneNode(true));
                xmlTableNode = XmlAssist.GetMandatoryNode(xmlDoc, ".//ACCESSAUDIT");

                xmlElem = xmlDoc.CreateElement("ACCESSAUDITGUID");
                xmlElem.InnerText = GuidAssist.CreateGUID();
                xmlTableNode.AppendChild(xmlElem);
                strCreationDate = XmlAssist.GetNodeText(xmlTableNode,".//PASSWORDCREATIONDATE");
                if (strCreationDate.Trim().Length == 0)
                {
                    strCreationDate = DateTime.Now.ToString("DD/MM/YYYY HH:MM:SS");
                }
                xmlElem = xmlDoc.CreateElement("ACCESSDATETIME");
                xmlElem.InnerText = strCreationDate;
                xmlTableNode.AppendChild(xmlElem);
                objIomAUClassDef = new omAUClassDef();

                xmlClassDefDoc = objIomAUClassDef.LoadAccessAuditData();
                DOAssist.CreateEx(xmlDoc.DocumentElement, xmlClassDefDoc);
                xmlClassDefDoc = null;
                // ------------------------------------------------------------------------------------------
                // create CHANGEPASSWORD record if password has been changed by someone other
                // than the user (e.g. system administrator)
                // ------------------------------------------------------------------------------------------
                // Find the xmlnode ONBEHALFOFUSERID in the current path
              //  strONBEHALFOFUSERID = 
                strONBEHALFOFUSERID = XmlAssist.GetNodeText(vxmlTableElement, ".//ONBEHALFOFUSERID");
              
                // If a value has been found
                if (strONBEHALFOFUSERID.Length > 0)
                {

                    if (strONBEHALFOFUSERID != strUserId)
                    {

                        if (((IAuditDO)this).IsPasswordChange(strAuditRecType))
                        {

                           // create CHANGEPASSWORD record
                           strTableName = "CHANGEPASSWORD";
                           XmlNode xmlDocNode;
                           xmlDocNode = (XmlNode )xmlDoc.DocumentElement;
                           XmlAssist.ChangeNodeName(xmlDocNode, "ACCESSAUDIT", strTableName);
                           xmlTableNode = XmlAssist.GetMandatoryNode(xmlDoc, ".//" + strTableName);
                           xmlClassDefDoc = objIomAUClassDef.LoadChangePasswordData();
                           DOAssist.CreateEx(xmlDoc.DocumentElement, xmlClassDefDoc);
                           xmlClassDefDoc = null;
                        }
                    }
                }
                // ------------------------------------------------------------------------------------------
                // create APPLICATIONACCESS for application access/release
                // ------------------------------------------------------------------------------------------
                if (((IAuditDO)this).IsApplicationAccess(strAuditRecType) || ((IAuditDO)this).IsApplicationRelease(strAuditRecType))
                {
                    // create APPLICATIONACCESS record
                    // SR 23/03/00 - Search in the current node only (not recursively)
                    strAppNo = XmlAssist.GetMandatoryNodeText(vxmlTableElement, "APPLICATIONNUMBER");

                    strTableName = "APPLICATIONACCESS";
                    XmlNode xmlDocNode;
                    xmlDocNode = (XmlNode)xmlDoc.DocumentElement;
                    XmlAssist.ChangeNodeName(xmlDocNode, "ACCESSAUDIT", strTableName);
                    xmlTableNode = XmlAssist.GetMandatoryNode(xmlDoc, ".//" + strTableName);


                    xmlElem = xmlDoc.CreateElement("APPLICATIONNUMBER");
                    xmlElem.InnerText  = strAppNo;
                    xmlTableNode.AppendChild(xmlElem);

                    xmlClassDefDoc = objIomAUClassDef.LoadApplicationAccessData();
                    DOAssist.CreateEx(xmlDoc.DocumentElement, xmlClassDefDoc);
                    xmlClassDefDoc = null;
                }
                
                ContextUtility.SetComplete();
                return ;
            }
            catch (Exception ex)
            {
                ErrAssistException exception = new ErrAssistException(ex);        
                if (exception.IsSystemError())
                {
                    ContextUtility.SetAbort();
                    exception.LogError(exception.Message, System.Diagnostics.EventLogEntryType.Error);
                }
                else
                {
                    ContextUtility.SetComplete();
                }
            }
		}

		#endregion
	}
}
