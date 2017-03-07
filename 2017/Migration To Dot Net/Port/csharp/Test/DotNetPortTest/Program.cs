using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Transactions;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort;
using Vertex.Fsd.Omiga.VisualBasicPort.Tests;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{

	public class MyClass : IComparable
	{
		private int _time = 0;
		private int _index;

		public int Index
		{
			get { return _index; }
			set { _index = value; }
		}

		public MyClass(int index)
		{
			_index = index;
		}

		public int CompareTo(object obj)
		{
			return this._time.CompareTo(((MyClass)obj)._time);
		}
	}

	public class MyComparer : IComparer<MyClass>
	{
		public int Compare(MyClass x, MyClass y)
		{
			return x.Index.CompareTo(y.Index);
		}
	}


	class Program
	{
		static void Foo(string x)
		{
			x = "world";
		}

		public static string ToString(double value)
		{
			return value.ToString("");
		}

		public static string ToString(float value)
		{
			return value.ToString();
		}

		public static string ToString(decimal value)
		{
			return value.ToString();
		}

		static void Main(string[] args)
		{
			//MessageTest messageTest = new MessageTest();
			//messageTest.TestMessageBOGetData();
			//messageTest.TestMessageDOGetMessageDetails();

			//SystemDatesTest systemDatesTest = new SystemDatesTest();
			//systemDatesTest.TestSystemDatesDOCheckNonWorkingOccurence();
			//systemDatesTest.TestSystemDatesBOCheckNonWorkingOccurence();
			//systemDatesTest.TestSystemDatesBOFindWorkingDay();


			GlobalParameterTest globalParameterTest = new GlobalParameterTest();
			//try
			//{
				//globalParameterTest.TestGlobalParameterDOCreate();
				//globalParameterTest.TestGlobalParameterDOUpdate();
				//globalParameterTest.TestGlobalParameterDODelete();
				//globalParameterTest.TestGlobalParameterDOGetCurrentParameter();
				//globalParameterTest.TestGlobalParameterDOGetCurrentParameterByType();
				//globalParameterTest.TestGlobalParameterDOGetCurrentParameterListEx();
				globalParameterTest.TestGlobalParameterBOGetCurrentParameter();
				//globalParameterTest.TestGlobalParameterBOIsTaskManager();
				//globalParameterTest.TestGlobalParameterBOIsMultipleLender();
				//globalParameterTest.TestGlobalParameterBOFindCurrentParameterList();
				//globalParameterTest.TestGlobalParameterBOGetCurrentParameterListEx();
				//globalParameterTest.TestGlobalBandedParameterDOGetCurrentParameter();
				//globalParameterTest.TestGlobalBandedParameterBOGetCurrentParameter();
			//}
			//catch (Exception exception)
			//{
			//	Console.WriteLine(exception.ToString());
			//}
			//CurrencyTest test = new CurrencyTest();
			//test.TestCurrencyBOFindList();
			//test.TestCurrencyDOFindList();

			//LogAssistTest test = new LogAssistTest();
			//test.Test();

			//DbXmlAssistTest test = new DbXmlAssistTest();
			//test.Test();

			//TraceAssistTest test = new TraceAssistTest();
			//test.TestTrace();

			//VersionAssistTest test = new VersionAssistTest();
			//test.TestGetVersionList();

			//TemplateHandlerBOTest test = new TemplateHandlerBOTest();
			//test.TestGetTemplateString();
			//test.TestFindAvailableTemplatesString();
			//test.TestFindAvailableTemplates();
			//test.TestGetTemplate();

			//TemplateHandlerDOTest test = new TemplateHandlerDOTest();
			//test.TestFindAvailableTemplates();

			//OmigaMessageQueueTest test = new OmigaMessageQueueTest();
			//test.TestSendToQueue();
			//test.TestAsyncSend();

			//DOAssistTest test = new DOAssistTest();
			//test.TestFindListMultipleEx();
			//test.TestFindListEx();
			//test.TestGenerateSequenceNumber();
			//test.TestGenerateSequenceNumberEx();
            //test.TestGetNextSequenceNumber();
            //test.TestUpdate();
            //test.TestCreate();
			//test.TestFindList();
			//test.TestGetData();
			//test.TestGetDataEx();
			//test.TestGetComponentData();
			//test.TestDelete();

			//ResumeNextItemTest test = new ResumeNextItemTest();
			//test.TestResumeNextItem();

			//GeneralAssistTest test = new GeneralAssistTest();
			//test.TestContainsSpecialChars();
			//test.TestIsAlpha();
			//test.TestHasDuplicatedChars();
			//test.TestConvertToMixedCase();

			//ComboAssistTest test = new ComboAssistTest();
			//test.TestBaseComboBOGetComboList();
			//test.TestBaseComboBOGetComboValue();
			//test.TestBaseComboDOGetQuickQuoteValuationTypeValueId();
			//test.TestBaseComboDOGetQuickQuoteLocationValueId();
			//test.TestBaseComboDOGetFirstComboValidation();
			//test.TestBaseComboDOGetNewLoanValue();
			//test.TestBaseComboDOGetComboValueIdFromValueName();
			//test.TestBaseComboDOGetFirstComboValueIdFromValueName();
			//test.TestBaseComboDOGetComboValueId();
			//test.TestBaseComboDOGetFirstComboValueId();
			//test.TestBaseComboDOIsItemInValidation();
			//test.TestBaseComboDOGetComboText();
			//test.TestBaseComboDOGetComboValue();
			//test.TestBaseComboDOGetComboList();
			//test.TestGetComboGroup();
			//test.TestGetValidationsTypeForValueID();
			//test.TestIsValidationTypeInValidationList();
			//test.TestGetFirstComboValueId();
			//test.TestGetValidationTypeForValueID();
			//test.TestGetValueIdsForValidationType();
			//test.TestGetValueIdsForValueName();
			//test.TestIsValidationType();
			//test.TestGetComboText();
						
			//GlobalAssistTest test = new GlobalAssistTest();
			//test.TestGetAllGlobalBandedParamValuesAsXml();
			//test.TestGetGlobalParamPercentage();
			//test.TestGetGlobalParamMaximumAmount();
			//test.TestGetGlobalParamAmount();
			//test.TestGetGlobalParamAmountAsDouble();
			//test.TestGetGlobalParamMaximumAmountAsDouble();
			//test.TestGetGlobalParamString();
			//test.TestGetGlobalParamBoolean();

			//SqlAssistTest test = new SqlAssistTest();
			//test.TestGuidConversions();
			//test.TestFormatGuid();

			//GeminiTest test = new GeminiTest();
			//test.TestSendToFulfillment();
			//test.TestIsFileVersionLocked();

			//AdoAssistTest test = new AdoAssistTest();
			//test.TestGetValueFromTable();
			//test.TestCheckSingleRecordExists();
			//test.TestExecuteSqlCommand();
			//test.TestDeleteAdditionalBorrowingFee();
			//test.TestGetAsXMLApplication();
			//test.TestCreateFromNodeAdditionalBorrowingFee();
			//test.TestLoadSchema();
			//test.TestGetStoredProcAsXML();
			//test.TestUpdateAdditionalBorrowingFee();
			//test.TestCreateAdditionalBorrowingFee();
			//test.TestGetGeneratedKeyShort();
			//test.TestGetGeneratedKeyGuid();
			//test.TestGetRecordSetAsXML();
			//test.TestBuildDbConnectionString();

			//ErrAssistExTest test = new ErrAssistExTest();
			//test.TestGetMessageText();
			//test.TestThrowError();

			//XmlAssistTest test = new XmlAssistTest();
			//test.TestMakeNodeElementBased();
			//test.TestParserError();
			//test.TestLoad();
			//test.TestGetRequestNode();

			//GuidAssistTest test = new GuidAssistTest();
			//test.TestCreateGUID();

			//EncryptAssistTest test = new EncryptAssistTest();
			//test.TestEncrypt();

			//ConvertAssistTest test = new ConvertAssistTest();
			//test.TestCSafeBool();

			/*
			using (TransactionScope ts = new TransactionScope(TransactionScopeOption.RequiresNew))
			{
				//Create and open the SQL connection.  The work done on this connection will be a part of the transaction created by the TransactionScope
				SqlConnection myConnection = new SqlConnection("server=SCH015107;Integrated Security=SSPI;database=northwind");
				SqlCommand myCommand = new SqlCommand();
				myConnection.Open();
				myCommand.Connection = myConnection;

				//Restore database to near it's original condition so sample will work correctly.
				myCommand.CommandText = "DELETE FROM Region WHERE (RegionID = 100) OR (RegionID = 101)";
				myCommand.ExecuteNonQuery();

				//Insert the first record.
				myCommand.CommandText = "Insert into Region (RegionID, RegionDescription) VALUES (100, 'MidWestern')";
				myCommand.ExecuteNonQuery();

				//Insert the second record.
				myCommand.CommandText = "Insert into Region (RegionID, RegionDescription) VALUES (101, 'MidEastern')";
				myCommand.ExecuteNonQuery();

				myConnection.Close();

				//Call complete on the TransactionScope or not based on input
				ConsoleKeyInfo c;
				while (true)
				{
					Console.Write("Complete the transaction scope? [Y|N] ");
					c = Console.ReadKey();
					Console.WriteLine();

					if ((c.KeyChar == 'Y') || (c.KeyChar == 'y'))
					{
						// Commit the transaction
						ts.Complete();
						break;
					}
					else if ((c.KeyChar == 'N') || (c.KeyChar == 'n'))
					{
						break;
					}
				}

			}
			 */
			try
			{
				//omAU.AuditBO auditBO = new omAU.AuditBO();
				//auditBO.CreateAccessAudit("");

				/*
				Type typAuditTxBO = Type.GetTypeFromProgID("omAU.AuditTxBO", true);
				object omAUObject = Activator.CreateInstance(typAuditTxBO);
				omAU.IAuditTxBO objIAuditTxBO = (omAU.IAuditTxBO)omAUObject;
				objIAuditTxBO.ToString();
				*/

				//int x = (int)OMIGAERROR.oeRecordNotFound;

				/*
				string request = 
					"<REQUEST " + 
						"CRUD_OP='READ' " + 
						"ENTITY_REF='GLOBALPARAMETER' " +
						"SCHEMA_NAME='omCRUD'>" +
						"<GLOBALPARAMETER NAME='GeminiEnvironment'/>" +
					"</REQUEST>";
				omCRUD.omCRUDBO bo = new omCRUD.omCRUDBO();
				 */
				//object obj = ContextUtility.CreateInstance("omCRUD.omCRUDBO");
				//omCRUD.omCRUDBO bo = (omCRUD.omCRUDBO)obj;
				//string response = 
				//    Convert.ToString(
				//        obj.GetType().InvokeMember(
				//            "OmRequest", 
				//            BindingFlags.Public | BindingFlags.InvokeMethod, 
				//            null, 
				//            obj, 
				//            new object[] { request }));
				//string response = bo.OmRequest(request);
				//Console.WriteLine(response);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.ToString());
			}
		}
/*
public static XmlNode FindQQRegularOutgoings()
{
	XmlNode xmlResponseNode = null;
	XmlDocument xmlOut = new XmlDocument();
	XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
	xmlOut.AppendChild((XmlNode)xmlResponseElement);

	ResumeNext resumeNext = new ResumeNext(OMIGAERROR.oeRecordNotFound, xmlResponseElement);
	try
	{
		//statement1;	// May raise a record not found error.
		//statement2;
	}
	catch (Exception exception)
	{
		xmlResponseNode = new ErrAssistException(exception).CreateErrorResponseNode();
		ContextUtility.SetComplete();
	}
	finally
	{
		resumeNext.Dispose();
	}

	return xmlResponseNode;
}
*/
		
		/*


FindQQRegularOutgoingsExit:
Exit Function

FindQQRegularOutgoingsVbErr:

    If Err.Number = omiga4RecordNotFound Then
        Resume Next
    End If

    If objErrAssist.IsWarning = True Then
        ' add message element to response block
        objErrAssist.AddWarning xmlResponseElement
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError _
TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindQQRegularOutgoings = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindQQRegularOutgoingsExit
End Function
		*/
	}
}
