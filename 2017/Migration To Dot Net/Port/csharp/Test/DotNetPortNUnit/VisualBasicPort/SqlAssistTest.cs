using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class SqlAssistTest
	{
		[Test]
		public void TestFormatGuid()
		{
			string guidText = null;
			guidText = SqlAssist.FormatGuid("DA6DA163412311D4B5FA00105ABB1680", GUIDSTYLE.guidHyphen);
			Assert.AreEqual(guidText, "{63A16DDA-2341-D411-B5FA-00105ABB1680}");
			guidText = SqlAssist.FormatGuid("{63A16DDA-2341-D411-B5FA-00105ABB1680}", GUIDSTYLE.guidBinary);
			Assert.AreEqual(guidText, "DA6DA163412311D4B5FA00105ABB1680");
			guidText = SqlAssist.FormatGuid("{63A16DDA-2341-D411-B5FA-00105ABB1680}", GUIDSTYLE.guidLiteral);
			Assert.AreEqual(guidText, "0xDA6DA163412311D4B5FA00105ABB1680");
		}

		[Test]
		public void TestGuidConversions()
		{
			string guidText1 = null;
			string guidText2 = null;
			byte[] guidBytes = null;

			guidText1 = "DA6DA163412311D4B5FA00105ABB1680";
			guidBytes = SqlAssist.GuidStringToByteArray(guidText1);
			guidText2 = SqlAssist.GuidToString(guidBytes, System.Data.SqlDbType.Binary, GUIDSTYLE.guidHyphen);
			Assert.AreEqual(guidText2, "{63A16DDA-2341-D411-B5FA-00105ABB1680}");

			guidText1 = "63A16DDA2341D411B5FA00105ABB1680";
			guidBytes = SqlAssist.GuidStringToByteArray(guidText1);
			guidText2 = SqlAssist.GuidToString(guidBytes, System.Data.SqlDbType.UniqueIdentifier, GUIDSTYLE.guidBinary);
			Assert.AreEqual(guidText2, "DA6DA163412311D4B5FA00105ABB1680");

			using (SqlConnection adoConnection = new SqlConnection(AdoAssist.GetDbConnectString()))
			{
				using (SqlCommand adoCommand = new SqlCommand("select * from dotnetporttemp", adoConnection))
				{
					adoCommand.CommandType = CommandType.Text;
					adoConnection.Open();

					using (SqlDataReader adoRecordSet = adoCommand.ExecuteReader(CommandBehavior.CloseConnection))
					{
						if (adoRecordSet.Read())
						{
							Guid guid = adoRecordSet.GetGuid(0);
							guidText1 = SqlAssist.GuidToString(guid);
							Assert.AreEqual(guidText1, "D86F6228BBB46945BE469664ECA6C790");

							SqlBinary sqlBinary = adoRecordSet.GetSqlBinary(1);
							guidText1 = SqlAssist.GuidToString(sqlBinary);
							Assert.AreEqual(guidText1, "10045B4381DA41A88491FB435032B6CC");
						}
					}
				}
			}
		}
	}
}
