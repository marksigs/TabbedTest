using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class TraceAssistTest
	{
		private const string _moduleName = "TraceAssistTest";

		[Test]
		public void TestTrace()
		{
			const string methodName = "TestTrace";

			TraceAssist tracer = new TraceAssist(_moduleName);
			tracer.TraceMethodEntry(_moduleName, methodName, "MethodEntry");
			tracer.TraceMethodExit(_moduleName, methodName, "MethodExit");
			tracer.TraceMethodError(_moduleName, methodName, new Exception("This is an exception"), "Exception message");
			tracer.TraceRequest("<REQUEST/>");
			tracer.TraceResponse("<RESPONSE/>");
			string response = 
				"<RESPONSE>" + 
					"<ERROR>" +
						"<NUMBER>0</NUMBER>" + 
						"<DESCRIPTION>This is an error.</DESCRIPTION>" + 
					"</ERROR>" +
				"</RESPONSE>";
			tracer.TraceIdentErrorResponse(ref response);
		}
	}
}
