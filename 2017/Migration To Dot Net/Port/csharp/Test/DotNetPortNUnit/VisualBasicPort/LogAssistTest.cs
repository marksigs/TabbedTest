using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class LogAssistTest
	{
		[Test]
		public void Test()
		{
			LogAssist.SetProfilingOn();
			LogAssist logAssist1 = new LogAssist(System.Reflection.Assembly.GetExecutingAssembly());
			logAssist1.SetComponentProfilingOn();
			logAssist1.StartTimer("Test1", true);
			Thread.Sleep(5000);
			logAssist1.StopTimer();

			using (LogAssist logAssist2 = new LogAssist(System.Reflection.Assembly.GetExecutingAssembly(), "Test2"))
			{
				Thread.Sleep(5000);
			}

		}
	}
}
