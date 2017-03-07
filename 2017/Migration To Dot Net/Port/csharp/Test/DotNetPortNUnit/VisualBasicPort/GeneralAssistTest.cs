using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class GeneralAssistTest
	{
		[Test]
		public void TestContainsSpecialChars()
		{
			bool success = false;
			success = GeneralAssist.ContainsSpecialChars("ABCdef");
			Assert.IsFalse(success);
			success = GeneralAssist.ContainsSpecialChars("ABCdef%");
			Assert.IsTrue(success);
		}

		[Test]
		public void TestIsAlpha()
		{
			bool success = false;
			success = GeneralAssist.IsAlpha("ABCdef");
			Assert.IsTrue(success);
			success = GeneralAssist.IsAlpha("ABCdef1");
			Assert.IsFalse(success);
		}

		[Test]
		public void TestHasDuplicatedChars()
		{
			bool hasDuplicatedChars = false;
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("AAAa123", false);
			Assert.IsTrue(hasDuplicatedChars);
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("111AAA111", false);
			Assert.IsTrue(hasDuplicatedChars);
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("Q2q345", false);
			Assert.IsTrue(hasDuplicatedChars);
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("Fred11", false);
			Assert.IsTrue(hasDuplicatedChars);
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("ABCDEFG", false);
			Assert.IsFalse(hasDuplicatedChars);
			hasDuplicatedChars = GeneralAssist.HasDuplicatedChars("1234567", false);
			Assert.IsFalse(hasDuplicatedChars);
		}

		[Test]
		public void TestConvertToMixedCase()
		{
			string text = null;
			text = GeneralAssist.ConvertToMixedCase("this");
			Assert.AreEqual(text, "This");
			text = GeneralAssist.ConvertToMixedCase("this is ");
			Assert.AreEqual(text, "This Is ");
			text = GeneralAssist.ConvertToMixedCase("this is a test string");
			Assert.AreEqual(text, "This Is A Test String");
		}
	}
}
