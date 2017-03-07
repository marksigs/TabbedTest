using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Text;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class ComboAssistTest
	{
		[Test]
		public void TestGetComboGroup()
		{
			ComboAssist.GetComboGroup("OmigaComponents");
			ComboAssist.GetComboGroups("OmigaComponents,Title");
			ComboAssist.GetComboGroups("AccessArrangements,AccessAuditType");
		}

		[Test]
		public void TestBaseComboDOGetComboList()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetComboList("<REQUEST><LIST><LISTNAME>AccessArrangements</LISTNAME></LIST></REQUEST>");
			}
		}

		[Test]
		public void TestBaseComboDOGetComboValue()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetComboValue("<LIST><GROUPNAME>AccessArrangements</GROUPNAME><VALUEID>1</VALUEID></LIST>");
			}
		}

		[Test]
		public void TestBaseComboDOGetComboText()
		{
			string response = "";
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				response = comboDO.GetComboText("AccessArrangements", "1");
			}
			Assert.AreEqual(response, "Vendor");
		}

		[Test]
		public void TestBaseComboDOIsItemInValidation()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				bool is1 = comboDO.IsItemInValidation("AccessArrangements", "1", "O");
				Assert.IsTrue(is1);
				bool is2 = comboDO.IsItemInValidation("AccessArrangements", "1", "A");
				Assert.IsFalse(is2);
			}
		}

		[Test]
		public void TestBaseComboDOGetFirstComboValueId()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetFirstComboValueId("AccessArrangements", "O");
				Assert.AreEqual(response, "1");
			}
		}

		[Test]
		public void TestBaseComboDOGetComboValueId()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetComboValueId("AccessArrangements", "O");
			}
		}
		
		[Test]
		public void TestBaseComboDOGetFirstComboValueIdFromValueName()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetFirstComboValueIdFromValueName("AccessArrangements", "Vendor");
				Assert.AreEqual(response, "1");
			}
		}

		[Test]
		public void TestBaseComboDOGetComboValueIdFromValueName()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetComboValueIdFromValueName("AccessArrangements", "Vendor");
			}
		}
		
		[Test]
		public void TestBaseComboDOGetNewLoanValue()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetNewLoanValue();
			}
		}
		
		[Test]
		public void TestBaseComboDOGetFirstComboValidation()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetFirstComboValidation("AccessArrangements", "1");
				Assert.AreEqual(response, "O");
			}
		}
		
		[Test]
		public void TestBaseComboDOGetQuickQuoteLocationValueId()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetQuickQuoteLocationValueId();
				Assert.AreEqual(response, "10");
			}
		}

		[Test]
		public void TestBaseComboDOGetQuickQuoteValuationTypeValueId()
		{
			using (omBase.ComboDO comboDO = new omBase.ComboDO())
			{
				string response = comboDO.GetQuickQuoteValuationTypeValueId();
				Assert.AreEqual(response, "3");
			}
		}

		[Test]
		public void TestBaseComboBOGetComboValue()
		{
			using (omBase.ComboBO comboBO = new omBase.ComboBO())
			{
				string response = comboBO.GetComboValue("<REQUEST><LIST><GROUPNAME>AccessArrangements</GROUPNAME><VALUEID>1</VALUEID></LIST></REQUEST>");
			}
		}

		[Test]
		public void TestBaseComboBOGetComboList()
		{
			using (omBase.ComboBO comboBO = new omBase.ComboBO())
			{
				string response = comboBO.GetComboList("<REQUEST><LIST><LISTNAME>AccessArrangements</LISTNAME></LIST></REQUEST>");
			}
		}

		[Test]
		public void TestGetComboText()
		{
			for (int index = 0; index < 100000; index++)
			{
				ComboAssist.GetComboText("AccessArrangements", 1);
			}

			for (int index = 0; index < 100000; index++)
			{
				ComboAssist.GetComboText("AccessArrangements", 1);
			}

			string text = null;
			text = ComboAssist.GetComboText("AccessArrangements", 1);
			Assert.AreEqual(text, "Vendor");
			text = ComboAssist.GetComboText("AccessArrangements", 1);
			Assert.AreEqual(text, "Vendor");
			text = ComboAssist.GetComboText("AccessAuditType", 2);
			Assert.AreEqual(text, "System logoff");
			text = ComboAssist.GetComboText("AccessArrangements", 1);
			Assert.AreEqual(text, "Vendor");
		}

		[Test]
		public void TestIsValidationType()
		{
			bool success = false;
			success = ComboAssist.IsValidationType("AccessArrangements", 1, "A");
			Assert.IsFalse(success);
			success = ComboAssist.IsValidationType("AccessArrangements", 1, "V");
			Assert.IsTrue(success);
			success = ComboAssist.IsValidationType("AccessArrangements", 1, "A");
			Assert.IsFalse(success);
			success = ComboAssist.IsValidationType("AccessArrangements", 1, "V");
			Assert.IsTrue(success);
		}

		/*
		[Test]
		public void TestIsValidationTypeInValidationList()
		{
			bool success = false;
			success = ComboAssist.IsValidationTypeInValidationList("AccessArrangements", "A", 1);
			Assert.IsFalse(success);
			success = ComboAssist.IsValidationTypeInValidationList("AccessArrangements", "V", 1);
			Assert.IsTrue(success);
			success = ComboAssist.IsValidationTypeInValidationList("AccessArrangements", "A", 1);
			Assert.IsFalse(success);
			success = ComboAssist.IsValidationTypeInValidationList("AccessArrangements", "V", 1);
			Assert.IsTrue(success);
		}
		*/

		[Test]
		public void TestGetValueIdsForValidationType()
		{
			ReadOnlyCollection<int> valueIds = null;
			valueIds = ComboAssist.GetValueIdsForValidationType("AccessArrangements", "V");
			Assert.AreEqual(valueIds.Count, 2);
			valueIds = ComboAssist.GetValueIdsForValidationType("AccessArrangements", "V");
			Assert.AreEqual(valueIds.Count, 2);
		}

		[Test]
		public void TestGetValueIdsForValueName()
		{
			ReadOnlyCollection<int> valueIds = null;
			valueIds = ComboAssist.GetValueIdsForValueName("AccessArrangements", "Vendor");
			Assert.AreEqual(valueIds.Count, 2);
			valueIds = ComboAssist.GetValueIdsForValueName("AccessArrangements", "Vendor");
			Assert.AreEqual(valueIds.Count, 2);
		}

		[Test]
		public void TestGetValidationTypeForValueID()
		{
			string validationType = null;
			validationType = ComboAssist.GetValidationTypeForValueID("AccessArrangements", 1);
			Assert.AreEqual(validationType, "O");
			validationType = ComboAssist.GetValidationTypeForValueID("AccessArrangements", 1);
			Assert.AreEqual(validationType, "O");
		}
		
		[Test]
		public void TestGetFirstComboValueId()
		{
			int valueId = 0;
			valueId = ComboAssist.GetFirstComboValueId("AccessArrangements", "V");
			Assert.AreEqual(valueId, 1);
			valueId = ComboAssist.GetFirstComboValueId("AccessArrangements", "V");
			Assert.AreEqual(valueId, 1);
		}

		[Test]
		public void TestGetValidationsTypeForValueID()
		{
			StringCollection validationTypes = null;
			validationTypes = ComboAssist.GetValidationTypesForValueID("AccessArrangements", 2);
			Assert.IsTrue(validationTypes.Count == 3);
			Assert.AreEqual(validationTypes[1], "O");
			validationTypes = ComboAssist.GetValidationTypesForValueID("AccessArrangements1", 2);
			Assert.IsTrue(validationTypes.Count == 3);
			Assert.AreEqual(validationTypes[1], "O");
		}		

	}
}

