using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.omBase.Obsolete
{
	[Obsolete("Use ComboDO instead")]
	public class IComboDO
	{
		[Obsolete("Use ComboDO.GetNewLoanValue instead")]
		public string GetNewLoanValue()
		{
			string response = "";
			using (ComboDO comboDO = new ComboDO())
			{
				response = comboDO.GetNewLoanValue();
			}
			return response;
		}

		[Obsolete("Use ComboDO.GetQuickQuoteLocationValueId instead")]
		public string GetQuickQuoteLocationValueId()
		{
			string response = "";
			using (ComboDO comboDO = new ComboDO())
			{
				response = comboDO.GetQuickQuoteLocationValueId();
			}
			return response;
		}

		[Obsolete("Use ComboDO.GetQuickQuoteValuationTypeValueId instead")]
		public string GetQuickQuoteValuationTypeValueId()
		{
			string response = "";
			using (ComboDO comboDO = new ComboDO())
			{
				response = comboDO.GetQuickQuoteValuationTypeValueId();
			}
			return response;
		}

		[Obsolete("Use ComboDO.GetDormantLegalFeeValueId instead")]
		public string GetDormantLegalFeeValueId()
		{
			string response = "";
			using (ComboDO comboDO = new ComboDO())
			{
				response = comboDO.GetDormantLegalFeeValueId();
			}
			return response;
		}

		[Obsolete("Use ComboDO.GetRegSaleRelatedComboList instead")]
		public XmlNode GetRegSaleRelatedComboList(XmlElement xmlElement)
		{
			XmlNode xmlNode = null;
			using (ComboDO comboDO = new ComboDO())
			{
				xmlNode = comboDO.GetRegSaleRelatedComboList(xmlElement);
			}
			return xmlNode;
		}

	}
}
