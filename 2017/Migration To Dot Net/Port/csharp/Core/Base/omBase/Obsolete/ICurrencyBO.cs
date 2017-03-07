using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.omBase.Obsolete
{
	[Obsolete("Use CurrencyBO instead")]
	public class ICurrencyBO
	{
		[Obsolete("Use CurrencyBO.FindListAsXml instead")]
		public XmlNode FindList()
		{
			XmlNode xmlNode = null;
			using (CurrencyBO currencyBO = new CurrencyBO())
			{
				xmlNode = currencyBO.FindListAsXml();
			}
			return xmlNode;
		}
	}
}
