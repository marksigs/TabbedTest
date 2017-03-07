using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Vertex.Fsd.Omiga.omBase.Obsolete
{
	[Obsolete("Use CurrencyDO instead")]
	public class ICurrencyDO
	{
		[Obsolete("Use CurrencyDO.FindList instead")]
		public XmlNode FindList()
		{
			XmlNode xmlNode = null;
			using (CurrencyDO currencyDO = new CurrencyDO())
			{
				xmlNode = currencyDO.FindList();
			}
			return xmlNode;
		}
	}
}
