using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.omBase
{
	public static class omBaseClassDef
	{
		public static XmlDocument LoadMessageData() 
		{
			const string xmlText = 
				"<TABLENAME>" +
					"MESSAGE" +
					"<PRIMARYKEY>MESSAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" +
					"<OTHERS>MESSAGETEXT<TYPE>dbdtString</TYPE></OTHERS>" +
					"<OTHERS>MESSAGETYPE<TYPE>dbdtString</TYPE></OTHERS>" +
				"</TABLENAME>";
			return XmlAssist.Load(xmlText);
		}

		public static XmlDocument LoadCurrencyData() 
		{
			const string xmlText = 
				"<TABLENAME>" +
					"CURRENCY" +
					"<PRIMARYKEY>CURRENCYCODE<TYPE>dbdtInt</TYPE></PRIMARYKEY>" +
					"<OTHERS>BASECURRENCYIND<TYPE>dbdtBoolean</TYPE></OTHERS>" +
					"<OTHERS>CURRENCYNAME<TYPE>dbdtString</TYPE></OTHERS>" +
					"<OTHERS>CONVERSIONRATE<TYPE>dbdtDouble</TYPE></OTHERS>" +
					"<OTHERS>ROUNDINGPRECISION<TYPE>dbdtInt</TYPE></OTHERS>" +
					"<OTHERS>ROUNDINGDIRECTION<TYPE>dbdtComboId</TYPE>" +
					"<COMBO>ROUNDINGDIRECTION</COMBO></OTHERS>" +
				"</TABLENAME>";
			return XmlAssist.Load(xmlText);
		}
	}
}
