using System;
using System.Xml;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.omBase
{
	[Obsolete("Use omBaseClassDef instead")]
    public class IomBaseClassDef
    {
		[Obsolete("Use omBaseClassDef.LoadMessageData instead")]
        public XmlDocument LoadMessageData() 
        {
			return omBaseClassDef.LoadMessageData();
        }

		[Obsolete("Use omBaseClassDef.LoadCurrencyData instead")]
        public XmlDocument LoadCurrencyData() 
        {
			return omBaseClassDef.LoadCurrencyData();
        }
    }
}
