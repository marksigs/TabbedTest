/*
--------------------------------------------------------------------------------------------
Workfile:			ISystemDatesDO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		31/07/2007	First .Net version. Ported from ISystemDatesDO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omBase
{
	[Obsolete("Use SystemDatesDO instead")]
    public class ISystemDatesDO
    {
		[Obsolete("Use SystemDatesDO.CheckNonWorkingOccurence instead")]
		public XmlNode CheckNonWorkingOccurence(XmlElement xmlRequestElement) 
        {
			XmlNode xmlResponseNode = null;
			using (SystemDatesDO systemDatesDO = new SystemDatesDO())
			{
				xmlResponseNode = systemDatesDO.CheckNonWorkingOccurence(xmlRequestElement);
			}
			return xmlResponseNode;
        }
    }
}
