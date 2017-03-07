/*
--------------------------------------------------------------------------------------------
Workfile:			ISystemDatesBO.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		
--------------------------------------------------------------------------------------------
Net Specific History:

Prog	Date		Description
AS		31/07/2007	First .Net version. Ported from ISystemDatesBO.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Xml;

namespace Vertex.Fsd.Omiga.omBase
{
	[Obsolete("Use SystemDatesBO instead")]
	public class ISystemDatesBO
    {
		[Obsolete("Use SystemDatesBO.CheckNonWorkingOccurence instead")]
		public XmlNode CheckNonWorkingOccurence(XmlElement xmlRequestElement) 
        {
			XmlNode xmlResponseNode = null;
			using (SystemDatesBO systemDatesBO = new SystemDatesBO())
			{
				xmlResponseNode = systemDatesBO.CheckNonWorkingOccurence(xmlRequestElement);
			}
			return xmlResponseNode;
        }

		[Obsolete("Use SystemDatesBO.FindWorkingDay instead")]
		public XmlNode FindWorkingDay(XmlElement xmlRequestElement) 
        {
			XmlNode xmlResponseNode = null;
			using (SystemDatesBO systemDatesBO = new SystemDatesBO())
			{
				xmlResponseNode = systemDatesBO.FindWorkingDay(xmlRequestElement);
			}
			return xmlResponseNode;
        }
    }
}
