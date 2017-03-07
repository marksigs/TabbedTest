/*
--------------------------------------------------------------------------------------------
Workfile:			HomeTrackBO.cs
Copyright:			Copyright © 2005 Marlborough Stirling

Description:		It is a desktop valuation service. 
					Called from Omiga through the Direct Gateway.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
HM		10/10/2005	MAR41 Created
JD		21/10/2005	MAR263 Add SUCCESS attribute to returned response if successful.
MV		26/10/2005	MAR278 Included logging 
JD		08/11/2005	MAR263 changes to error reporting
HMA     09/11/2005  MAR462 Remove req.AccountID in line with latest xsd.
JD		11/11/2005	MAR503 customer node passed in is attribute based.
JD		11/11/2005	MAR506 save the valuationresult to the database.
JD		15/11/2005	MAR547 ensure the date returned from hometrack is in the right format.
JD		16/11/2005  MAR562 Check number of bedrooms before sending. Ensure RESPONSE returned is always valid.
GHun	16/11/2005	MAR584 Fix ValuationRules request XML, and added extra logging
JD		25/11/2005	MAR685 Check for errors correctly.
JD		30/11/2005	MAR719 Check for legal propertytype
JD		30/11/2005	MAR721 don't update the mortgagesubquote.amountrequested after calc'ing LTV
JD		01/12/2005	MAR643 Return the esurv taskid if there are errorcodes returned from Hometrack
JD		02/12/2005	MAR785 Check for Gateway error as well as Hometrack error
PSC		06/12/2005	MAR814 Correct responses sent back if an error occurs and rewrite SaveValuation
						   to correct errors and used parameterized SQL
PSC		07/12/2005	MAR793 Amend to get the combo values and global parameters for the call to 
					ValuationRules
GHun	12/12/2005	MAR852 set Operator in generic request			   
PSC		13/12/2005	MAR457 Use omLogging wrapper
RF		28/12/2005	MAR940 Home Track task is showing N/A status and  calling Esurv irrespective of confidence level - 
						   need to add extra debug info and late bind to valuation rules
DRC     05/01/2005  MAR996 ValuationRules Request needs global parameter HomeTrackMinLTV	as well				   
JD		17/02/2006	MAR1280	Now using omRB with template APValnRBTemplate to get data to pass to the valuation rules
IK		08/03/2006	EP1999 Epsom stub
--------------------------------------------------------------------------------------------
*/

using System;
using System.Net; 
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Reflection; 

namespace Vertex.Fsd.Omiga.omHomeTrack
{
	/// <summary>
	/// Summary description for HomeTrack
	/// </summary>
	/// 
//	[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
//	public interface IHomeTrackBO
//	{
//		[DispId(600)]string RunHomeTrackValuation (string strRequest);
//	}
//	[ClassInterface(ClassInterfaceType.AutoDual)]
	[Guid("FEFCE385-505C-4431-963C-C01B28326721")]
	[ComVisible(true)]
	[ProgId("omHomeTrack.HomeTrackBO")]
	public class HomeTrackBO //: IHomeTrackBO
	{
		public HomeTrackBO()
		{
		}
		
		//Return the latest hometrack valuation amount for the passed in appliaction number
		public string GetPresentValuation(string strRequest)
		{
			return "<RESPONSE TYPE='SUCCESS'><HOMETRACKVALUATION/></RESPONSE>";
		}

		public string RunHomeTrackValuation (string strRequest)
		{
			return "<RESPONSE TYPE='SUCCESS'><HOMETRACKVALUATION/></RESPONSE>";
		}
	}
}
