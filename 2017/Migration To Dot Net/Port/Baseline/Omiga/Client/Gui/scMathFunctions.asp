<%@  Language=JScript %>
<%/*
Workfile:      scXMLFunctions.htm
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Helper functions for XML parser.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		31/01/00	Converted to .asp for optimisation
AY		16/02/00	SYS0240 - HEAD and TITLE elements removed
PF		23/04/01	SYS2252 - new Rounding variant allows specification
					of rounding direction
MV		10/04/02	SYS4376 - Modified RoundValue to Cater for Negative Values
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History
Prog	Date		Description
GHun	28/07/2004	BMIDS823 (BBG498) Check that the number of decimal places is not > the
						decimal precision required in PadToPrecision to avoid infinite
						loop if PadToPrecision is called directly (SHOULD ONLY BE CALLED
						INDIRECTLY THROUGH ROUNDVALUE).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Mars History
Prog	Date		Description
GHun	08/05/2006	MAR1648 Changed RoundValue to use precision passed in, and not always use 2
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
List the external function calls and their parameters here:

iNumber = RoundValue(iNumber, iDecimalPrecision)
	iNumber				- the number to be rounded
	iDecimalPrecision	- the number of decimal places to round to
			
	Rounds a number to the Decimal Precision indicated.
	This routine calls GetRoundingFactor.


iNumber = RoundWithDirection(iNumber, iDecimalPrecision, iRoundingDirection)
	iNumber				- the number to be rounded
	iDecimalPrecision	- the number of decimal places to round to
	iRoundingDirection  - the direction in which the Rounding should be forced
			
	Rounds a number to the precision specified.
	Ignores normal rounding rules and goes with the direction specified, where 10 is Up and 20 is Down.
	
				
iRoundingFactor = GetRoundingFactor(iDecimalPrecision)
	iDecimalPrecision	- the number of decimal places to round to
			
	Calculates the rounding factor to apply in rounding a number to a certain 
	Precision. Not an external routine.

		
iNumber = Truncate(iNumber)
	Truncates a number to just its integer part i.e. 1.234 -> 1			
	iNumber			- the number to be truncated

			
nYears = GetYearsBetweenDates(dteFromDate, dteToDate)
	dteFromDate		- date object for start date
	dteToDate		- date object for end date
	nYears			- number of whole years between the start and end dates
			
	Returns the number of whole years between the 2 dates or -1 if either date
	is invalid, or if the start date > end date. 
	eg. if start date = 18/11/1969 and end date = 19/11/1999 then nYears would be 30
	    if start date = 20/11/1969 and end date = 19/11/1999 then nYears would be 29
	    if start date = 19/11/1969 and end date = 19/11/1999 then nYears would be 30 
*/
%>
<HTML id=MathFunctions>
<script language="JavaScript">

public_description = new CreateMathFunctions;

<% /* Initialises the pointers to the externally accessible functions. */ %>
function CreateMathFunctions()
{
	this.Truncate = Truncate;
	this.RoundValue = RoundValue;
	this.GetRoundingFactor = GetRoundingFactor;
	this.PadToPrecision = PadToPrecision;
	this.GetYearsBetweenDates = GetYearsBetweenDates;
	this.RoundWithDirection = RoundWithDirection;
}

function RoundValue(sNumber, iDecimalPrecision)
{
<%	// Rounds a number (sNumber) to the specified number of
	// decimal places (iDecimalPrecision)
%>	var iNumber = parseFloat(sNumber);

<%	// SR 26/02/2002 : SYSMCP0196 - if the number passed in does not have places more than 
	// iDecimalPrecision, just return the number
%>
	if (isNaN(iNumber)) return "" ;
	
	var iDecimalPlaces;
	sNumber = sNumber.toString(10);
	var iDPIndex = sNumber.indexOf(".");
	
	if(iDPIndex == -1) 
		return this.PadToPrecision(iNumber, iDecimalPrecision);	<% //MAR1648 GHun %>
	else
	{
		iDecimalPlaces = sNumber.length - iDPIndex - 1 ;
		if(iDecimalPlaces <= iDecimalPrecision) 
			return this.PadToPrecision(iNumber, iDecimalPrecision);
	}
	
	if (!isNaN(iNumber))
	{
		var iRoundingFactor = this.GetRoundingFactor(iDecimalPrecision);
		
		if ( iNumber < 0 ) 
			iNumber = iNumber * (-1) ;	
		
		if (iRoundingFactor != 1) iNumber = iNumber + (0.5 / iRoundingFactor);

		iNumber = iNumber * iRoundingFactor;
		iNumber = parseInt(iNumber);
		iNumber = iNumber / iRoundingFactor;

		if (parseFloat(sNumber) < 0 ) 
		{
			var ReturnVal = this.PadToPrecision(iNumber, iDecimalPrecision);
			return (-1) * ReturnVal ; 
		}
		else
			return this.PadToPrecision(iNumber, iDecimalPrecision);
	}

	return "";
}

function GetRoundingFactor(iDecimalPrecision)
{
<%	// Calculates the rounding factor to apply from the required decimal precision
	// i.e. iDecimalPrecision = 0 -> iRoundingFactor = 1
	//		iDecimalPrecision = 1 -> iRoundingFactor = 10
	//		...
	//		iDecimalPrecision = 3 -> iRoundingFactor = 10*10*10 = 1000
	// and so on...
%>	if((!isNaN(iDecimalPrecision)) && (iDecimalPrecision > 0)) return Math.pow (10,iDecimalPrecision)

	return 1;
}

<%	/* BMIDS823 (BBG498) If this was called directly and the number of decimal places
	of iNumber was greater than the iDecimalPrecision required, this function 
	went into an infinite loop, because iLoop could never be negative. THIS
	SHOULD NEVER BE CALLED DIRECTLY, BUT SHOULD BE CALLED INDIRECTLY VIA THE
	RoundValue FUNCTION ABOVE*/
%>
function PadToPrecision(iNumber, iDecimalPrecision)
{
<%	// Pads a value (iNumber) to the required number of decimal places
	// (iDecimalPrecision). Because of the following values i.e. 25.50
	// being translated to 25.5 we need to pad the figure to the correct
	// number of decimal places
%>	var sNumber = iNumber.toString();
	var iDPIndex = sNumber.indexOf(".");

	if ((iDPIndex != -1) || (iDecimalPrecision != 0))
	{
		var iNumberOfDP = 0;

		if (iDPIndex == -1)
		{
			sNumber	+= ".";
			iNumberOfDP	= 0;
		}
		else iNumberOfDP = sNumber.length - 1 - iDPIndex;
		
		<%	/* BMIDS823 Check that the number of decimal places is not > the decimal precision
			required to avoid infinite loop */ %>		
		if(iNumberOfDP < iDecimalPrecision) 
		{
			for (var iLoop = 0; iLoop != (iDecimalPrecision - iNumberOfDP); iLoop++) sNumber += "0";
		}
	}

	return sNumber;
}

function Truncate(iNumber)
{
	if (!isNaN(iNumber)) iNumber = Math.floor(iNumber);

	return iNumber;
}

function GetYearsBetweenDates(dteFromDate, dteToDate)
{
	var nNumYears = -1;

	if(dteFromDate != null && dteToDate != null)
	{
		if(dteFromDate < dteToDate)
		{
			nNumYears = dteToDate.getFullYear() - dteFromDate.getFullYear();
			if(dteToDate.getMonth() <= dteFromDate.getMonth())
			{
				if(dteToDate.getMonth() < dteFromDate.getMonth()) nNumYears -= 1;
				else if(dteToDate.getDate() < dteFromDate.getDate()) nNumYears -=1;
			}
		}
		else alert("Date is invalid");
	}

	return(nNumYears);
}


function RoundWithDirection(Result, Precision, Direction)
	{
		//take the number and multiply it by 10 for the appropriate precision value
		var tempResult = Result * (Math.pow(10, Precision));
		//get the value of the figure immediately after the decimal point
		var strResult = tempResult.toString();
		var iDPIndex = strResult.indexOf(".");
		if (iDPIndex != "-1")
			{
			var RoundBit = strResult.charAt(iDPIndex + 1);
		
			if (Direction != "10")
				{
				if (Direction != "20")
					{
						if (Direction != "30")
							{
								alert("Unknown Rounding Direction");
							}
					}
				//chop it back to the right format
				else 
				
					{
						var formattedResult = tempResult / (Math.pow(10, Precision));
						return formattedResult;
					}
				}
				
			else
				{	//convert it to an integer
					tempResult = Math.floor(tempResult);
					if (RoundBit > 0)
						{
							if (Direction == "10")
								{
									tempResult = tempResult + 1;
								}
						}
					//chop it back to the right format
					var formattedResult = tempResult / (Math.pow(10, Precision));
					return formattedResult;
				}
			}
		//chop it back to the right format
					var formattedResult = tempResult / (Math.pow(10, Precision));
					return formattedResult;		
	}


</script>
