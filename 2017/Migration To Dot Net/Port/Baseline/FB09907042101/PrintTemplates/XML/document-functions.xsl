<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	PE		03/08/2006	EP1031		Abstracted javascript functions into seperate xsl.
	PE		03/08/2006	EP1031		Added to SourceSafe
	PE		19/09/2006	EP1107		Added GetYears and GetMonths Javascript functions.
	OS		13/12/2006	EP2_455		Modified DealWithSalutation function
	OS		13/12/2006	EP2_455		Modified GetDate function
	MAH	15/12/2006	EP2_472		Added PackagerCase element for Ingested cases
	OS		29/12/2006	EP2_455		Added new templates for Packager data
    DS		29/12/2006	EP2_504		Added new templates for Broker data
	DS		08/01/2007	EP2_504 		Modified ARFIRM  and PRINCIPALFIRM template  for REVISEDOFFERCOPYTOBROKER
    DS		08/01/2007  EP2_501  	Modified PRINCIPALFIRM template  for OFFERCOPYTOPAKAGER according to new schema
	OS		08/01/2007	EP2_481		Added new method for handling FSA addresses
	OS		08/01/2007	EP2_721		Added new template for handling Guarantor data
	MAH		09/01/2007	EP2_473		Added IntroducerName where exists
	MAH		10/01/2007	EP2_473		Added function toPC to convert text strings to Proper Case	
	OS		10/01/2007	EP2_455		Added IntroducerName for Packager
	MAH		11/01/2007	EP2_483		Modified CustCount selection code
	PB		12/01/2007	EP2_734		Added DealWithAddressAndSalutation function.
	PB		17/01/2007	EP2_803		Added templates APPLICATION and INTRODUCER
	OS		17/01/2007	EP2_479		Modified template CUSTOMERVERSION to handle Guarantor data
	MAH		19/01/2007	EP2_472		Added Packagercase to APPLICATION and put date in braces
	OS		25/01/2007	EP2_896		Added separate attributes for AR and Principal firm names
	DS		26/01/2007	EP2_483		Added GetApplicantNames function and attribute for Applicant and Guarantor
	PB		27/01/2007	EP2_866		Added GetFullName function to return  the a name with the correct title and first name in full (not initial).
	PB		27/01/2007	EP2_866		Added CheckFortypes to detect a validation type within a pipe-separated list
	DS		31/01/2007	EP2_804		Added TENANCY transformation for Applicant and Gurantor
	GHun	31/01/2007	EP2_1147	Merged EP1225 Added APPDATE under Applicant
	PB		01/02/2007	EP2_821		Added FormatAddress
	DS		06/02/2007    EP2_489      Modified DealWithTypeOfApplication as per updated spec
	PB		06/02/2007	EP2_805		Updates to APPLICANT and GUARANTOR
	PB		06/02/2007	EP2_805		Added 2 functions from document-functions-applicant.xsl to avoid compatibility clashes
	PB		07/02/2007	EP2_805		Added missing function which was causing JS error, and renamed 2 functions as a precaution
	MAH		08/02/2007	EP2_815		Addded code to implement toPC() in address functions prior to addition of postcode which must be uppercase
	PB		08/02/2007	EP2_729		Added FormatNames() to format 1-2 basic names
	DS		08/02/2007	EP2_504		Modified APPLICATIONINTRODUCER template to take care of Individual Broker's address, applies to all the broker templates
    DS		09/02/2007	EP2_739   	Date formatting as per required format 
	DS		09/02/2007	EP2_722   	Added GUARANTORPOSTALADDRESS attribute for formatted address
	OS		12/02/2007	EP2_722		Reverted changes made to APPLICANT and GUARANTOR templates.
	DS		13/02/2007	EP2_1360	 	Formatted date
	DS		14/02/2007	EP2_489		Modified APPLICATIONINTRODUCER template to take care of Individual Broker's address, added  House number
	DS		14/02/2007	EP2_482       Modified toPC() function to take care of apostrophe(') in address.	
	OS		19/02/2007	EP2_1512	Modified DealWithTypeOfApplication to remove "RTB"
	PB		20/02/2007	EP2_808		Added NotFirstPage()
	OS		23/02/2007	EP2_479		Made changes to address handling functions, to exclude Firm Names from the capitalization function (toPC)
	PB		23/02/2007	EP2_801		Updated DealWithFSAAddress() function to cope with inconveniently formatted adresses
	===============================================================================================================-->
	<msxsl:script language="JavaScript" implements-prefix="msg"><![CDATA[

       var PageTemp = '';
       var LatestLeavingDate = '';
    String.prototype.trim = function() 
    {
		   return this.replace(/^\s+|\s+$/, '')
	}
	
		// PB 20/02/2007 EP2_808
		// Returns a 'Y' if this is not the first time the function is called
		var Called = '';
		function NotFirstPage()
		{
			var ReturnValue = '';
			ReturnValue = Called;
			Called = 'Y';
			return( ReturnValue );
		}
		
		// PB 08/07/2007 EP2_729
		// Returns a single name (without title) or two names separated with 'and'
		function FormatNames( Forename1, Surname1, Forename2, Surname2 )
		{
			var Output='';
			if( Forename2!='' )
			{
				Output = Forename1 + ' ' + Surname1 + ' and ' + Forename2 + ' ' + Surname2;
			}
			else
			{
				Output = Forename1 + ' ' + Surname1;
			}
			return( Output );
		}
		
		// PB 06/02/2007 EP2_805
		// This function can be called with either 1 or 2 applicants to return the correct salutation WITH the first initial
		function GetSingleOrJointSalutation2( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 )
		{
			var Salutation='';
			
			if( Title2 != '' )
			{
				//Joint
				Salutation = DealWithJointTitle2( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 );
			}
			else
			{
				//Single
				Salutation = GetApplicantNameWithInitials(Title1, TitleOther1, Forename1, Surname1);
			}
			return( Salutation )
		}
		
		// PB 06/02/2007 EP2_805
		// This function can be called with either 1 or 2 applicants to return the correct salutation WITHOUT the first initial
		function GetSingleOrJointShortSalutation( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 )
		{
			var Salutation='';
			
			if( Title2 != '' )
			{
				//Joint
				Salutation = DealWithJointSalutation( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 );
			}
			else
			{
				//Single
				Salutation = DealWithSalutation(Title1, TitleOther1, Forename1, Surname1);
			}
			return( Salutation )
		}
	
		function DealWithJointSalutation(Title1, TitleOther1, FirstName1, Surname1, Title2, TitleOther2, FirstName2, Surname2)
    {
		var strSalutation = "";
		
		if (Title1 != '' && Title2 != '' && Surname1 == Surname2)
		{
			if ((Title1 == 'Mr' && Title2 == 'Mrs') || (Title1 == 'Mrs' && Title2 == 'Mr'))
			{
				strSalutation = DealWithSalutation('Mr and Mrs', '', '', Surname1);
			}
			else
			{
				strSalutation = DealWithSalutation(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + DealWithSalutation(Title2, TitleOther2, FirstName2, Surname2);
			}
		}
		else
		{
			strSalutation = DealWithSalutation(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + DealWithSalutation(Title2, TitleOther2, FirstName2, Surname2);
		}
			
		return (strSalutation);
    }
    
	function DealWithJointTitle2(Title1, TitleOther1, FirstName1, Surname1, Title2, TitleOther2, FirstName2, Surname2)
    {
		var strAddressTitle = "";
		
		if (Title1 != '' && Title2 != '' && Surname1 == Surname2)
		{
			if (Title1 == 'Mr' && Title2 == 'Mrs')
			{
				strAddressTitle = GetApplicantNameWithInitials('Mr and Mrs', '',FirstName1, Surname1);
			}
			else if (Title1 == 'Mrs' && Title2 == 'Mr')
			{
				strAddressTitle = GetApplicantNameWithInitials('Mr and Mrs', '',FirstName2, Surname2);
			}
			else
			{
				strAddressTitle = GetApplicantNameWithInitials(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + GetApplicantNameWithInitials(Title2, TitleOther2, FirstName2, Surname2);
			}
		}
		else
		{
			strAddressTitle = GetApplicantNameWithInitials(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + GetApplicantNameWithInitials(Title2, TitleOther2, FirstName2, Surname2);
		}
			
		return (strAddressTitle);
    }
    
		// Returns true if "TypeToCheckFor" is present in the pipe-separated list "ListOfTypes"
		// If not found returns false
		function CheckForType( ListOfTypes, TypeToCheckFor )
		{
			var returnValue=false;
			var typesArray=ListOfTypes.split('|');
			for( loop=0 ; loop < typesArray.length; loop++)
			{
				if( typesArray[loop]==TypeToCheckFor )
				{
					returnValue=true;
					break;
				}
			}
			return returnValue;
		}
		
		function FormatAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
    	{
			var strAddress = "";
			if (FlatNumber != '')
			{
				strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
			}
			if (HouseName != '')
			{
				strAddress = strAddress + HouseName + '\\par ';
			}
			if (HouseNumber != '')
			{
				strAddress = strAddress + HouseNumber + ' ';
			}
			if (Street != '')
			{
				strAddress = strAddress + Street;
			}
			if (District != '')
			{
				if((Street != '') || (HouseNumber != ''))
				{
					strAddress = strAddress + '\\par ' + District;			
				}	
				else 
				{
					strAddress = strAddress + District;			
				}	
			}
			if (Town != '')
			{
				strAddress = strAddress + '\\par ' + Town;
			}
			if (County != '')
			{
				strAddress = strAddress + '\\par ' + County;
			}
			//MAH EP2_815
			strAddress = toPC(strAddress);
			
			if (PostCode != '')
			{
				strAddress = strAddress + '\\par ' + PostCode;
			}		
			return (strAddress);
		}
		
		function GetFullName( Title, TitleOther, Forename, Surname )
		{
			if( TitleOther!='' )
			{
				Title=TitleOther;
			}
			return( Title + ' ' + Forename + ' ' + Surname );
		}
		
        function getPageBreak()
        {
			var PageTemp2 = PageTemp;
			PageTemp = '\\page';
            return (PageTemp2);
        }   
    
        var Temp = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
        function UpdateCounter(WhichCounter)
        {
            Temp[WhichCounter] = Temp[WhichCounter] + 1;
            return (Temp[WhichCounter]);
        }

        function GetCounter(WhichCounter)
        {
            return (Temp[WhichCounter]);
        }        

        function DealWithStraightAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {

		var strAddress = "";
			
        if (FlatNumber != '')
		{
			strAddress = strAddress  + 'Flat ' + FlatNumber + ', ';
		}
        if (HouseName != '')
		{
			strAddress = strAddress + HouseName + ', ' ;
		}
        if (HouseNumber != '')
		{
			strAddress = strAddress + HouseNumber + ' ';
		}
        if (Street != '')
		{
			strAddress = strAddress + Street + ', ';;
		}
        if (District != '')
		{
			strAddress = strAddress + District + ', ';			
		}
        if (Town != '')
		{
			strAddress = strAddress + Town + ', ';
		}
        if (County != '')
		{
			strAddress = strAddress + County + ', ';
		}
		//MAH EP2_815
		strAddress = toPC(strAddress);

		if (PostCode != '')
		{
			strAddress = strAddress +  PostCode;
		}		
		return (strAddress);
        }
        
	function DealWithAddress(Title, TitleOther, FirstName, Surname,CompanyName,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
    	{
		var strAddress = "";
		
		if (Title != '')
		{
		   if ((TitleOther != '') && (Title == 'Other'))
		      strAddress = strAddress + toPC(TitleOther) + ' ';
		   else 
				strAddress = strAddress + toPC(Title) + ' ';			
		}	 
		    
		if ((FirstName !='') || (Surname !='')) 
		{
			if (FirstName !='')
			{
				strAddress = strAddress + toPC(FirstName).substring(0,1);
			}
			strAddress = strAddress +  ' ' + toPC(Surname) + '\\par ' ;
		}

		if (CompanyName != '')
		{
			strAddress = strAddress + CompanyName + '\\par ';
		}

		if (FlatNumber != '')
		{
			strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
		}
		    
		if (HouseName != '')
		{
			strAddress = strAddress + toPC(HouseName) + '\\par ';
		}
			
		if (HouseNumber != '')
		{
			strAddress = strAddress + HouseNumber + ' ';
		}
			
		if (Street != '')
		{
			strAddress = strAddress + toPC(Street);
		}
			
		if (District != '')
		{
			if((Street != '') || (HouseNumber != ''))
			{
				strAddress = strAddress + '\\par ' + toPC(District);			
			}	
			else 
			{
				strAddress = strAddress + toPC(District);			
			}	
		}
		    
		if (Town != '')
		{
			strAddress = strAddress + '\\par ' + toPC(Town);
		}
		    
		if (County != '')
		{
			strAddress = strAddress + '\\par ' + toPC(County);
		}
			
		if (PostCode != '')
		{
			strAddress = strAddress + '\\par ' + PostCode;
		}		
			
		return (strAddress);
	}
	
	//PB 12/01/2007 EP2_734
	function DealWithAddressAndSalutation( Salutation, FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode )
    	{
		var strAddress = "";
		
		if (Salutation != '')
		{
			strAddress = strAddress + Salutation + '\\par ';
		}

		if (FlatNumber != '')
		{
			strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
		}
		    
		if (HouseName != '')
		{
			strAddress = strAddress + HouseName + '\\par ';
		}
			
		if (HouseNumber != '')
		{
			strAddress = strAddress + HouseNumber + ' ';
		}
			
		if (Street != '')
		{
			strAddress = strAddress + Street;
		}
			
		if (District != '')
		{
			if((Street != '') || (HouseNumber != ''))
			{
				strAddress = strAddress + '\\par ' + District;			
			}	
			else 
			{
				strAddress = strAddress + District;			
			}	
		}
		    
		if (Town != '')
		{
			strAddress = strAddress + '\\par ' + Town;
		}
		    
		if (County != '')
		{
			strAddress = strAddress + '\\par ' + County;
		}
		//MAH EP2_815
		strAddress = toPC(strAddress);
			
		if (PostCode != '')
		{
			strAddress = strAddress + '\\par ' + PostCode;
		}		
			
		return (strAddress);
	}
	
	function DealWithFSAAddress(FirmName, Address1, Address2, Address3, Address4, Address5, Address6, PostCode)
    	{
    	
    	var addressArray = new Array();
		var strAddress = "";
		var index = 0;

		// Shuffle-up any blank lines as these will scupper the formatting
		var shuffleArray = new Array( Address1, Address2, Address3, Address4, Address5, Address6 );
		var iLoop	=0;
		for( iLoop=0; iLoop<=4; iLoop++ ) // indices will be 0-5
		{
			if( shuffleArray[ iLoop ]=='' )
			{
				shuffleArray[ iLoop ]=shuffleArray[ iLoop+1 ];
				shuffleArray[ iLoop+1 ]='';
			}
		}
		Address1 = shuffleArray[ 0 ];
		Address2 = shuffleArray[ 1 ];
		Address3 = shuffleArray[ 2 ];
		Address4 = shuffleArray[ 3 ];
		Address5 = shuffleArray[ 4 ];
		Address6 = shuffleArray[ 5 ];
		
		// Check whether first line is a street number which needs appending to line 2
		if( Address1!='' )
		{
			if( !isNaN( Address1.substr(1,1) ) )
			{
				// Line starts with a numeric value
				// Check for space
				if( Address1.indexOf(' ') >0 )
				{
					// Contains a space, so do nothing
				}
				else
				{
					// No space found - must be a street number
					// so append to the next line and shuffle the rest up
					Address1 = Address1 + ' ' + Address2;
					Address2 = Address3;
					Address3 = Address4;
					Address4 = Address5;
					Address5 = Address6;
					Address6 = '';
				}
			}
		}
		
		if (Address1 != '')
		{
			addressArray[index] = Address1;
			index++;
		}
		    
		if (Address2 != '')
		{
			addressArray[index] = Address2;
			index++;
		}
			
		if (Address3 != '')
		{
			addressArray[index] = Address3;
			index++;
		}
			
		if (Address4 != '')
		{
			addressArray[index] = Address4;
			index++;
		}
		
		if (Address5 != '')
		{
			addressArray[index] = Address5;
			index++;
		}
		
		if (Address6 != '')
		{
			addressArray[index] = Address6;
			index++;
		}
		
		if (addressArray.length <= 4)
		{
			for(counter in addressArray)
			{
				strAddress = strAddress + addressArray[counter] + '\\par ';
			}
		}
		else if (addressArray.length == 5)
		{
			strAddress = addressArray[0] + ', ' + addressArray[1] + '\\par ' + addressArray[2] + '\\par ' + addressArray[3] + '\\par ' + addressArray[4] + '\\par ' ;
		}
		else if (addressArray.length == 6)
		{
			strAddress = addressArray[0] + ', ' + addressArray[1] + '\\par ' + addressArray[2] + '\\par ' + addressArray[3] + '\\par ' + addressArray[4] + ', ' + addressArray[5] + '\\par ' ;
		}
		
		//MAH EP2_815
		strAddress = toPC(strAddress);
		
		if (FirmName != '')
		{
			strAddress =  FirmName + '\\par ' + strAddress;
		}
			
		if (PostCode != '')
		{
			strAddress = strAddress + PostCode;
		}
 
		return (strAddress);
	}
	
	function DealWithSalutation(Title, TitleOther, FirstName, Surname)
    {
		var strSalutation = "";
		
		   
		if ((Title != '') && (Surname !='')) 
		{
		   if ((TitleOther != '') && (Title == 'Other'))
		      strSalutation = strSalutation + TitleOther + ' ';
		   else 
				strSalutation = strSalutation + Title + ' ';
			
			strSalutation = strSalutation + toPC(Surname);
		}
		
		if (strSalutation == '')
		{
			strSalutation = 'Sir or Madam';
		}	 
			
		return (strSalutation);
    }

        function GetApplicantNameWithInitials(Title,TitleOther,FirstName,Surname)
        {
	var strName = "";
		if (Title == 'Other' )
		{
			Title = TitleOther;
		}
        if (Title != '')
		{
		strName = strName + Title + ' ';
		}
	    if (FirstName !='')
	    {
		strName = strName + toPC(FirstName.substring(0,1)) + ' ';
	    }
        if (Surname != '')
		{
		strName = strName + toPC(Surname);
		}
	return (strName);
        }    
    	    
	function GetDate(dteDate)
	{
		var today='';
		var blnNewDate;
		
		blnNewDate = false;
		if (typeof(dteDate)=='undefined')
		{
			blnNewDate = true;			
		}
		else
		{
			if(dteDate=='')
			{
				blnNewDate = true;
			}
		}
		
		if (blnNewDate)
		{
			dt = new Date();
			day = dt.getDate();
			year = dt.getFullYear();
			month = dt.getMonth();			
		}
		else
		{	        
			day = parseInt(dteDate.substring(0,2),10);
			month = parseInt(dteDate.substring(3,5),10)-1;
			year = parseInt(dteDate.substring(6,10),10);
		}	
  			
		if (month == '0')
		{
			month='January';
		}
		else if (month == '1')
		{
			month='February';
		}
		else if (month == '2')
		{
			month='March';
		}   
		else if (month == '3')
		{
			month='April';
		}
		else if (month == '4')
		{
			month='May';
		}
		else if (month == '5')
		{
			month='June';
		}
		else if (month == '6')
		{
			month='July';
		}
		else if (month == '7')
		{
			month='August';
		}
		else if (month == '8')
		{
			month='September';
		}
		else if (month == '9')
		{
			month='October';
		}
		else if (month == '10')
		{
			month='November';
		}
		else if (month == '11')
		{
			month='December';
		}

		/*
		if ((day == 1) || (day == 21) || (day == 31))
		{
			day = day + "st";
		}
		else if ((day == 2) || (day == 22))
		{
			day = day + "nd";
		}
		else if ((day == 3) || (day == 23))
		{
			day = day + "rd";
		}
		else
		{
			day = day + "th";
		}
		*/
   		
		today = day + ' ' + month + ' ' + year;
			
		return(today);
	}

	function GetAge(strDOB)
	{
			day = parseInt(strDOB.substring(0,2),10);
			month = parseInt(strDOB.substring(3,5),10) - 1;
			year = parseInt(strDOB.substring(6,10),10);
			
			dteDOB = new Date();
			dteDOB.setFullYear(year,month,day);
			
			dteNow = new Date();
			intYears = dteNow.getFullYear() - year;
			
			if((month > dteNow.getMonth()) || ((month == dteNow.getMonth()) && (day > dteNow.getDate())))
			{
				intYears = intYears - 1;
			}

			return(intYears);
	}
	
	function GetYears(strStart, strEnd)
	{
			startday = parseInt(strStart.substring(0,2),10);
			startmonth = parseInt(strStart.substring(3,5),10);
			startyear = parseInt(strStart.substring(6,10),10);
			
			dteStart = new Date();
			dteStart.setFullYear(startyear,startmonth,startday);

			if(strEnd=='')
			{
				dteEnd = new Date();

				endday = dteEnd.getDate();
				endmonth = dteEnd.getMonth() +1;
				endyear = dteEnd.getFullYear();
			}
			else
			{
				endday = parseInt(strEnd.substring(0,2),10);
				endmonth = parseInt(strEnd.substring(3,5),10);
				endyear = parseInt(strEnd.substring(6,10),10);
	
				dteEnd = new Date();
				dteEnd.setFullYear(endyear,endmonth,endday);
			}
			
			intYears = endyear - startyear 
			
			if((startmonth > endmonth) || ((startmonth == endmonth) && (startday > endday)))
			{
				intYears = intYears - 1;
			}

			return(intYears);
	}

	function GetMonths(strStart, strEnd)
	{
			startday = parseInt(strStart.substring(0,2),10);
			startmonth = parseInt(strStart.substring(3,5),10);
			startyear = parseInt(strStart.substring(6,10),10);
			
			dteStart = new Date();
			dteStart.setFullYear(startyear,startmonth,startday);

			if(strEnd=='')
			{
				dteEnd = new Date();

				endday = dteEnd.getDate();
				endmonth = dteEnd.getMonth() +1;
				endyear = dteEnd.getFullYear();
			}
			else
			{
				endday = parseInt(strEnd.substring(0,2),10);
				endmonth = parseInt(strEnd.substring(3,5),10);
				endyear = parseInt(strEnd.substring(6,10),10);
	
				dteEnd = new Date();
				dteEnd.setFullYear(endyear,endmonth,endday);
			}	
			
			if(startmonth > endmonth)
			{
				intMonths = (12-startmonth) + endmonth
			}
			else
			{
				intMonths = endmonth	- startmonth
			}			
			
			return(intMonths);
	}
	
	function DealWithTypeOfApplication(typeOfApplication)
	{
		var returnTypeOfApplication = "Other"
		var typesArray = typeOfApplication.split('|')
		for( ctr=0 ; ctr < typesArray.length; ctr++)
		{
			if (typesArray[ctr]  == "FT" || typesArray[ctr]  == "HM"  || typesArray[ctr]  == "R" )
			{
				returnTypeOfApplication = "Mortgage"
				break
			}
			
			else if (typesArray[ctr]  == "ABO" )
			{
				 if (typesArray[ctr+1]  == "ABTOE" )
				{
					returnTypeOfApplication = "Additional Borrowing with Transfer of Equity"
					break
				}
				else
				{
					returnTypeOfApplication = "Additional Borrowing"
					break
				}
			}
			
			else if (typesArray[ctr]  == "TOE" )
			{
				returnTypeOfApplication = "Transfer of Equity"
				break
			}
			
			else if (typesArray[ctr]  == "PSW" )
			{
				returnTypeOfApplication = "Product Switch"
				break
			}
			else if (typesArray[ctr]  == "CLI" )
			{
				returnTypeOfApplication = "Credit Limit Increase"
				break
			}
			else if (typesArray[ctr]  == "NP" )
			{
				returnTypeOfApplication = "Porting"
				break
			}
			else
			{
				returnTypeOfApplication = "Mortgage"
			}
		}
		return (returnTypeOfApplication)
	}
	
	function CheckTypeOfApplication(typeOfApplication)
	{
		var returnTypeOfApplication = "Other"
		var typesArray = typeOfApplication.split('|')
		for( ctr=0 ; ctr < typesArray.length; ctr++)
		{
			if (typesArray[ctr]  == "N" )
			{
				returnTypeOfApplication = "2"
				break
			}
			else if (typesArray[ctr]  == "ABO" )
			{
				returnTypeOfApplication = "1"
				break
			}
			else if (typesArray[ctr]  == "TOE" )
			{
				returnTypeOfApplication = "Transfer of Equity"
				break
			}
			else if (typesArray[ctr]  == "PSW" )
			{
				returnTypeOfApplication = "1"
				break
			}
			else if (typesArray[ctr]  == "CLI" )
			{
				returnTypeOfApplication = "1"
				break
			}
		}
		return (returnTypeOfApplication)
	}
	function toPC(v)
	{
		var s = " " + v.toLowerCase();
		var a;
		while(a = s.match(/ [a-z]|'[a-z]|-[a-z]|mc[a-z]|Mc[a-z]/))
		//while(a = s.match(/ [a-z]|mc[a-z]|Mc[a-z]/))
			s = s.substr(0,a.lastIndex -1) + s.substr(a.lastIndex -1,1).toUpperCase() + s.substring(a.lastIndex);
		s = s.replace(/ Bfpo /g," BFPO ");
		s = s.replace(/ Hms /g," HMS ");
		s = s.replace(/ Po /g," PO ");
		s = s.replace(/ Von /g," von ");
		s = s.replace(/ Van /g," van ");
		s = s.replace(/ De /g," de ");
				
		return(s.substring(1));
	}
	
	function GetApplicantNames(FirstName1, Surname1,  FirstName2, Surname2)
    {
		var strApplicantNames = "";
		
		if (FirstName1 != '' && Surname1 != '' )
		{
			
			strApplicantNames = FirstName1+ ' ' + Surname1;
			
			
		}
		
		if (FirstName2 != '' && Surname2 != '' )
		{
		
			strApplicantNames = strApplicantNames + ' and ' + FirstName2+ ' ' + Surname2;
			
		}
	
			
		return (strApplicantNames);
    }
	
    ]]></msxsl:script>
	<!--============================================================================================================-->
	<!-- NOTE: Requires import of document-functions-applicant.xsl for GetSingleOrJointSalutation function -->
	<xsl:template match="APPLICATION">
		<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"></xsl:value-of></xsl:attribute>
		<xsl:attribute name="APPLICATIONNUMBER">{<xsl:value-of select="./@APPLICATIONNUMBER"></xsl:value-of>}</xsl:attribute>
		<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string(./APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
		<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:GetSingleOrJointSalutation( 
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION/@TITLE_TEXT ), 
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION/@TITLEOTHER ),
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION/@FIRSTFORENAME ),
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=1]/CUSTOMERVERSION/@SURNAME ),
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION/@TITLE_TEXT ), 
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION/@TITLEOTHER ),
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION/@FIRSTFORENAME ),
																										string( ./CUSTOMERROLE[@CUSTOMERORDER=2]/CUSTOMERVERSION/@SURNAME )
																										)"/></xsl:attribute>
		<xsl:if test="./@INGESTIONAPPLICATIONNUMBER">
				<xsl:attribute name="PACKAGERCASEREF"><xsl:value-of select="concat('{',./@INGESTIONAPPLICATIONNUMBER,'}')"/></xsl:attribute>
				<xsl:element name="PACKAGERCASE"/>
			</xsl:if>
			
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="INTRODUCER">
		<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddressAndSalutation(msg:DealWithSalutation(
																										string( ./ORGANISATIONUSER/@USERTITLE_TEXT ),
																										'',
																										string( ./ORGANISATIONUSER/@USERFORENAME ),
																										string( ./ORGANISATIONUSER/@USERSURNAME )
																										),
																										./@FLATNUMBER,
																										./@BUILDINGORHOUSENAME,
																										./@BUILDINGORHOUSENUMBER,
																										./@STREET,
																										./@DISTRICT,
																										./@TOWN,
																										./@COUNTY,
																										./@POSTCODE
																										)"></xsl:value-of></xsl:attribute>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template name="APPLICANT">
		<xsl:param name="RESPONSE"/>
		<xsl:element name="APPLICANT">
			<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',$RESPONSE/CUSTOMERROLE[1]/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
			<xsl:if test="$RESPONSE/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES"> 
				<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string($RESPONSE/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
			</xsl:if> 
			<xsl:if test="../../../APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES">			
				<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string($RESPONSE/../../../APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
			</xsl:if>				
			<xsl:if test="../../APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES">			
				<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string($RESPONSE/../../APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
			</xsl:if>			
			<xsl:if test="$RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES">			
				<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string($RESPONSE/APPLICATION/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
			</xsl:if>	
			<xsl:if test="$RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER">
				<xsl:attribute name="PACKAGERCASEREF"><xsl:value-of select="concat('{',$RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER,'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="//ACCOUNT/@ACCOUNTNUMBER">
				<xsl:attribute name="LENDERREFERENCE"><xsl:value-of select="concat('{',//ACCOUNT/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$RESPONSE/APPLICATION/@APPLICATIONDATE">
				<xsl:attribute name="APPDATE"><xsl:value-of select="concat('{',msg:GetDate(string($RESPONSE/APPLICATION/@APPLICATIONDATE)),'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:variable name="Cust1">
				<xsl:value-of select="msg:GetApplicantNameWithInitials(
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERVERSION/@TITLE_TEXT),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERVERSION/@TITLEOTHER),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERVERSION/@FIRSTFORENAME),																																	    									
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERVERSION/@SURNAME))"/>
			</xsl:variable>
			<xsl:variable name="Cust2">
				<xsl:value-of select="msg:GetApplicantNameWithInitials(
					string($RESPONSE/CUSTOMERROLE[2]/CUSTOMERVERSION/@TITLE_TEXT),
					string($RESPONSE/CUSTOMERROLE[2]/CUSTOMERVERSION/@TITLEOTHER),
					string($RESPONSE/CUSTOMERROLE[2]/CUSTOMERVERSION/@FIRSTFORENAME),																																	    									
					string($RESPONSE/CUSTOMERROLE[2]/CUSTOMERVERSION/@SURNAME))"/>
			</xsl:variable>
			<xsl:variable name="CustCount"> 
				<xsl:value-of select="count($RESPONSE/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'])"/>
			</xsl:variable>
			<xsl:if test="$CustCount='1'">
				<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="$Cust1"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="$CustCount='2'">
				<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="concat($Cust1,' and ',$Cust2)"/></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@STREET),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@DISTRICT),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@TOWN),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@COUNTY),
					string($RESPONSE/CUSTOMERROLE[1]/CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/>
			</xsl:attribute>
			<xsl:if test="$RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER">
				<xsl:element name="PACKAGERCASE"/>
			</xsl:if>
		</xsl:element>
	</xsl:template>
	
	<!--============================================================================================================-->
	<xsl:template name="GUARANTOR">
		<xsl:param name="RESPONSE"/>
		<xsl:element name="GUARANTOR">
			<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE[1]/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
			<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
			<xsl:variable name="Cust1">
				<xsl:value-of select="msg:GetApplicantNameWithInitials(
					string($RESPONSE/CUSTOMERVERSION/@TITLE_TEXT),
					string($RESPONSE/CUSTOMERVERSION/@TITLEOTHER),
					string($RESPONSE/CUSTOMERVERSION/@FIRSTFORENAME),																																	    									
					string($RESPONSE/CUSTOMERVERSION/@SURNAME))"/>
			</xsl:variable>
			<xsl:attribute name="GUARANTORNAME"><xsl:value-of select="$Cust1"/></xsl:attribute>
			<xsl:if test="//ACCOUNT/@ACCOUNTNUMBER">
				<xsl:attribute name="LENDERREFERENCE"><xsl:value-of select="concat('{',//ACCOUNT/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="GUARANTORADDRESS"><xsl:value-of select="msg:DealWithAddress(
					string($RESPONSE/CUSTOMERVERSION/@TITLE_TEXT),
					string($RESPONSE/CUSTOMERVERSION/@TITLEOTHER),
					string($RESPONSE/CUSTOMERVERSION/@FIRSTFORENAME),
					string($RESPONSE/CUSTOMERVERSION/@SURNAME),	
					'',
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@STREET),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@DISTRICT),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@TOWN),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@COUNTY),
					string($RESPONSE/CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
			<xsl:if test="$RESPONSE/APPLICATION/@INGESTIONAPPLICATIONNUMBER">
				<xsl:element name="PACKAGERCASE"/>
			</xsl:if>
			<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(
				string($RESPONSE/CUSTOMERVERSION/@TITLE_TEXT),
				string($RESPONSE/CUSTOMERVERSION/@TITLEOTHER),
				string($RESPONSE/CUSTOMERVERSION/@FIRSTFORENAME),
				string($RESPONSE/CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
			<xsl:choose>
				<xsl:when test="string($RESPONSE/CUSTOMERVERSION/@TITLE_TEXT) != '' and string($RESPONSE/CUSTOMERVERSION/@SURNAME)  != ''">
					<xsl:element name="SINCERELY"/>
				</xsl:when>
			<xsl:otherwise>
				<xsl:element name="FAITHFULLY"/>
			</xsl:otherwise>
		</xsl:choose>
		</xsl:element>
	</xsl:template>
	
	<!--============================================================================================================-->
	<xsl:template match="CUSTOMERVERSION">
		<!--The root node is set to APPLICANT or GUARANTOR, based on CustomerRoleType-->
		<xsl:variable name="RootElement">
			<xsl:choose>
				<xsl:when test="../@CUSTOMERROLETYPE_TYPES='G'">
					<xsl:value-of select="string('GUARANTOR')"></xsl:value-of>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="string('APPLICANT')"></xsl:value-of>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:element name="{$RootElement}">
			<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',../@APPLICATIONNUMBER,'}')"/></xsl:attribute>
			<xsl:attribute name="TYPEOFAPPLICATION"><xsl:value-of select="msg:DealWithTypeOfApplication(string(../../APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:attribute>
			<xsl:if test="../APPLICATION/@INGESTIONAPPLICATIONNUMBER">
				<xsl:attribute name="PACKAGERCASEREF"><xsl:value-of select="concat('{',../APPLICATION/@INGESTIONAPPLICATIONNUMBER,'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="../../APPLICATIONFACTFIND/RISKASSESSMENT/@RISKASSESSMENTDATETIME">
				<xsl:attribute name="AIPDATE"><xsl:value-of select="concat('{',msg:GetDate(string(../../APPLICATIONFACTFIND/RISKASSESSMENT/@RISKASSESSMENTDATETIME)),'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="../../APPLICATION/@APPLICATIONDATE">
				<xsl:attribute name="APPDATE"><xsl:value-of select="concat('{',msg:GetDate(string(../../APPLICATION/@APPLICATIONDATE)),'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="../../APPLICATIONFACTFIND/ACTIVEQUOTE/MORTGAGESUBQUOTE/LOANCOMPONENT/@NETMONTHLYCOST">
				<xsl:attribute name="MONTHLYPAYMENT"><xsl:value-of select="../../APPLICATIONFACTFIND/ACTIVEQUOTE/MORTGAGESUBQUOTE/LOANCOMPONENT/@NETMONTHLYCOST"/></xsl:attribute>
			</xsl:if>
			<xsl:if test="ACCOUNT/@ACCOUNTNUMBER">
				<xsl:attribute name="LENDERREFERENCE"><xsl:value-of select="concat('{',ACCOUNT/@ACCOUNTNUMBER,'}')"/></xsl:attribute>
			</xsl:if>
			<xsl:attribute name="APPLICANTNAME"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(@TITLE_TEXT),
					string(@TITLEOTHER),
					string(@FIRSTFORENAME),																																	    									
					string(@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="APPLICANTNAMES"><xsl:value-of select="msg:GetApplicantNames(string(@FIRSTFORENAME),																																	    									
					string(@SURNAME),'','')"/></xsl:attribute>
			<xsl:attribute name="APPLICANTADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
					string(../CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
					string(../CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
					string(../CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
					string(../CUSTOMERADDRESS/ADDRESS/@STREET),
					string(../CUSTOMERADDRESS/ADDRESS/@DISTRICT),
					string(../CUSTOMERADDRESS/ADDRESS/@TOWN),
					string(../CUSTOMERADDRESS/ADDRESS/@COUNTY),
					string(../CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
		</xsl:element>
		
		
		<xsl:if test="TENANCY">
			<xsl:element name="LANDLORDDETAILS">
				<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
				<xsl:apply-templates select="TENANCY[1]/THIRDPARTY/ADDRESS"/>
				<xsl:apply-templates select="TENANCY[1]/THIRDPARTY/CONTACTDETAILS"/>
				<xsl:apply-templates select="TENANCY[1]/NAMEANDADDRESSDIRECTORY/ADDRESS"/>
				<xsl:apply-templates select="TENANCY[1]/NAMEANDADDRESSDIRECTORY/CONTACTDETAILS"/>
			</xsl:element>
		</xsl:if>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="ADDRESS">
		<xsl:choose>
			<xsl:when test="../../INTERMEDIARYORGANISATION">
				<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(
						string(../CONTACTDETAILS/@CONTACTTITLE_TEXT),
						string(../CONTACTDETAILS/@CONTACTTITLEOTHER),
						string(../CONTACTDETAILS/@CONTACTFORENAME),
						string(../CONTACTDETAILS/@CONTACTSURNAME),	
						string(../../INTERMEDIARYORGANISATION/@NAME),									
						string(@FLATNUMBER),
						string(@BUILDINGORHOUSENAME),
						string(@BUILDINGORHOUSENUMBER),
						string(@STREET),
						string(@DISTRICT),
						string(@TOWN),
						string(@COUNTY),
						string(@POSTCODE))"/></xsl:attribute>
			</xsl:when>
			<xsl:otherwise>
				<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(
						string(../CONTACTDETAILS/@CONTACTTITLE_TEXT),
						string(../CONTACTDETAILS/@CONTACTTITLEOTHER),
						string(../CONTACTDETAILS/@CONTACTFORENAME),
						string(../CONTACTDETAILS/@CONTACTSURNAME),	
						string(../@COMPANYNAME),									
						string(@FLATNUMBER),
						string(@BUILDINGORHOUSENAME),
						string(@BUILDINGORHOUSENUMBER),
						string(@STREET),
						string(@DISTRICT),
						string(@TOWN),
						string(@COUNTY),
						string(@POSTCODE))"/></xsl:attribute>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="ARFIRM">
		<xsl:variable name="AFIRMNAME">
			<xsl:choose>
				<xsl:when test="ARFIRMALTNAME/@ALTERNATIVENAME">
					<xsl:value-of select="ARFIRMALTNAME/@ALTERNATIVENAME"></xsl:value-of>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="@ARFIRMNAME"></xsl:value-of>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:attribute name="ARINTRONAME"><xsl:value-of select="$AFIRMNAME"></xsl:value-of></xsl:attribute>
		<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithFSAAddress(
				string($AFIRMNAME),									
				string(@ADDRESSLINE1),
				string(@ADDRESSLINE2),
				string(@ADDRESSLINE3),
				string(@ADDRESSLINE4),
				string(@ADDRESSLINE5),
				string(@ADDRESSLINE6),
				string(@POSTCODE))"/>
		</xsl:attribute>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="PRINCIPALFIRM">
		<xsl:variable name="PFIRMNAME">
			<xsl:choose>
				<xsl:when test="PRINCIPALFIRMALTNAME/@ALTERNATIVENAME">
					<xsl:value-of select="PRINCIPALFIRMALTNAME/@ALTERNATIVENAME"></xsl:value-of>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="@PRINCIPALFIRMNAME"></xsl:value-of>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:attribute name="PINTRONAME"><xsl:value-of select="$PFIRMNAME"></xsl:value-of></xsl:attribute>
		<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithFSAAddress(
				string($PFIRMNAME),									
				string(@ADDRESSLINE1),
				string(@ADDRESSLINE2),
				string(@ADDRESSLINE3),
				string(@ADDRESSLINE4),
				string(@ADDRESSLINE5),
				string(@ADDRESSLINE6),
				string(@POSTCODE))"/>
		</xsl:attribute>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template match="CONTACTDETAILS">
		<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(
				string(@CONTACTTITLE_TEXT),
				string(@CONTACTTITLEOTHER),
				string(@CONTACTFORENAME),
				string(@CONTACTSURNAME))"/></xsl:attribute>
		<xsl:choose>
			<xsl:when test="string(@CONTACTTITLE_TEXT) != '' and string(@CONTACTSURNAME) !=''">
				<xsl:element name="SINCERELY"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:element name="FAITHFULLY"/>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	<!--============================================================================================================-->
	<xsl:template name="APPLICATIONINTRODUCER">
	<xsl:choose>
	<xsl:when test="//APPLICATIONBROKER/@INTRODUCERID">
		<xsl:variable name="FIRMNAME">
			<xsl:choose>
				<xsl:when test="//APPLICATIONBROKER/@ARFIRMID">
					<xsl:value-of select="//ARFIRM/@ARFIRMNAME"></xsl:value-of>
				</xsl:when>
				<xsl:when test="//APPLICATIONBROKER/@PRINCIPALFIRMID">
					<xsl:value-of select="//ARFIRM/@PRINCIPALFIRMNAME"></xsl:value-of>
				</xsl:when>
				</xsl:choose>
		</xsl:variable>
		<xsl:attribute name="ADDRESS"><xsl:value-of select="msg:DealWithAddress(
					'',
					'',
					'',
					'',
					string($FIRMNAME),	
					string(APPLICATIONBROKER/@FLATNUMBER),
					string(APPLICATIONBROKER/@BUILDINGORHOUSENAME),
					string(APPLICATIONBROKER/@BUILDINGORHOUSENUMBER),
					string(APPLICATIONBROKER/@STREET),
					string(APPLICATIONBROKER/@DISTRICT),
					string(APPLICATIONBROKER/@TOWN),
					string(APPLICATIONBROKER/@COUNTY),
					string(APPLICATIONBROKER/@POSTCODE))"/></xsl:attribute>
		<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(
				string(APPLICATIONBROKER/@TITLE),
				'',
				string(APPLICATIONBROKER/@USERFORENAME),
				string(APPLICATIONBROKER/@USERSURNAME))"/></xsl:attribute>
		<xsl:attribute name="APPINTRODUCERCONTACT"><xsl:value-of select="msg:GetApplicantNameWithInitials(
			string(APPLICATIONBROKER/@TITLE),
			'',
			string(APPLICATIONBROKER/@USERFORENAME),
			string(APPLICATIONBROKER/@USERSURNAME))"/>
		</xsl:attribute>				
		<xsl:choose>
			<xsl:when test="string(APPLICATIONBROKER/@TITLE) != '' and string(APPLICATIONBROKER/@USERSURNAME) !=''">
				<xsl:element name="SINCERELY"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:element name="FAITHFULLY"/>
			</xsl:otherwise>
		</xsl:choose>
		</xsl:when>
		
		<xsl:otherwise>
				<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:DealWithSalutation(
				string(APPLICATIONPACKAGER/@TITLE),
				'',
				string(APPLICATIONPACKAGER/@USERFORENAME),
				string(APPLICATIONPACKAGER/@USERSURNAME))"/></xsl:attribute>
				<xsl:attribute name="APPINTRODUCERCONTACT"><xsl:value-of select="msg:GetApplicantNameWithInitials(
					string(APPLICATIONPACKAGER/@TITLE),
					'',
					string(APPLICATIONPACKAGER/@USERFORENAME),
					string(APPLICATIONPACKAGER/@USERSURNAME))"/>
				</xsl:attribute>
		<xsl:choose>
			<xsl:when test="string(APPLICATIONPACKAGER/@TITLE) != '' and string(APPLICATIONPACKAGER/@USERSURNAME) !=''">
				<xsl:element name="SINCERELY"/>
			</xsl:when>
			<xsl:otherwise>
				<xsl:element name="FAITHFULLY"/>
			</xsl:otherwise>
		</xsl:choose>
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<!--============================================================================================================-->
</xsl:stylesheet>
 