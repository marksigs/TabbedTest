<?xml version="1.0" encoding="UTF-8"?>
<!-- edited with XMLSpy v2005 rel. 3 U (http://www.altova.com) by ITS (Marlborough Stirling plc) -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<xsl:output method="xml" indent="yes" encoding="UTF-8"/>
	<msxsl:script language="JavaScript" implements-prefix="msg"><![CDATA[
    

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


		var titlename ='';	    
	    function SetSalutation(Title,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					titlename= titlename  + Title ;
				if (Surname !='')
					titlename=titlename + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					titlename= titlename  + ' and ' + Title ;
				if (Surname !='')
					titlename=titlename + ' ' + Surname ;
			}	
		return(titlename);
	    }
        
		function GetSalutation()
		{
			return(titlename);
		}

        function DealWithAddress(Title,FirstName,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	
		 if (Title != '')
		{
		strAddress = strAddress + Title + ' ';
		}	 
	    if ((FirstName !='') || (Surname !='')) 
		{
		    if (FirstName !='')
		    {
			strAddress = strAddress + FirstName.substring(0,1);
		    }
		strAddress = strAddress +  ' ' + Surname + '\\par ' ;
		}
		
		if ((Title == '') && (FirstName == '')) 
		{
		strAddress = Surname + '\\par ' ;
		}
        if (FlatNumber != '')
		{
		strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
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
		if (PostCode != '')
		{
		strAddress = strAddress +  PostCode;
		}		
	return (strAddress);
        }


function DealWithLandlordAddress(CompanyName,Title,FirstName,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	
	    if (CompanyName != '')
		{
		strAddress = strAddress + CompanyName + '\\par ';
		}
		 if (Title != '')
		{
		strAddress = strAddress + Title + ' ';
		}	 
	    if ((FirstName !='') || (Surname !='')) 
		{
		    if (FirstName !='')
		    {
			strAddress = strAddress + FirstName.substring(0,1);
		    }
		strAddress = strAddress +  ' ' + Surname + '\\par ' ;
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
		if (PostCode != '')
		{
		strAddress = strAddress + '\\par ' + PostCode;
		}		
	return (strAddress);
        }



        function GetApplicantNameWithInitials(Title,FirstName,Surname)
        {
	var strName = "";
        if (Title != '')
		{
		strName = strName + Title + ' ';
		}
	    if (FirstName !='')
	    {
		strName = strName + FirstName + ' ';
	    }
        if (Surname != '')
		{
		strName = strName + Surname;
		}
	return (strName);
        }    
    
	
	
	function GetGenderDescription(Gender)
	{
			var strGender='';
	
			if (Gender !='')			
			{
				if (Gender == '1' ) 
				{
					strGender='his';
				}	
				else if (Gender == '2')
				{
					strGender='her';
				}	
				else	
				{
					strGender='their';
				}	
			}		
		
		    return (strGender);	
	}
    
function GetDate()
{
	var today='';
   dt = new Date();        
   day = dt.getDate();
   year = dt.getFullYear();
   month = dt.getMonth();
		
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
    		
   today= day + ' ' + month + ' ' + year;
   return(today);
}    
    
    
    ]]></msxsl:script>
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:element name="INFORMATION">
				<xsl:element name="CUSTOMERDETAILS">
					<xsl:variable name="MultiApplicant">
						<xsl:value-of select="count(//TEMPLATEDATA/ADDRESS)"/>
					</xsl:variable>

					<xsl:attribute name="HASORHAVE">
						<xsl:value-of select="string('has')"/>
					</xsl:attribute>

					<xsl:attribute name="LANDLORDADDRESS">
						<xsl:value-of select="msg:DealWithLandlordAddress(
							string(.//TENANCY/THIRDPARTY/@COMPANYNAME),
							string(.//CONTACTDETAILS/@CONTACTTITLE_TEXT),
							string(.//CONTACTDETAILS/@CONTACTFORENAME),
							string(.//CONTACTDETAILS/@CONTACTSURNAME),	
							string(.//THIRDPARTY/ADDRESS/@FLATNUMBER),
							string(.//THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME),
							string(.//THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER),
							string(.//THIRDPARTY/ADDRESS/@STREET),
							string(.//THIRDPARTY/ADDRESS/@DISTRICT),
							string(.//THIRDPARTY/ADDRESS/@TOWN),
							string(.//THIRDPARTY/ADDRESS/@COUNTY),
							string(.//THIRDPARTY/ADDRESS/@POSTCODE))"/>
					</xsl:attribute>

					<xsl:attribute name="CURRENTDATE">
						<xsl:value-of select="msg:GetDate()"/>
					</xsl:attribute>	
																				 
					<xsl:attribute name="APPLICANTADDRESS">
						<xsl:value-of select="msg:DealWithAddress(
							string(''),
							string(''),
							string(msg:GetSalutation()),
							string(.//CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
							string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
							string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
							string(.//CUSTOMERADDRESS/ADDRESS/@STREET),
							string(.//CUSTOMERADDRESS/ADDRESS/@DISTRICT),
							string(.//CUSTOMERADDRESS/ADDRESS/@TOWN),
							string(.//CUSTOMERADDRESS/ADDRESS/@COUNTY),
							string(.//CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/>
					</xsl:attribute>

					<xsl:attribute name="APPLICATIONNUMBER">
						<xsl:value-of select="concat('{',//APPLICATIONFACTFIND/@APPLICATIONNUMBER,'}')"/>
					</xsl:attribute>

       	    				<xsl:attribute name="APPLICANT">
               					<xsl:for-each select="//CUSTOMERVERSION">
                    					<xsl:variable name="AppCounter" select="msg:UpdateCounter(2)"/>
							<xsl:variable name="Applicant"/>
       	            					<xsl:variable name="Title" select="@TITLE_TEXT"/>
               	    					<xsl:variable name="Initial1" select="@FIRSTFORENAME"/>
                    					<xsl:variable name="Initial2" select="@SECONDFORENAME"/>
                    					<xsl:variable name="Surname" select="@SURNAME"/>

                    					<xsl:if test="($AppCounter != 1)">
                        					<xsl:value-of select="concat($Applicant, ' and ')"/>
                    					</xsl:if>

                    					<xsl:if test="($Title != '')">
                        					<xsl:value-of select="concat($Applicant, $Title, ' ')"/>
                    					</xsl:if>

                    					<xsl:if test="($Initial1 != '')">
                        					<xsl:value-of select="concat($Applicant, substring($Initial1, 1, 1), ' ')"/>
                    					</xsl:if>

                    					<xsl:if test="($Initial2 != '')">
                        					<xsl:value-of select="concat($Applicant, substring($Initial2, 1, 1), ' ')"/>
                   					</xsl:if>

                    					<xsl:if test="($Surname != '')">
                        					<xsl:value-of select="concat($Applicant, $Surname)"/>
                    					</xsl:if>

                				</xsl:for-each>
           				 </xsl:attribute>


					<xsl:attribute name="GENDER">
						<xsl:value-of select="msg:GetGenderDescription(string(.//COMBOVALUE/@VALUEID))"/>
					</xsl:attribute>

				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
