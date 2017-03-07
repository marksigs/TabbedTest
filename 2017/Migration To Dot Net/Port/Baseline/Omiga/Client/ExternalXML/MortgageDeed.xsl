<?xml version="1.0" encoding="UTF-8"?>
<!-- edited with XMLSpy v2005 rel. 3 U (http://www.altova.com) by ﻿ITS (Marlborough Stirling plc) -->
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
        
          
         var CustCount = new Array(0); 
        function Count(WhichCounter)
        {
            CustCount[WhichCounter] = CustCount[WhichCounter] + 1;
            return (CustCount[WhichCounter]);
        }          
            


		var titlename ='';	    
	    function SetApplicantName(Title,FirstName,SecondName,OtherName,Surname)
	    {
				if (Title !='')
					titlename= titlename  + Title ;
				if (FirstName !='')
					titlename=titlename + ' ' + FirstName;
				if (SecondName !='')
					titlename=titlename + ' ' + SecondName;					
				if (OtherName !='')
					titlename=titlename + ' ' + OtherName;					
				if (Surname !='')
					titlename=titlename + ' ' + Surname;
		return(titlename);
	    }
        
		function GetApplicantName()
		{
			return(titlename);
		}

		var strtitlename ='';	    
	    function SetCoApplicantName(Title,FirstName,SecondName,OtherName,Surname)
	    {
				if (Title !='')
					strtitlename= strtitlename  + Title ;
				if (FirstName !='')
					strtitlename=strtitlename + ' ' + FirstName;
				if (SecondName !='')
					strtitlename=strtitlename + ' ' + SecondName;					
				if (OtherName !='')
					strtitlename=strtitlename + ' ' + OtherName;					
				if (Surname !='')
					strtitlename=strtitlename + ' ' + Surname;
		return(strtitlename);
	    }
        
		function GetCoApplicantName()
		{
			return(strtitlename);
		}

        function DealWithAddress(SalutationOne,Title,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	

		 if (Title != '')
		{
		strAddress = strAddress +  Title + ' ' ;
		}

		 if (Surname != '')
		{
			 if (SalutationOne != '')
				strAddress = strAddress +  Surname + ' ' ;
			else
				strAddress = strAddress +  Surname + ' \\par' ;
		}
		
		 if (SalutationOne != '')
		{
		strAddress = strAddress +  SalutationOne + ' \\par ';
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


        function GetAddress(Title,FirstName,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
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



        function DealWithNewAddress(FlatNumber, Street, Town, County)
        {
	var strAddress = "";
	
	    if (FlatNumber != '')
		{
		 strAddress = strAddress + FlatNumber + ', ';
		}
	    if (Street != '')
		{
		 strAddress = strAddress + Street + ', ';
		}
	    if (Town != '')
		{
		 strAddress = strAddress + Town + ', ';
		}
	    if (County != '')
		{
		 strAddress = strAddress + County + ' ';
		}
	return (strAddress);
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
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="Applicant"/>
    <xsl:variable name="CoApplicant"/>
    <xsl:variable name="Order_No"/> 
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustCount" select="count(RESPONSE/MORTGAGEDEED/UNIQUEAPPLICATIONADDRESSES)"/>
			<xsl:variable name="CustRole" select="count(RESPONSE/MORTGAGEDEED/CUSTOMERROLE)"/>
				<xsl:for-each select="//CUSTOMERROLE">
					<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
					<xsl:if test="($Order_No = '1')">
							<xsl:variable name="Applicant" select="msg:SetApplicantName(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																														string(.//CUSTOMERVERSION/@FIRSTFORENAME),									 
																														string(.//CUSTOMERVERSION/@SECONDFORENAME),
																														string(.//CUSTOMERVERSION/@OTHERFORENAMES),
																														string(.//CUSTOMERVERSION/@SURNAME))"/>
					</xsl:if>		
					<xsl:if test="($Order_No = '2')">
							<xsl:variable name="CoApplicant" select="msg:SetCoApplicantName(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																														string(.//CUSTOMERVERSION/@FIRSTFORENAME),									 
																														string(.//CUSTOMERVERSION/@SECONDFORENAME),
																														string(.//CUSTOMERVERSION/@OTHERFORENAMES),
																														string(.//CUSTOMERVERSION/@SURNAME))"/>
					</xsl:if>
				</xsl:for-each>
				<xsl:attribute name="HMLAND"><xsl:value-of select="string('[                          ]    .')"/></xsl:attribute>	
				<xsl:element name="INFORMATION">
						<xsl:element name="CUSTOMERDETAILS">
							<xsl:attribute name="APPLICANT"><xsl:value-of select="msg:GetApplicantName()"/> </xsl:attribute>						
							<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetCoApplicantName()"/> </xsl:attribute>
							<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
							<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
										string(.//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
										string(.//NEWPROPERTYADDRESS/ADDRESS/@STREET),
										string(.//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
										string(.//NEWPROPERTYADDRESS/ADDRESS/@COUNTY))"/></xsl:attribute>
							<xsl:attribute name="POSTCODE"><xsl:value-of select="string(.//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE)"/></xsl:attribute>			
							<xsl:attribute name="HMLAND"><xsl:value-of select="string('[                          ]    .')"/></xsl:attribute>
							
						</xsl:element>
				</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
