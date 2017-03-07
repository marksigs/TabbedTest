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



        function DealWithNewAddress(FlatNumber, Street, Town, County, PostCode)
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
		 strAddress = strAddress + County + ', ';
		}
	    if (PostCode != '')
		{
		 strAddress = strAddress + PostCode ;
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


function GetBothNames(Title,Surname)
{
	var strAddress = "";
	
		 if (Title != '')
		{
		strAddress = strAddress +  Title + ' ' ;
		}

		 if (Surname != '')
		{
				strAddress = strAddress +  Surname + ' ' ;
		}
		return (strAddress);
}
  
		var strCustomerNo ='';	    
	    function SetCustomerNo(CustomerNo,Counter)
	    {
			if (Counter == '1') 
			{
				if (CustomerNo !='')
					strCustomerNo= strCustomerNo  + CustomerNo ;
			}
		  else if (Counter == '2') 
			{
				if (CustomerNo !='')
					strCustomerNo= strCustomerNo  + ' and ' + CustomerNo ;
			}	
		return(strCustomerNo);
	    }
        
		function GetCustomerNo()
		{
			return(strCustomerNo);
		}
  

		var strValue='';
		function TypeOfMortgage(ValueId,Desc)  
		{
			if (ValueId == '10')
			{
				strValue='Purchase';
			}
			if (ValueId == '20')
			{
				strValue=Desc;
			}				
		return(strValue);
		}
  
  
    ]]></msxsl:script>
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="CustomerNo"/>
    <xsl:variable name="Order_No"/> 
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustCount" select="count(RESPONSE/ORIGINALDOCUMENTS/UNIQUEAPPLICATIONADDRESSES)"/>
			<xsl:variable name="CustRole" select="count(RESPONSE/ORIGINALDOCUMENTS/CUSTOMERROLE)"/>
			<xsl:if test="($CustCount = '1')">
				<xsl:if test="($CustRole = '2')">
						<xsl:for-each select="//CUSTOMERROLE">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),1)"/>
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@SURNAME),2)"/>
									<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),2)"/>
							</xsl:if>
						</xsl:for-each>	
						<xsl:element name="INFORMATION">
								<xsl:element name="CUSTOMERDETAILS">
									<xsl:attribute name="TITLES"><xsl:value-of select="msg:GetSalutation()"/> </xsl:attribute>						
									<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:DealWithAddress(
												string(msg:GetSalutation()),										
												string(''),
												string(''),
												string(.//CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
												string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
												string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
												string(.//CUSTOMERADDRESS/ADDRESS/@STREET),
												string(.//CUSTOMERADDRESS/ADDRESS/@DISTRICT),
												string(.//CUSTOMERADDRESS/ADDRESS/@TOWN),
												string(.//CUSTOMERADDRESS/ADDRESS/@COUNTY),
												string(.//CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
									<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>							
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="msg:GetCustomerNo()"/></xsl:attribute>
									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
									<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
												string(.//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
												string(.//NEWPROPERTYADDRESS/ADDRESS/@STREET),
												string(.//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
												string(.//NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
												string(.//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<xsl:attribute name="TYPEOFMORTGAGE">
										<xsl:value-of select="msg:TypeOfMortgage(string(//APPLICATIONFACTFIND/COMBOVALUE/@VALUEID),
																											   string(//APPLICATIONFACTFIND/COMBOVALUE/@VALUENAME))"/>
									</xsl:attribute>								
								</xsl:element>
						</xsl:element>
				</xsl:if>		
			</xsl:if>
			<xsl:if test="(($CustCount = '1' and $CustRole = '1') or ($CustCount = '2' and $CustRole = '2'))">
				<xsl:for-each select="//CUSTOMERROLE">
					<xsl:element name="INFORMATION">
							<xsl:element name="CUSTOMERDETAILS">
								<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:GetAddress(
											string(.//CUSTOMERVERSION/@TITLE_TEXT),
											string(.//CUSTOMERVERSION/@FIRSTFORENAME),
											string(.//CUSTOMERVERSION/@SURNAME),
											string(.//CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
											string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
											string(.//CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
											string(.//CUSTOMERADDRESS/ADDRESS/@STREET),
											string(.//CUSTOMERADDRESS/ADDRESS/@DISTRICT),
											string(.//CUSTOMERADDRESS/ADDRESS/@TOWN),
											string(.//CUSTOMERADDRESS/ADDRESS/@COUNTY),
											string(.//CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
								<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
								<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetBothNames(
																																string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																string(.//CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
								<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}')"/></xsl:attribute>
								<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
								<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
								<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
								<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
											string(.//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
											string(.//NEWPROPERTYADDRESS/ADDRESS/@STREET),
											string(.//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
											string(.//NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
											string(.//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
								<xsl:attribute name="TYPEOFMORTGAGE">
									<xsl:value-of select="msg:TypeOfMortgage(string(//APPLICATIONFACTFIND/COMBOVALUE/@VALUEID),
																										   string(//APPLICATIONFACTFIND/COMBOVALUE/@VALUENAME))"/>
								</xsl:attribute>								
								<xsl:if test="($TempCounter > 1)">
									<xsl:element name="PAGEBREAK"/>
								</xsl:if>			
							</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:if>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
