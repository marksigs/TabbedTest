<?xml version="1.0" encoding="UTF-8"?>
<?altova_samplexml C:\Program Files\Marlborough Stirling\Omiga 4\XML\omResponse.xml?>
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
    

        function DealWithAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
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
		if (PostCode != '')
		{
		strAddress = strAddress +  PostCode;
		}		
	return (strAddress);
        }


function DealWithSolicitorAddress(CompanyName,Title,FirstName,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
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

	
	function GetSexDescription(Sex)
	{
			var strSex='';
	
			if (Sex == '1')
				strSex='Sir';
			else if (Sex == '2')
				strSex='Madam';
			else
				strSex='Sirs';
			return (strSex);			
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
    
		var titlename ='';	    
	    function SetSalutation(Title,FirstName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					titlename= titlename  + Title ;
				if (FirstName !='')
					titlename=titlename + ' ' + FirstName;	
				if (Surname !='')
					titlename=titlename + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					titlename= titlename  + ' and ' + Title ;
				if (FirstName !='')
					titlename=titlename + ' ' + FirstName;	
				if (Surname !='')
					titlename=titlename + ' ' + Surname ;
			}	
		return(titlename);
	    }
        
		function GetSalutation()
		{
			return(titlename);
		}     



		var strGender ='';	    
	    function GetInformation(Gender)
	    {
			if ((Gender!='F') || (Gender!='f'))
			{
				strGender='Madam';
			}

			if ((Gender!='M') || (Gender!='m')) 
			{
				strGender='Sir';
			}		
		return(strGender);
        }
     
     
     
    ]]></msxsl:script>
    <xsl:variable name="COAPPLICANT"/> 
    <xsl:variable name="Order_No"/>
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustRole" select="count(RESPONSE/SOLICITOROFFERLETTER/CUSTOMERROLE)"/>
						<xsl:for-each select="//CUSTOMERROLE">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<xsl:variable name="COAPPLICANT" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
																																			   string(.//CUSTOMERVERSION/@SURNAME),1)"/>
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="COAPPLICANT" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
																																			   string(.//CUSTOMERVERSION/@SURNAME),2)"/>
							</xsl:if>
						</xsl:for-each>	
				<xsl:element name="INFORMATION">
					<xsl:element name="SOLICITORDETAILS">
						    <xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:DealWithSolicitorAddress(
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/@COMPANYNAME),
																												  string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE_TEXT),
																												  string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTFORENAME),
																												  string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTSURNAME),	
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@FLATNUMBER),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@STREET),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@DISTRICT),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@TOWN),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@COUNTY),
																												  string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@POSTCODE))"/></xsl:attribute>
							<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
							<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithAddress(
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@STREET),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@DISTRICT),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@TOWN),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
																																					string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
							<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
							<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@CUSTOMERNUMBER,'}')"/></xsl:attribute>							
							<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>
							<xsl:attribute name="SALUTATION">
								<xsl:value-of select="msg:GetInformation(string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE_TYPES))"/>
							</xsl:attribute>							
							<xsl:variable name="Temp" select="msg:UpdateCounter(1)" />
							<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
							<xsl:if test="($TempCounter > 1)">
								<xsl:element name="PAGEBREAK"/>
							</xsl:if>
						</xsl:element>
				</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
