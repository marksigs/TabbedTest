<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
<!-- PB 13/06/2006 EP721	Not using correspondence address -->
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
		var Applicant1Title='';
        var App1Surname='';
        var App1Title='';
        var App1Initial= '';
		var Applicant2Title='';
		var A

	/*  PM 28/6/2006 EP890 Removed original SetSalutation function that only had 1 title parameter */
	    function SetSalutation(TitleType, TitleText, TitleOther, FirstName, Surname, Counter)
	    {
			if (Counter == '1') 
			{
				if (TitleType == 'O')
				{
					if (TitleOther != '')
					{
						titlename = titlename + TitleOther;
						Applicant1Title = Applicant1Title + TitleOther;
						App1Title = TitleOther						
					}
				}
				else
				{
					if (TitleText != '')
					{
						titlename = titlename + TitleText;
						Applicant1Title = Applicant1Title + TitleText;
						App1Title = TitleText
					}
				}
				
				if (FirstName !='')
				{		
					Applicant1Title= Applicant1Title + ' ' + FirstName.substring(0,1);
					App1Initial= FirstName.substring(0,1);
				}

				if (Surname !='')
				{
					titlename=titlename + ' ' + Surname;
					Applicant1Title= Applicant1Title + ' ' + Surname;
					App1Surname = Surname; 
				}
			}
		  else if (Counter == '2') 
			{
				if ((Surname == App1Surname) && (TitleText== 'Mrs'))
				{
					titlename= App1Title + ' & ' + TitleText;
					titlename=titlename + ' ' + Surname ;
					Applicant1Title= App1Title + ' & ' + TitleText;
					Applicant1Title= Applicant1Title + ' ' + App1Initial +  ' ' + Surname ;
				}
				else
				{				
					if (TitleType == 'O')
					{
						if (TitleOther != '')
						{
							titlename = titlename + ' & ' + TitleOther;
							Applicant2Title = Applicant2Title + TitleOther;														
						}
					}
					else
					{
						if (TitleText != '')
						{
							titlename = titlename + ' & ' + TitleText;
							Applicant2Title = Applicant2Title + TitleText;														
						}
					}
					
					if (FirstName !='')
					{
						//titlename= titlename + ' ' + FirstName.substring(0,1);
						Applicant2Title= Applicant2Title + ' ' + FirstName.substring(0,1);
					}
					if (Surname !='')
					{
						titlename=titlename + ' ' + Surname ;
						Applicant2Title= Applicant2Title + ' ' + Surname;					
					}
				}
			}	
		return(titlename);
	    }


		function GetSalutation()
		{
			return(titlename);
		}




		var strApplicant1Title='';
		var strApplicant2Title='';

	/*  PM 28/6/2006 EP890 Removed original GetApplicantNameWithInitials function that only had 1 title parameter */
        function GetApplicantNameWithInitials(TitleType, TitleText, TitleOther, FirstName, SecondName, Surname, Counter)
        {
			if (Counter == '1')	
			{
				if (TitleType == 'O')
				{
					if (TitleOther != '')
					{
						strApplicant1Title = strApplicant1Title + TitleOther + ' ';
					}
				}
				else
				{
					if (TitleText != '')
					{
						strApplicant1Title = strApplicant1Title + TitleText + ' ';
					}
				}

				if (FirstName !='')
				{
				strApplicant1Title = strApplicant1Title + FirstName + ' ';
				}
				if (SecondName !='')
				{
				strApplicant1Title = strApplicant1Title + SecondName + ' ';
				}	    
				if (Surname != '')
				{
				strApplicant1Title = strApplicant1Title + Surname;
				}
                            return (strApplicant1Title);
			}
			else if (Counter == '2')	
			{
				if (TitleType == 'O')
				{
					if (TitleOther != '')
					{
						strApplicant2Title = strApplicant2Title + TitleOther + ' ';
					}
				}
				else
				{
					if (TitleText != '')
					{
						strApplicant2Title = strApplicant2Title + TitleText + ' ';
					}
				}

				if (FirstName !='')
				{
				strApplicant2Title = strApplicant2Title + FirstName + ' ';
				}
				if (SecondName !='')
				{
				strApplicant2Title = strApplicant2Title + SecondName + ' ';
				}	    
				if (Surname != '')
				{
				strApplicant2Title = strApplicant2Title + Surname;
				}
                            return (strApplicant2Title);
			}
		
        }    


		function GetApplicant1Title()
		{
			return(strApplicant1Title);
		}
		
		function GetApplicant2Title()
		{
			return(strApplicant2Title);
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
		 
		if (Applicant2Title != '')
	                strAddress = strAddress +  Applicant1Title + ' & ' + Applicant2Title + ' \\par ';	 
		else
                        strAddress = strAddress +  Applicant1Title + ' \\par ';
		
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




	/*  PM 28/6/2006 EP890 Removed original GetAddress function that only had 1 title parameter */
        function GetAddress(TitleType, TitleText, TitleOther, FirstName, Surname, FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	
		if (TitleType == 'O')
		{
			if (TitleOther != '')
			{
				strAddress = strAddress +  TitleOther + ' ' ;
			}
		}
		else
		{
			if (TitleText != '')
			{
				strAddress = strAddress + TitleText + ' ';
			}
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




var intTenureType = '';
function SetTenureType(Tenure)
{
	 intTenureType=Tenure;

   return(intTenureType);
}   

function GetTenureType()
{
   return(intTenureType);
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
   		
		today = day + ' ' + month + ', ' + year;
			
		return(today);
	}




/*  PM 28/6/2006 EP890 Removed original GetBothNames function that only had 1 title parameter */
function GetBothNames(TitleType, TitleText, TitleOther, Surname)
{
	var strAddress = "";
	
		if (TitleType == 'O')
		{
			if (TitleOther != '')
			{
				strAddress = strAddress +  TitleOther + ' ' ;
			}
		}
		else
		{
			if (TitleText != '')
			{
				strAddress = strAddress + TitleText + ' ';
			}
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
					strCustomerNo= strCustomerNo  + ' & ' + CustomerNo ;
			}	
		return(strCustomerNo);
	    }
        
		function GetCustomerNo()
		{
			return(strCustomerNo);
		}
  
  
var TenType = '';
function SetTenure(Tenure)
{
	 TenType=Tenure;
   return(TenType);
}   

function GetTenure()
{
   return(TenType);
}
  
  
  
    ]]></msxsl:script>
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="Order_No"/> 
    <xsl:variable name="TenureType"/>   
    <xsl:variable name="TenureType1"/>       
    <xsl:variable name="TenureType2"/>       
    <xsl:variable name="CustomerNo"/>      
    <xsl:variable name="ApplicantTitles"/>    
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustCount" select="count(RESPONSE/INTERESTRATECHANGENOTIFICATIONCUSTOMER/UNIQUEAPPLICATIONADDRESSES)"/>
			<xsl:variable name="CustRole" select="count(RESPONSE/INTERESTRATECHANGENOTIFICATIONCUSTOMER/CUSTOMERROLE)"/>

			<xsl:variable name="vProductDesc" select="string(.//APPLICATIONFACTFIND/ACTIVEQUOTE/MORTGAGESUBQUOTE/LOANCOMPONENT/MORTGAGEPRODUCTLANGUAGE/@PRODUCTTEXTDETAILS)"/>
			<xsl:variable name="vRevisedInterestRate" select="string(.//APPLICATIONFACTFIND/ACTIVEQUOTE/MORTGAGESUBQUOTE/LOANCOMPONENT/@RESOLVEDRATE)"/>
			<xsl:variable name="vPayment" select="string(.//APPLICATIONFACTFIND/ACTIVEQUOTE/MORTGAGESUBQUOTE/LOANCOMPONENT//@GROSSMONTHLYCOST)"/>

// DRC Epsom 1 address with 2 customers
			<xsl:if test="($CustCount = '1')">
				<xsl:if test="($CustRole = '2')">
						<xsl:for-each select="//CUSTOMERROLE">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<!-- PM 28/06/2006 EP890 - Start -->
									<!-- <xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),1)"/> -->
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(
										string(.//CUSTOMERVERSION/@TITLE_TYPES), 
										string(.//CUSTOMERVERSION/@TITLE_TEXT), 
										string(.//CUSTOMERVERSION/@TITLEOTHER), 
										string(.//CUSTOMERVERSION/@FIRSTFORENAME),
										string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<!-- PM 28/06/2006 EP890 - End -->
									<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),1)"/>									
									<xsl:variable name="TenureType1" select="msg:SetTenure(.//NEWPROPERTY/@TENURETYPE)"/>									
									<!-- PM 28/06/2006 EP890 - Start -->
									<!-- <xsl:variable name="AddressName" select="msg:GetApplicantNameWithInitials(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME), string(.//CUSTOMERVERSION/@SECONDFORENAME),  string(.//CUSTOMERVERSION/@SURNAME),1)"/> -->
									<xsl:variable name="AddressName" select="msg:GetApplicantNameWithInitials(
										string(.//CUSTOMERVERSION/@TITLE_TYPES), 
										string(.//CUSTOMERVERSION/@TITLE_TEXT), 
										string(.//CUSTOMERVERSION/@TITLEOTHER), 
										string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
										string(.//CUSTOMERVERSION/@SECONDFORENAME),  
										string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<!-- PM 28/06/2006 EP890 - End -->
							</xsl:if>

							<xsl:if test="($Order_No = '2')">
									<!-- PM 28/06/2006 EP890 - Start -->
									<!-- <xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),2)"/> -->
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(
										string(.//CUSTOMERVERSION/@TITLE_TYPES), 
										string(.//CUSTOMERVERSION/@TITLE_TEXT), 
										string(.//CUSTOMERVERSION/@TITLEOTHER), 
										string(.//CUSTOMERVERSION/@FIRSTFORENAME),
										string(.//CUSTOMERVERSION/@SURNAME),2)"/>
									<!-- PM 28/06/2006 EP890 - End -->
									<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),2)"/>									
									<!-- PM 28/06/2006 EP890 - Start -->
									<!-- <xsl:variable name="AddressName" select="msg:GetApplicantNameWithInitials(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME), string(.//CUSTOMERVERSION/@SECONDFORENAME),  string(.//CUSTOMERVERSION/@SURNAME),1)"/> -->
									<xsl:variable name="AddressName" select="msg:GetApplicantNameWithInitials(
										string(.//CUSTOMERVERSION/@TITLE_TYPES), 
										string(.//CUSTOMERVERSION/@TITLE_TEXT), 
										string(.//CUSTOMERVERSION/@TITLEOTHER), 
										string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
										string(.//CUSTOMERVERSION/@SECONDFORENAME),  
										string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<!-- PM 28/06/2006 EP890 - End -->
							</xsl:if>
						</xsl:for-each>	
						<xsl:element name="INFORMATION">
								<xsl:element name="CUSTOMERDETAILS">
									<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>
									<!-- PB 13/06/2006 EP721 Begin
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
												string(.//CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute> -->
									<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:DealWithAddress(
												string(msg:GetSalutation()),										
												string(''),
												string(''),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@FLATNUMBER),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENAME),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENUMBER),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@STREET),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@DISTRICT),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@TOWN),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@COUNTY),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<!-- PB EP721 End -->
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
									<xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="msg:GetCustomerNo()"/></xsl:attribute>									
									<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>

									<xsl:attribute name="PRODUCTDESCRIPTION"><xsl:value-of select="$vProductDesc"/></xsl:attribute>
									<xsl:attribute name="REVISEDINTERESTRATE"><xsl:value-of select="$vRevisedInterestRate"/></xsl:attribute>
									<xsl:attribute name="PAYMENT"><xsl:value-of select="$vPayment"/></xsl:attribute>

									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
								</xsl:element>			
						</xsl:element>
				</xsl:if>		
			</xsl:if>
//DRC Epsom 1 address with 1 customer OR 2 addresses with 2 customers
			<xsl:if test="(($CustCount = '1' and $CustRole = '1') or ($CustCount = '2' and $CustRole = '2'))">
				<xsl:for-each select="//CUSTOMERROLE">
					<xsl:element name="INFORMATION">
							<xsl:element name="CUSTOMERDETAILS">
								<xsl:attribute name="CURRENTDATE"><xsl:value-of select="msg:GetDate()"/></xsl:attribute>							
								<!-- PB 13/06/2006 EP721 Begin
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
											string(.//CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute> -->
								<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msg:GetAddress(
											string(.//CUSTOMERVERSION/@TITLE_TYPES), 
											string(.//CUSTOMERVERSION/@TITLE_TEXT), 
											string(.//CUSTOMERVERSION/@TITLEOTHER), 
											string(.//CUSTOMERVERSION/@FIRSTFORENAME),
											string(.//CUSTOMERVERSION/@SURNAME),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@FLATNUMBER),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENAME),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENUMBER),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@STREET),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@DISTRICT),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@TOWN),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@COUNTY),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@POSTCODE))"/></xsl:attribute>
								<!-- PB EP721 End -->
								<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>

								<!-- PM 28/6/2006 EP890 Start -->
								<!--<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetBothNames(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@SURNAME))"/></xsl:attribute> -->
								<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetBothNames(
											string(.//CUSTOMERVERSION/@TITLE_TYPES), 
											string(.//CUSTOMERVERSION/@TITLE_TEXT), 
											string(.//CUSTOMERVERSION/@TITLEOTHER), 
											string(.//CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
								<!-- PM 28/6/2006 EP890 End -->
								
								<xsl:attribute name="PRODUCTDESCRIPTION"><xsl:value-of select="$vProductDesc"/></xsl:attribute>
								<xsl:attribute name="REVISEDINTERESTRATE"><xsl:value-of select="$vRevisedInterestRate"/></xsl:attribute>
								<xsl:attribute name="PAYMENT"><xsl:value-of select="$vPayment"/></xsl:attribute>
															
								<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
								<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
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
