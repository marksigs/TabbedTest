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
	    function SetSalutation(Title,FirstName,SecondName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					titlename= titlename + Title ;
				
				if (FirstName !='')
					titlename= titlename + ' ' + FirstName ;

				if (SecondName !='')
					titlename= titlename  + ' ' + SecondName;
										
				if (Surname !='')
					titlename=titlename + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					titlename= titlename  + ' and ' + Title ;

				if (FirstName !='')
					titlename= titlename + ' ' + FirstName ;

				if (SecondName !='')
					titlename= titlename  + ' ' + SecondName;
										
				if (Surname !='')
					titlename=titlename + ' ' + Surname ;
			}	
		return(titlename);
	    }
        
		function GetSalutation()
		{
			return(titlename);
		}


		var Solname ='';	    
	    function GetSolicitorName(Title,FirstName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					Solname= Solname + Title ;
				
				if (FirstName !='')
					Solname= Solname + ' ' + FirstName ;

				if (Surname !='')
					Solname=Solname + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					Solname= Solname  + ' and ' + Title ;

				if (FirstName !='')
					Solname= Solname + ' ' + FirstName ;

				if (Surname !='')
					Solname=Solname + ' ' + Surname ;
			}	
		return(Solname);
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
   month =month +1;		
		
    		
   today= day + '/' + month + '/' + year;
   return(today);
}


function GetBothNames(Title,Firstname,Secondname,Surname)
{
	var strAddress = "";
	
		 if (Title != '')
		{
		strAddress = strAddress +  Title + ' ' ;
		}

		 if (Firstname != '')
		{
				strAddress = strAddress +  Firstname + ' ' ;
		}

		 if (Secondname != '')
		{
				strAddress = strAddress +  Secondname + ' ' ;
		}		
		
		 if (Surname != '')
		{
				strAddress = strAddress +  Surname + ' ' ;
		}
		return (strAddress);
}
  
	var strMortgageAmt ='';	    
	    function SetMortgageAmt(MortgageAmount)
	    {
			strMortgageAmt= MortgageAmount;
	    }
        
		function GetMortgageAmt()
		{
			return(strMortgageAmt);
		}
  
	var strMaxQuoteNo ='';	    
	    function SetQuoteNo(QuoteNo)
	    {
			strMaxQuoteNo=QuoteNo;
			return(strMaxQuoteNo);
	    }
        
		function GetQuoteNo()
		{
			return(strMaxQuoteNo);
		}
  


		var Applicant1 ='';	    
	    function SetApplicant1(Title,FirstName,SecondName,Surname)
	    {
				if (Title !='')
					Applicant1= Applicant1 + Title ;
				
				if (FirstName !='')
					Applicant1= Applicant1 + ' ' + FirstName ;

				if (SecondName !='')
					Applicant1= Applicant1 + ' ' + SecondName ;

				if (Surname !='')
					Applicant1=Applicant1 + ' ' + Surname;
		return(Applicant1);
	    }  

	    function GetApplicant1()
	    {
			return(Applicant1);
	    }
  

		var Applicant2 ='';	    
	    function SetApplicant2(Title,FirstName,SecondName,Surname)
	    {
				if (Title !='')
					Applicant2= Applicant2 + Title ;
				
				if (FirstName !='')
					Applicant2= Applicant2 + ' ' + FirstName ;

				if (SecondName !='')
					Applicant2= Applicant2 + ' ' + SecondName ;					

				if (Surname !='')
					Applicant2=Applicant2 + ' ' + Surname;
		return(Applicant2);
	    }  

	    function GetApplicant2()
	    {
			return(Applicant2);
	    }




  
    ]]></msxsl:script>
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="Order_No"/> 
    <xsl:variable name="MaxSubQuote"/>     
    <xsl:variable name="CurrentSubQuoteNo"/>
    <xsl:variable name="Applicant1"/>
    <xsl:variable name="Applicant2"/>        
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustRole" select="count(RESPONSE/COTDOCUMENT/CUSTOMERROLE)"/>
			<xsl:variable name="MaxSubQuote" select="count(RESPONSE/COTDOCUMENT/CUSTOMERROLE/MORTGAGESUBQUOTE)"/>
						<xsl:for-each select="//CUSTOMERROLE">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<xsl:variable name="Applicant1" select="msg:SetApplicant1(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME))"/>																																			   
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME),2)"/>
									<xsl:variable name="Applicant2" select="msg:SetApplicant2(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME))"/>																																			   
							</xsl:if>
						</xsl:for-each>	
						<xsl:element name="INFORMATION">
								<xsl:element name="CUSTOMERDETAILS">
									<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>							
									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
									<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@STREET),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@TOWN),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<xsl:attribute name="APPLICANTONE"><xsl:value-of select="msg:GetApplicant1()"/></xsl:attribute>
									<xsl:attribute name="APPLICANTTWO"><xsl:value-of select="msg:GetApplicant2()"/></xsl:attribute>	
									<xsl:attribute name="MORTGAGEDATE">
										<xsl:value-of select="//APPLICATIONOFFER/@OFFERISSUEDATE"/>
									</xsl:attribute>		
							   </xsl:element>					
						</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
