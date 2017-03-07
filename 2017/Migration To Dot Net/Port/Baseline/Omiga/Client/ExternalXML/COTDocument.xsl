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
  
  
  
    ]]></msxsl:script>
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="Order_No"/> 
    <xsl:variable name="MaxSubQuote"/>     
    <xsl:variable name="CurrentSubQuoteNo"/>
    <xsl:variable name="PanelDetails"/>    
    <xsl:variable name="ThirdPartyDetails"/>        
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustRole" select="count(RESPONSE/COTDOCUMENT/CUSTOMERROLE)"/>
			<xsl:variable name="MaxSubQuote" select="count(RESPONSE/COTDOCUMENT/CUSTOMERROLE/MORTGAGESUBQUOTE)"/>
			<xsl:variable name="PanelDetails" select="count(RESPONSE/COTDOCUMENT/APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT)"/>
			<xsl:variable name="ThirdPartyDetails" select="count(RESPONSE/COTDOCUMENT/REPORTONTITLE/ROTSOLICITORSBANKACCOUNT)"/>			
				<xsl:if test="($CustRole > '0')">
						<xsl:for-each select="//CUSTOMERROLE">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME),1)"/>
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																			   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																			   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																			   string(.//CUSTOMERVERSION/@SURNAME),2)"/>
							</xsl:if>
						</xsl:for-each>	
						<xsl:element name="INFORMATION">
								<xsl:element name="CUSTOMERDETAILS">
									<xsl:attribute name="SOLICITORDETAILS"><xsl:value-of select="msg:DealWithAddress(
												msg:GetSolicitorName(string(//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE_TEXT),
																				    string(//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTFORENAME),
																					string(//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTSURNAME),1),										
												string(''),
												string(''),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@FLATNUMBER),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENAME),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@BUILDINGORHOUSENUMBER),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@STREET),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@DISTRICT),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@TOWN),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@COUNTY),
												string(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>							
									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
									<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@STREET),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@TOWN),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
								    <xsl:attribute name="DXID"><xsl:value-of select="concat('{',.//COTDOCUMENT/APPLICATIONLEGALREP/@DXID,'}')"/></xsl:attribute>				
									<xsl:for-each select="//CUSTOMERROLE/MORTGAGESUBQUOTE">
											<xsl:variable name="CurrentSubQuoteNo" select="msg:SetQuoteNo(.//@TOTALLOANAMOUNT)"/>
											<xsl:attribute name="MORTGAGEAMOUNT">
													<xsl:value-of select="string(msg:GetQuoteNo())"/>
											</xsl:attribute>		
									</xsl:for-each>
									<xsl:attribute name="PURCHASEPRICE">
										<xsl:value-of select=".//CUSTOMERROLE/APPLICATIONFACTFIND/@PURCHASEPRICEORESTIMATEDVALUE"/>
									</xsl:attribute>
									<xsl:attribute name="SOLICITORPANELNO">
										<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/@PANELID,'}')"/>
									</xsl:attribute>	
								    <xsl:attribute name="PanelDetails"><xsl:value-of select="string($PanelDetails)"/></xsl:attribute>				
								    <xsl:attribute name="ThirdPartyDetails"><xsl:value-of select="string($ThirdPartyDetails)"/></xsl:attribute>
									<xsl:if test="($PanelDetails > 0 and $ThirdPartyDetails > 0)">									   
											<xsl:attribute name="SOLICITORBANKNAME">
												<xsl:value-of select="//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@BANKNAME"/>
											</xsl:attribute>									
											<xsl:attribute name="SOLICITORSORTCODE">
												<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@BANKSORTCODE,'}')"/>
											</xsl:attribute>
											<xsl:attribute name="SOLICITORACCOUNTNO">
												<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@ACCOUNTNUMBER,'}')"/>
											</xsl:attribute>
									</xsl:if>		
									<xsl:if test="($PanelDetails = 0 and $ThirdPartyDetails > 0)">									   
											<xsl:attribute name="SOLICITORBANKNAME">
												<xsl:value-of select="//REPORTONTITLE/ROTSOLICITORSBANKACCOUNT/@BANKNAME"/>
											</xsl:attribute>									
											<xsl:attribute name="SOLICITORSORTCODE">
												<xsl:value-of select="concat('{',//REPORTONTITLE/ROTSOLICITORSBANKACCOUNT/@BANKSORTCODE,'}')"/>
											</xsl:attribute>
											<xsl:attribute name="SOLICITORACCOUNTNO">
												<xsl:value-of select="concat('{',//REPORTONTITLE/ROTSOLICITORSBANKACCOUNT/@BANKACCOUNTNUMBER,'}')"/>
											</xsl:attribute>
									</xsl:if>									
									<xsl:if test="($PanelDetails > 0 and $ThirdPartyDetails = 0)">									   
											<xsl:attribute name="SOLICITORBANKNAME">
												<xsl:value-of select="//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@BANKNAME"/>
											</xsl:attribute>									
											<xsl:attribute name="SOLICITORSORTCODE">
												<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@BANKSORTCODE,'}')"/>
											</xsl:attribute>
											<xsl:attribute name="SOLICITORACCOUNTNO">
												<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/PANELBANKACCOUNT/@ACCOUNTNUMBER,'}')"/>
											</xsl:attribute>
									</xsl:if>									
									<xsl:attribute name="MORTGAGEDATE">
										<xsl:value-of select="//APPLICATIONOFFER/@OFFERISSUEDATE"/>
									</xsl:attribute>
									<xsl:attribute name="DATE"><xsl:value-of select=".//REPORTONTITLE/@COMPLETIONDATE"/></xsl:attribute>
									<xsl:attribute name="TITLENO"><xsl:value-of select="concat('{',.//REPORTONTITLE/@TITLENUMBER,'}')"/></xsl:attribute>									
									<xsl:attribute name="OTHERSYSTEMACCOUNTNO"><xsl:value-of select="concat('{',//APPLICATION/@OTHERSYSTEMACCOUNTNUMBER,'}')"/></xsl:attribute
>
									<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:GetSexDescription(
																											string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE))"/></xsl:attribute>									
								</xsl:element>
						</xsl:element>
				</xsl:if>		
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
