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
        
          
         var CustCount = new Array(0); 
        function Count(WhichCounter)
        {
            CustCount[WhichCounter] = CustCount[WhichCounter] + 1;
            return (CustCount[WhichCounter]);
        }          
            


		var Custname1 ='';
		var Custname2 ='';    
	    function SetSalutation(Title,FirstName,SecondName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					Custname1= Custname1+ Title ;
				
				if (FirstName !='')
					Custname1= Custname1+ ' ' + FirstName ;

				if (SecondName !='')
					Custname1= Custname1+ ' ' + SecondName;
										
				if (Surname !='')
					Custname1=Custname1+ ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					Custname2 = Custname2 + Title ;

				if (FirstName !='')
					Custname2 = Custname2 + ' ' + FirstName ;

				if (SecondName !='')
					Custname2 = Custname2 + ' ' + SecondName;
										
				if (Surname !='')
					Custname2 =Custname2 + ' ' + Surname ;
			}	
		return(Custname2 );
	    }
        
        var Name='';
		function GetSalutation()
		{
		  if (Custname2 !='')
		       Name=Custname1 + ' and ' + Custname2;
		  if (Custname2 =='')
		       Name=Custname1;

		return(Name);
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
        



        function DealWithAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
		
        if (FlatNumber.length > 0)
		{
		strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
		}
        if (HouseName.length > 0)
		{
		strAddress = strAddress + HouseName + ' \\par ';
		}
        if (HouseNumber.length > 0)
		{
		strAddress = strAddress + HouseNumber + ' ';
		}
        if (Street.length > 0)
		{
		strAddress = strAddress + Street;
		}
        if (District.length > 0)
		{
		  if((Street.length > 0) || (HouseNumber.length > 0))
		  {
			strAddress = strAddress + ' \\par ' + District;			
		  }	
		  else 
		  {
  			strAddress = strAddress + District;			
		  }	
		}
        if (Town.length > 0)
		{
		strAddress = strAddress + ' \\par ' + Town;
		}
        if (County.length > 0)
		{
		strAddress = strAddress + ' \\par ' + County;
		}
	if (PostCode.length > 0)
		{
		strAddress = strAddress + ' \\par ' + PostCode;
		}		
	return (strAddress);
        }


        function GetAddress(Title,FirstName,Surname,FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	

	if (Title.length > 0)
	{
		strAddress = strAddress + Title + ' ';
	}	 
	if ((FirstName.length > 0) || (Surname.length > 0)) 
		{
		    if (FirstName.length > 0)
		    {
			strAddress = strAddress + FirstName.substring(0,1);
		    }
		strAddress = strAddress +  ' ' + Surname + '\\par ' ;
		}
		
        if (FlatNumber.length > 0)
		{
		strAddress = strAddress + 'Flat ' + FlatNumber + ', ';
		}
        if (HouseName.length > 0)
		{
		strAddress = strAddress + HouseName + '\\par ';
		}
        if (HouseNumber.length > 0)
		{
		strAddress = strAddress + HouseNumber + ' ';
		}
        if (Street.length > 0)
		{
		strAddress = strAddress + Street;
		}
        if (District.length > 0)
		{
		  if((Street.length > 0) || (HouseNumber.length > 0))
		  {
			strAddress = strAddress + '\\par ' + District;			
		  }	
		  else 
		  {
  				strAddress = strAddress + District;			
		  }	
		}
        if (Town.length > 0)
		{
		strAddress = strAddress + '\\par ' + Town;
		}
        if (County.length > 0)
		{
		strAddress = strAddress + '\\par ' + County;
		}
		if (PostCode.length > 0)
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
  
  function GetCountry(Country)
  {
  return(Country);
  }
  
  
  		var Applicant1='';
	    function SetResidentName(Title,Firstname,Secondname,Othername,Surname)
	    {
	    var Applicant='';

				if (Title !='')
					Applicant= Title + ' ';
				if (Firstname !='')
					Applicant=Applicant + Firstname.substring(0,1) + ' ';
				if (Secondname !='')
					Applicant=Applicant + Secondname.substring(0,1) + ' ';
				if (Othername !='')
					Applicant=Applicant + Othername.substring(0,1) + ' ';					
				if (Surname !='')
					Applicant= Applicant + Surname;
			
			 Applicant1=Applicant;

		return(Applicant1);
	    }
        
		function GetResidentName()
		{
			return(Applicant1);
		}


		var strApplicant1Title='';
		var strApplicant2Title='';
        function GetApplicantNameWithInitials(Title,FirstName,SecondName,Surname,Counter)
        {
			if (Counter == '1')	
			{
				if (Title != '')
				{
				strApplicant1Title = strApplicant1Title + Title + ' ';
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
			}
			else if (Counter == '1')	
			{
				if (Title != '')
				{
				strApplicant2Title = strApplicant2Title + Title + ' ';
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
			}
		return (strApplicant1Title);
        }    



		function GetApplicant1Title()
		{
			return(strApplicant1Title);
		}
		
		function GetApplicant2Title()
		{
			return(strApplicant2Title);
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
		  

  
  
    ]]></msxsl:script>
  
	<xsl:template match="/">
	<xsl:element name="TEMPLATEDATA">
		<xsl:variable name="Applicant"/>
  
		<xsl:variable name="CustRole" select="count(//CUSTOMERROLE)"/>
           		

		<xsl:for-each select="//NEWPROPERTYADDRESS/ADDRESS">
       	        <xsl:attribute name="NEWADDRESS">
      	        	<xsl:value-of select="msg:DealWithNewAddress(
						string(@FLATNUMBER),
						string(@BUILDINGORHOUSENAME),
						string(@BUILDINGORHOUSENUMBER),
						string(@STREET),
						string(@DISTRICT),
						string(@TOWN),
						string(@COUNTY),
						string(@POSTCODE))"/>
        	</xsl:attribute>
       	    	</xsl:for-each>

		<xsl:attribute name="CUSTOMERNUMBER">
                	<xsl:value-of select="concat('{', string(//TEMPLATEDATA/ADDRESS/CUSTOMER/@CUSTOMERNUMBER), '}')" /> 
            	</xsl:attribute>

	    	<xsl:attribute name="CURRENTADDRESS">
      	        	<xsl:value-of select="msg:DealWithAddress(
						string(//TEMPLATEDATA/ADDRESS/@FLATNUMBER),
						string(//TEMPLATEDATA/ADDRESS/@BUILDINGORHOUSENAME),
						string(//TEMPLATEDATA/ADDRESS/@BUILDINGORHOUSENUMBER),
						string(//TEMPLATEDATA/ADDRESS/@STREET),
						string(//TEMPLATEDATA/ADDRESS/@DISTRICT),
						string(//TEMPLATEDATA/ADDRESS/@TOWN),
						string(//TEMPLATEDATA/ADDRESS/@COUNTY),
						string(//TEMPLATEDATA/ADDRESS/@POSTCODE))"/>
	    	</xsl:attribute>

                <xsl:attribute name="APPLICATIONNUMBER">
                	<xsl:value-of select="concat('{', string(//TEMPLATEDATA/ADDRESS/@APPLICATIONNUMBER), '}')"/>
            	</xsl:attribute>

            	<xsl:attribute name="CURRENTDATE">
                	<xsl:value-of select="msg:GetDate()" /> 
            	</xsl:attribute>

       	    	<xsl:attribute name="APPLICANT">
               	<xsl:for-each select="//TEMPLATEDATA/ADDRESS/CUSTOMER">
                    <xsl:variable name="AppCounter" select="msg:UpdateCounter(2)"/>
       	            <xsl:variable name="Title" select="@TITLE"/>
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



		<!-- <xsl:if test="($CustRole > '0')"> -->
			<xsl:for-each select="//CUSTOMERROLE">
				<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
				<xsl:if test="($Order_No = '1')">
					<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT), string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																						       string(.//CUSTOMERVERSION/@SECONDFORENAME),string(.//CUSTOMERVERSION/@SURNAME),1)"/>
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																						        string(.//CUSTOMERVERSION/@SECONDFORENAME),string(.//CUSTOMERVERSION/@SURNAME),2)"/>
							</xsl:if>
						</xsl:for-each>			
				

								
									<xsl:attribute name="SOLICITORDETAILS"><xsl:value-of select="msg:GetAddress(
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
																
									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
													
		
					
		<xsl:attribute name="MORTGAGEAMOUNT">
			<xsl:value-of select=".//APPLICATIONFACTFIND/QUOTATION/MORTGAGESUBQUOTE/@TOTALLOANAMOUNT"/></xsl:attribute>

				
		<xsl:attribute name="SOLICITORPANELNO">
			<xsl:value-of select="concat('{',//APPLICATIONLEGALREP/PANEL/@PANELID,'}')"/>
		</xsl:attribute>	
								    									
									
									
			<xsl:attribute name="ROTDATE">
			<xsl:value-of select=".//REPORTONTITLE/@COMPLETIONDATE"/></xsl:attribute>
									<xsl:attribute name="TITLENO"><xsl:value-of select="concat('{',.//REPORTONTITLE/@TITLENUMBER,'}')"/></xsl:attribute>									
									<xsl:attribute name="OTHERSYSTEMACCOUNTNO"><xsl:value-of select="concat('{',//APPLICATION/@OTHERSYSTEMACCOUNTNUMBER,'}')"/></xsl:attribute
>
									<xsl:attribute name="SALUTATION"><xsl:value-of select="msg:GetSexDescription(
																											string(.//APPLICATIONLEGALREP/CONTACTDETAILS/@CONTACTTITLE))"/></xsl:attribute>										<xsl:variable name="Country" select="msg:GetCountry(.//APPLICATIONLEGALREP/THIRDPARTY/ADDRESS/@COUNTRY)"/>
								      <xsl:attribute name="Country"><xsl:value-of select="string($Country)"/></xsl:attribute>

								      <xsl:if test="($Country = '2' or $Country = '4')">
								      		<xsl:attribute name="COTROT"><xsl:value-of select="string('certificate of')"/></xsl:attribute>
								      		<xsl:attribute name="COTROTCAPS"><xsl:value-of select="string('Certificate of')"/></xsl:attribute>
								      		<xsl:attribute name="COTROTTITLE"><xsl:value-of select="string('Certificate of ownership')"/></xsl:attribute>
								      		<xsl:attribute name="TRANSFER"><xsl:value-of select="string('transfer')"/></xsl:attribute>
								      		<xsl:attribute name="ROTCOT"><xsl:value-of select="string('certificate')"/></xsl:attribute>								      		
								      		<xsl:element name="COT"/>
								      </xsl:if>
								      <xsl:if test="($Country = '3')">
								      		<xsl:attribute name="COTROT"><xsl:value-of select="string('report on')"/></xsl:attribute>
								      		<xsl:attribute name="COTROTCAPS"><xsl:value-of select="string('Report on')"/></xsl:attribute>								      
								      		<xsl:attribute name="COTROTTITLE"><xsl:value-of select="string('Report on title')"/></xsl:attribute>
								      		<xsl:attribute name="TRANSFER"><xsl:value-of select="string('Disposition')"/></xsl:attribute>
								      		<xsl:attribute name="ROTCOT"><xsl:value-of select="string('report on title')"/></xsl:attribute>								      		
								      		<xsl:element name="ROT"/>
								      </xsl:if>	





								      
	
								      
								
				<!-- </xsl:if>		-->
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
