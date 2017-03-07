<?xml version="1.0" encoding="UTF-8"?>
<!-- edited with XMLSpy v2005 rel. 3 U (http://www.altova.com) by ITS (Marlborough Stirling plc) -->
<!-- BC	03/04/2006	MAR1523	New template																					-->
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
        
        function ResetCounter(WhichCounter)
        {
            Temp[WhichCounter] = 0;
            return (0);
        }   
        
         var CustCount = new Array(0); 
        function Count(WhichCounter)
        {
            CustCount[WhichCounter] = CustCount[WhichCounter] + 1;
            return (CustCount[WhichCounter]);
        }          
            

    var Name1='';
    var Name2='';    
    function SetSalutation(Title,FirstName,SecondName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					Name1=  Title ;
				
				if (FirstName !='')
					Name1 = Name1 + ' ' + FirstName ;

				if (SecondName !='')
					Name1 = Name1 + ' ' + SecondName;
										
				if (Surname !='')
					Name1 =Name1 + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					Name2= Title ;

				if (FirstName !='')
					Name2 = Name2 + ' ' + FirstName ;

				if (SecondName !='')
					Name2 = Name2 + ' ' + SecondName;
										
				if (Surname !='')
					Name2 =Name2 + ' ' + Surname ;
			}	
		return(Name2 );
	    }
        
             var Name='';
		function GetSalutation()
		{
			if ((Name2 !='') && (Name1  !=''))
			    Name= Name1  + ' and ' + Name2;
			if ((Name2 =='') && (Name1 !=''))
			    Name=  Name1;
			if ((Name2 !='') && (Name1 ==''))
			    Name= Name2;
		     	
			return(Name);
		}




	var Custname1='';
	var Custname2='';
	    function SetName(Title,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					Custname1=  Title ;
				
				if (Surname !='')
					Custname1=Custname1+ ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					Custname2 =  Title ;

				if (Surname !='')
					Custname2 =Custname2 + ' ' + Surname ;
			}	
		return(Custname2 );
	    }
        
             var Name='';
		function GetName()
		{
			if ((Custname2 !='') && (Custname1  !=''))
			    Name= Custname1  + ' and ' + Custname2;
			if ((Custname2 =='') && (Custname1 !=''))
			    Name= Custname1;
			if ((Custname2 !='') && (Custname1 ==''))
			    Name= Custname2;
		     	
			return(Name);
		}


function DealWithAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
		
        if (FlatNumber.length > 0)
		{
		strAddress = strAddress + FlatNumber + ', ';
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
		strAddress = strAddress + Street + ' \\par ';
		}
        if (Town.length > 0)
		{
		strAddress = strAddress + Town + ' \\par ';
		}
        if (District.length > 0)
		{
  			strAddress = strAddress + District + ' \\par ';
		}
        if (County.length > 0)
		{
		strAddress = strAddress + County + ' \\par ';
		}
	if (PostCode.length > 0)
		{
		strAddress = strAddress + PostCode + ' \\par ';
		}
	return (strAddress);
        }



        function DealWithNewAddress(FlatNumber,BuildingOrHousename,BuildingOrHouseNumber, Street, District, Town, County, PostCode)
        {
	var strAddress = "";
	
	    if (FlatNumber != '')
		{
		 strAddress = strAddress + FlatNumber + ', ';
		}
	    if (BuildingOrHousename != '')
		{
		 strAddress = strAddress + BuildingOrHousename + ', ';
		}		
	    if (BuildingOrHouseNumber != '')
		{
		 strAddress = strAddress + BuildingOrHouseNumber + ', ';
		}				
	    if (Street != '')
		{
		 strAddress = strAddress + Street + ', ';
		}
	    if (Town != '')
		{
		 strAddress = strAddress + Town + ', ';
		}
	    if (District != '')
		{
		 strAddress = strAddress + District + ', ';
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
  
  
	    var strCustomerNo1 ='';	    
	    var strCustomerNo2 ='';	    
	    function SetCustomerNo(CustomerNo,Counter)
	    {
			if (Counter == '1') 
			{
				if (CustomerNo !='')
					strCustomerNo1= CustomerNo ;
			}
		  else if (Counter == '2') 
			{
				if (CustomerNo !='')
					strCustomerNo2= CustomerNo ;
			}	
		return(strCustomerNo);
	    }
        
		var strCustomerNo='';
		function GetCustomerNo()
		{
			if ((strCustomerNo2 !='') && (strCustomerNo1  !=''))
			    strCustomerNo= strCustomerNo1  + ' and ' + strCustomerNo2;
			if ((strCustomerNo2 =='') && (strCustomerNo1 !=''))
			    strCustomerNo= strCustomerNo1;
			if ((strCustomerNo2 !='') && (strCustomerNo1 ==''))
			    strCustomerNo= strCustomerNo2;
		     	
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
    	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
		       <xsl:variable name="Applicant"/>

           		 <xsl:variable name="Salutation"/>
           		 
		      <xsl:for-each select="//CUSTOMERROLE">
				<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
				<xsl:if test="($Order_No = '1')">
						<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																				     string(.//CUSTOMERVERSION/@FIRSTFORENAME),	
																				     string(.//CUSTOMERVERSION/@SECONDFORENAME),
																				     string(.//CUSTOMERVERSION/@SURNAME),1)"/>
						<xsl:variable name="Names" select="msg:SetName(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@SURNAME),1)"/>	
						<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//@CUSTOMERNUMBER,'}'),1)"/>
																				     																			     
				</xsl:if>		
				<xsl:if test="($Order_No = '2')">
						<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																				     string(.//CUSTOMERVERSION/@FIRSTFORENAME),	
																				     string(.//CUSTOMERVERSION/@SECONDFORENAME),
																				     string(.//CUSTOMERVERSION/@SURNAME),2)"/>
						<xsl:variable name="Names" select="msg:SetName(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																				     string(.//CUSTOMERVERSION/@SURNAME),2)"/>								
						<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//@CUSTOMERNUMBER,'}'),2)"/>											
				</xsl:if>
			</xsl:for-each>           		 

	            <xsl:attribute name="APPLICATIONNUMBER">
	                <xsl:value-of select="concat('{', string(//CUSTOMERROLE/@APPLICATIONNUMBER), '}')"/>
	            </xsl:attribute>
	            
	            <xsl:attribute name="CURRENTDATE">
	                <xsl:value-of select="msg:GetDate()" /> 
	            </xsl:attribute>
		
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
		    <xsl:attribute name="ADDRESS">
	      	        <xsl:value-of select="msg:DealWithAddress(
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@FLATNUMBER),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@STREET),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@DISTRICT),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@TOWN),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@COUNTY),
							string(//CUSTOMERROLE/CUSTOMERADDRESS/ADDRESS/@POSTCODE))"/>
		    </xsl:attribute>
	          <xsl:attribute name="APPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>
	          
	          <xsl:attribute name="SALUTATION"><xsl:value-of select="msg:GetName()"/></xsl:attribute>
	          <xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="msg:GetCustomerNo()"/></xsl:attribute>  
		  <xsl:attribute name="OTHERSYSTEMACCOUNTNUMBER">
			<xsl:value-of select="concat('{',//CUSTOMERROLE/CUSTOMER/@OTHERSYSTEMCUSTOMERNUMBER,'}')"/>
	    	  </xsl:attribute>
	</xsl:element>
    </xsl:template>
</xsl:stylesheet>
