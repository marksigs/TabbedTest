<?xml version="1.0" encoding="UTF-8"?>
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
  
  
    ]]></msxsl:script>
    <xsl:template match="/">
	<xsl:element name="TEMPLATEDATA">

            <xsl:variable name="Applicant"/>

            <xsl:variable name="Salutation"/>

            <xsl:attribute name="APPLICATIONNUMBER">
                <xsl:value-of select="concat('{', string(//TEMPLATEDATA/ADDRESS/@APPLICATIONNUMBER), '}')"/>
            </xsl:attribute>
            <xsl:attribute name="CURRENTDATE">
                <xsl:value-of select="msg:GetDate()" /> 
            </xsl:attribute>

       	    <xsl:attribute name="NEWADDRESS">
	    	<xsl:value-of select="msg:DealWithNewAddress(
			string(//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
			string(//NEWPROPERTYADDRESS/ADDRESS/@STREET),
			string(//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
			string(//NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
			string(//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/>
            </xsl:attribute>

	    <xsl:attribute name="ADDRESS">
      	        <xsl:value-of select="msg:DealWithAddress(
						string(msg:GetSalutation()),										
						string(''),
						string(''),
						string(//TEMPLATEDATA/ADDRESS/@FLATNUMBER),
						string(//TEMPLATEDATA/ADDRESS/@BUILDINGORHOUSENAME),
						string(//TEMPLATEDATA/ADDRESS/@BUILDINGORHOUSENUMBER),
						string(//TEMPLATEDATA/ADDRESS/@STREET),
						string(//TEMPLATEDATA/ADDRESS/@DISTRICT),
						string(//TEMPLATEDATA/ADDRESS/@TOWN),
						string(//TEMPLATEDATA/ADDRESS/@COUNTY),
						string(//TEMPLATEDATA/ADDRESS/@POSTCODE))"/>
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
            <xsl:variable name="TempCounter" select="msg:ResetCounter(2)"/>
            <xsl:attribute name="SALUTATION">
       		<xsl:variable name="NumberOfApplicants" select="string(//TEMPLATEDATA/ADDRESS/@CUSTOMERSATADDRESS)"/>
                <xsl:variable name="Order" select="1"/>

                <xsl:variable name="Title1" select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[1]/@TITLE)"/>
                <xsl:variable name="FirstInitial1" select="substring(//TEMPLATEDATA/ADDRESS/CUSTOMER[1]/@FIRSTFORENAME, 1, 1)"/>
                <xsl:variable name="SecondInitial1" select="substring(//TEMPLATEDATA/ADDRESS/CUSTOMER[1]/@SECONDFORENAME, 1, 1)"/>
                <xsl:variable name="Surname1" select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[1]/@SURNAME)"/>
                <xsl:variable name="Title2" select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[2]/@TITLE)"/>
                <xsl:variable name="FirstInitial2" select="substring(//TEMPLATEDATA/ADDRESS/CUSTOMER[2]/@FIRSTFORENAME, 1, 1)"/>
                <xsl:variable name="SecondInitial2" select="substring(//TEMPLATEDATA/ADDRESS/CUSTOMER[2]/@SECONDFORENAME, 1, 1)"/>
                <xsl:variable name="Surname2" select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[2]/@SURNAME)"/>
                <xsl:choose>
                    <xsl:when test="($NumberOfApplicants = '1')">
                        <xsl:if test="($Title1 != '')">
                            <xsl:value-of select="concat($Salutation, $Title1, ' ')"/>
                        </xsl:if>
                        <xsl:if test="($Surname1 != '')">
                            <xsl:value-of select="concat($Salutation, $Surname1)"/>
                        </xsl:if>
                    </xsl:when>
                    <xsl:when test="($NumberOfApplicants = '2' and $Title1 = 'Mr' and $Title2 = 'Mrs' and $Surname1 = $Surname2)">
                        <xsl:value-of select="concat($Salutation, 'Mr and Mrs ', $Surname1)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:for-each select="//CUSTOMER">
                            <xsl:variable name="AppCounter" select="msg:UpdateCounter(2)"/>
       	                    <xsl:variable name="Title" select="@TITLE"/>
               	            <xsl:variable name="Initial1" select="@FIRSTFORENAME"/>
                       	    <xsl:variable name="Initial2" select="@SECONDFORENAME"/>
                            <xsl:variable name="Surname" select="@SURNAME"/>
       	                    <xsl:if test="($AppCounter != 1)">
               	                <xsl:value-of select="concat($Salutation, ' and ')"/>
                       	    </xsl:if>
                            <xsl:if test="($Title != '')">
       	                        <xsl:value-of select="concat($Salutation, $Title, ' ')"/>
               	            </xsl:if>
                       	    <xsl:if test="($Order != '1')">
                               	<xsl:if test="($Surname = .//TEMPLATEDATA/ADDRESS[$Order - 1]/@SURNAME)">
                            	    <xsl:if test="($Initial1 != '')">
       	                                <xsl:value-of select="concat($Salutation, substring($Initial1, 1, 1), ' ')"/>
               	                    </xsl:if>
                       	        </xsl:if>
                            </xsl:if>
       	                    <xsl:if test="($Surname != '')">
               	                <xsl:value-of select="concat($Salutation, $Surname)"/>
                       	    </xsl:if>
                        </xsl:for-each>
       	            </xsl:otherwise>
               	</xsl:choose>
       	    </xsl:attribute>
	</xsl:element>
    </xsl:template>
</xsl:stylesheet>
