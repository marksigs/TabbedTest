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



        function DealWithNewAddress(HouseName, FlatNumber, HouseNumber, Street, Town, County, PostCode)
        {
	var strAddress = "";
	
	    if (HouseName != '')
		{
		 strAddress = strAddress + HouseName + ', ';
		}
	    if (FlatNumber != '')
		{
		 strAddress = strAddress + FlatNumber + ', ';
		}
	    if (HouseNumber != '')
		{
		 strAddress = strAddress + HouseNumber + ', ';
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
  
 var strPaymentDay=''; 
function PaymentDay(PaymentDay)
{
	if (strPaymentDay == '')
	{
		if ((PaymentDay = '1') || (PaymentDay = '21') || (PaymentDay = '31'))
			strPaymentDay=strPaymentDay + PaymentDay + 'st of the month';
	   else if ((PaymentDay = '2') || (PaymentDay = '22'))
			strPaymentDay=strPaymentDay + PaymentDay + 'nd of the month';
	   else if ((PaymentDay = '3') || (PaymentDay = '23'))
			strPaymentDay=strPaymentDay + PaymentDay + 'rd of the month';		
	   else	
			strPaymentDay=strPaymentDay + PaymentDay + 'th of the month';		
	}
	return(strPaymentDay);	
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
  
  var intYear = '';
  var dt;
  var CurYear;  
  function GetYears(Year)
  {
	   dt = new Date();        
       CurYear = dt.getFullYear();
 	   intYear = CurYear - Year;
	  return(intYear);
  }
  
    ]]></msxsl:script>
	<xsl:variable name="SALUTATIONONE"/>
	<xsl:variable name="Order_No"/>
	<xsl:variable name="PropertySubsidence"/>
	<xsl:variable name="LongStandingSubsidence"/>
	<xsl:variable name="LargePanelSystemAppraised"/>
	<xsl:variable name="HistoricBldgRepairReq"/>
	<xsl:variable name="AsbestosPoorCondition"/>
	<xsl:variable name="ExtensionsOrAlterations"/>
	<xsl:variable name="Tenanted"/>
	<xsl:variable name="SharedAccess"/>
	<xsl:variable name="MiningReport"/>
	<xsl:variable name="NonResidentLandInd"/>
	<xsl:variable name="DvlpProposals"/>
	<xsl:variable name="NonSTDConstructionType"/>
	<xsl:variable name="SingleConstruction"/>
	<xsl:variable name="IsSingleTwoStorey"/>
	<xsl:variable name="PropertyType"/>
	<xsl:variable name="PropertyDescription"/>
	<xsl:variable name="NewPropertyInd"/>
	<xsl:variable name="Tenure"/>
	<xsl:variable name="Conditions"/>
	<xsl:variable name="Saleability"/>
	<xsl:variable name="ReTypeInd"/>
	<xsl:variable name="CertificationType"/>	
	<xsl:variable name="TotalYears"/>
	<xsl:variable name="SignatureReturned"/>	
	<xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">
			<xsl:variable name="CustCount" select="count(RESPONSE/VALUATIONREPORT/UNIQUEAPPLICATIONADDRESSES)"/>
			<xsl:variable name="CustRole" select="count(RESPONSE/VALUATIONREPORT/CUSTOMERROLE)"/>
			<xsl:variable name="PropertySubsidence" select="string(//GETLATESTVALUATIONREPORT/@PROPERTYSUBSIDENCE)"/>
			<xsl:variable name="LongStandingSubsidence" select="//GETLATESTVALUATIONREPORT/@LONGSTANDINGSUBIDENCE"/>
			<xsl:variable name="LargePanelSystemAppraised" select="//GETLATESTVALUATIONREPORT/@LARGEPANELSYSTEMAPPRAISED"/>
			<xsl:variable name="HistoricBldgRepairReq" select="//GETLATESTVALUATIONREPORT/@HISTORICBUILDINGREPAIRSREQ"/>
			<xsl:variable name="AsbestosPoorCondition" select="//GETLATESTVALUATIONREPORT/@ASBESTOSPOORCONDITION"/>
			<xsl:variable name="ExtensionsOrAlterations" select="//GETLATESTVALUATIONREPORT/@EXTENSIONSORALTERATIONS"/>
			<xsl:variable name="Tenanted" select="//GETLATESTVALUATIONREPORT/@TENANTEDPROPERTYIND"/>
			<xsl:variable name="SharedAccess" select="//GETLATESTVALUATIONREPORT/@UNADOPTEDSHAREDACCESSISSUES"/>
			<xsl:variable name="MiningReport" select="//GETLATESTVALUATIONREPORT/@MININGREPORTSREQ"/>
			<xsl:variable name="NonResidentLandInd" select="//GETLATESTVALUATIONREPORT/@NONRESIDENTIALLANDIND"/>
			<xsl:variable name="DvlpProposals" select="//GETLATESTVALUATIONREPORT/@DEVELOPMENTPROPOSALS"/>
			<xsl:variable name="NonSTDConstructionType" select="//GETLATESTVALUATIONREPORT/@NONSTANDARDCONSTRUCTIONTYPE"/>
			<xsl:variable name="SingleConstruction" select="//GETLATESTVALUATIONREPORT/@ISSINGLESKINCONSTRUCTION"/>
			<xsl:variable name="IsSingleTwoStorey" select="//GETLATESTVALUATIONREPORT/@ISSINGLESKINTWOSTOREY"/>
			<xsl:variable name="PropertyType" select="//GETLATESTVALUATIONREPORT/@TYPEOFPROPERTY"/>
			<xsl:variable name="PropertyDescription" select="//GETLATESTVALUATIONREPORT/@PROPERTYDESCRIPTION"/>
			<xsl:variable name="NewPropertyInd" select="//GETLATESTVALUATIONREPORT/@NEWPROPERTYINDICATOR"/>
			<xsl:variable name="Tenure" select="//GETLATESTVALUATIONREPORT/@TENURE"/>
			<xsl:variable name="Conditions" select="//GETLATESTVALUATIONREPORT/@OVERALLCONDITION"/>
			<xsl:variable name="Saleability" select="//GETLATESTVALUATIONREPORT/@SALEABILITY"/>
			<xsl:variable name="ReTypeInd" select="//GETLATESTVALUATIONREPORT/@RETYPEIND"/>
			<xsl:variable name="CertificationType" select="//GETLATESTVALUATIONREPORT/@CERTIFICATIONTYPE"/>			
			<xsl:variable name="SignatureReturned" select="//GETLATESTVALUATIONREPORT/@SIGNATURERETURNED"/>			
			
			<xsl:for-each select="//CUSTOMERROLE">
				<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
				<xsl:if test="($Order_No = '1')">
					<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																   string(.//CUSTOMERVERSION/@SURNAME),1)"/>
					<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),1)"/>
				</xsl:if>
				<xsl:if test="($Order_No = '2')">
					<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),
																																   string(.//CUSTOMERVERSION/@FIRSTFORENAME),
																																   string(.//CUSTOMERVERSION/@SECONDFORENAME),
																																   string(.//CUSTOMERVERSION/@SURNAME),2)"/>
					<xsl:variable name="CustomerNo" select="msg:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),2)"/>
				</xsl:if>
			</xsl:for-each>
			<xsl:element name="INFORMATION">
				<xsl:element name="VALUATIONDETAILS">
					<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
					<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:GetSalutation()"/></xsl:attribute>
					<xsl:attribute name="NEWADDRESS"><xsl:value-of select="msg:DealWithNewAddress(
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@STREET),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@TOWN),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
												string(.//NEWPROPERTY/NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
					<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
					<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
					<xsl:if test="($TempCounter > 1)">
						<xsl:element name="PAGEBREAK"/>
					</xsl:if>
					<xsl:attribute name="PRESENTVALUATION"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@PRESENTVALUATION)"/></xsl:attribute>
					<xsl:attribute name="REINSTATEMENTVALUE"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@REINSTATEMENTVALUE)"/></xsl:attribute>
					<xsl:if test="$PropertySubsidence = '1' ">
						<xsl:attribute name="PROPERTYSUBSIDENCE"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertySubsidence ='0' ">
						<xsl:attribute name="PROPERTYSUBSIDENCE"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$LongStandingSubsidence = '1' ">
						<xsl:attribute name="LONGSTANDINGSUBIDENCE"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$LongStandingSubsidence = '0' ">
						<xsl:attribute name="LONGSTANDINGSUBIDENCE"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$LargePanelSystemAppraised = '1' ">
						<xsl:attribute name="LARGEPANELSYSTEMAPPRAISED"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$LargePanelSystemAppraised = '0' ">
						<xsl:attribute name="LARGEPANELSYSTEMAPPRAISED"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$HistoricBldgRepairReq = '1' ">
						<xsl:attribute name="HISTORICBUILDINGREPAIRSREQ"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$HistoricBldgRepairReq = '0' ">
						<xsl:attribute name="HISTORICBUILDINGREPAIRSREQ"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$AsbestosPoorCondition = '1' ">
						<xsl:attribute name="ASBESTOSPOORCONDITION"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$AsbestosPoorCondition = '0' ">
						<xsl:attribute name="ASBESTOSPOORCONDITION"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="ESSENTIALMATTERS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@ESSENTIALMATTERS)"/></xsl:attribute>
					<xsl:attribute name="GENERALOBSERVATIONS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@GENERALOBSERVATIONS)"/></xsl:attribute>
					<xsl:if test="$ExtensionsOrAlterations = '1' ">
						<xsl:attribute name="EXTENSIONSORALTERATIONS"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$ExtensionsOrAlterations = '0' ">
						<xsl:attribute name="EXTENSIONSORALTERATIONS"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Tenanted = '1' ">
						<xsl:attribute name="TENANTEDPROPERTYIND"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Tenanted = '0' ">
						<xsl:attribute name="TENANTEDPROPERTYIND"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$SharedAccess = '1' ">
						<xsl:attribute name="UNADOPTEDSHAREDACCESSISSUES"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$SharedAccess = '0' ">
						<xsl:attribute name="UNADOPTEDSHAREDACCESSISSUES"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$MiningReport = '1' ">
						<xsl:attribute name="MININGREPORTSREQ"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$MiningReport = '0' ">
						<xsl:attribute name="MININGREPORTSREQ"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$NonResidentLandInd = '1' ">
						<xsl:attribute name="NONRESIDENTIALLANDIND"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$NonResidentLandInd = '0' ">
						<xsl:attribute name="NONRESIDENTIALLANDIND"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$DvlpProposals = '1' ">
						<xsl:attribute name="DEVELOPMENTPROPOSALS"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$DvlpProposals = '0' ">
						<xsl:attribute name="DEVELOPMENTPROPOSALS"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="SOLICITORNOTES"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@SOLICITORNOTES)"/></xsl:attribute>
					<xsl:if test="$NonSTDConstructionType = '1' ">
						<xsl:attribute name="NONSTANDARDCONSTRUCTIONTYPE"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$NonSTDConstructionType = '0' ">
						<xsl:attribute name="NONSTANDARDCONSTRUCTIONTYPE"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="CONSTRUCTIONTYPEDETAILS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@CONSTRUCTIONTYPEDETAILS)"/></xsl:attribute>
					<xsl:if test="$SingleConstruction = '1' ">
						<xsl:attribute name="ISSINGLESKINCONSTRUCTION"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$SingleConstruction = '0' ">
						<xsl:attribute name="ISSINGLESKINCONSTRUCTION"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$IsSingleTwoStorey = '1' ">
						<xsl:attribute name="ISSINGLESKINTWOSTOREY"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$IsSingleTwoStorey = '0' ">
						<xsl:attribute name="ISSINGLESKINTWOSTOREY"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyType = '10'">
						<xsl:attribute name="HOUSE"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyType = '20'">
						<xsl:attribute name="BLOW"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyType = '30'">
						<xsl:attribute name="FLAT"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyType = '99'">
						<xsl:attribute name="OTHERS"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>					
					<xsl:attribute name="YEARBUILT"><xsl:value-of select="concat('{',//GETLATESTVALUATIONREPORT/@YEARBUILT,'}')"/></xsl:attribute>
					<xsl:if test="$PropertyDescription = '1' ">
						<xsl:attribute name="DTACH"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDescription = '2' ">
						<xsl:attribute name="SEMI"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$PropertyDescription = '3' ">
						<xsl:attribute name="TERRACE"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$NewPropertyInd = '1' ">
						<xsl:attribute name="NEWPROPIND"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$NewPropertyInd = '0' ">
						<xsl:attribute name="NEWPROPIND"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="HOUSENO"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@BUILDINGORHOUSENUMBER)"/></xsl:attribute>
					<xsl:variable name="TotalYears" select="msg:GetYears(string(//GETLATESTVALUATIONREPORT/@YEARBUILT))"/>					
					<xsl:attribute name="TotalYears"><xsl:value-of select="string($TotalYears)"/></xsl:attribute>
					<xsl:if test="($NewPropertyInd = '1' or $TotalYears &lt; 10) ">		
							<xsl:if test="$CertificationType = '1' ">
								<xsl:attribute name="NHBC"><xsl:value-of select="string('X')"/></xsl:attribute>
							</xsl:if>					
							<xsl:if test="$CertificationType = '2' ">
								<xsl:attribute name="ZURICH"><xsl:value-of select="string('X')"/></xsl:attribute>
							</xsl:if>					
							<xsl:if test="$CertificationType = '3' ">
								<xsl:attribute name="PREMGUARANTEE"><xsl:value-of select="string('X')"/></xsl:attribute>
							</xsl:if>					
					</xsl:if>	
					<xsl:if test="$Tenure = '1' ">
						<xsl:attribute name="FHOLD"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Tenure = '2' ">
						<xsl:attribute name="LHOLD"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Tenure = '3' ">
						<xsl:attribute name="CHOLD"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Tenure = '4' ">
						<xsl:attribute name="FEUDHOLD"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="UNEXPIREDLEASE"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@UNEXPIREDLEASE)"/></xsl:attribute>
					<xsl:attribute name="LIVROOMS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@LIVINGROOMS)"/></xsl:attribute>
					<xsl:attribute name="BEDROOMS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@NUMBEROFBEDROOMS)"/></xsl:attribute>
					<xsl:attribute name="KITCHENS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@NUMBEROFKITCHENS)"/></xsl:attribute>
					<xsl:attribute name="BATHROOMS"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@BATHROOMS)"/></xsl:attribute>
					<xsl:attribute name="GARAGES"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@GARAGES)"/></xsl:attribute>
					<xsl:attribute name="PARKSPACES"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@PARKINGSPACES)"/></xsl:attribute>
					<xsl:attribute name="RESIDENCEAREA"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@RESIDENCEAREA)"/></xsl:attribute>
					<xsl:if test="$Conditions = '20' ">
						<xsl:attribute name="ABAVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Conditions = '10' ">
						<xsl:attribute name="AVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Conditions = '30' ">
						<xsl:attribute name="BLAVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Saleability = '20' ">
						<xsl:attribute name="SABAVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Saleability = '10' ">
						<xsl:attribute name="SAVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$Saleability = '30' ">
						<xsl:attribute name="SBLAVG"><xsl:value-of select="string('X')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="VALUERNAME"><xsl:value-of select="string(//VALUERINSTRUCTION/@VALUERNAME)"/></xsl:attribute>
					<xsl:if test="$ReTypeInd = '1' ">
						<xsl:attribute name="RETYPEIND"><xsl:value-of select="string('Y')"/></xsl:attribute>
					</xsl:if>
					<xsl:if test="$ReTypeInd = '0' ">
						<xsl:attribute name="RETYPEIND"><xsl:value-of select="string('N')"/></xsl:attribute>
					</xsl:if>
					<xsl:attribute name="APPNTDATE"><xsl:value-of select="string(//VALUERINSTRUCTION/@APPOINTMENTDATE)"/></xsl:attribute>
					<xsl:attribute name="DATERECEIVED"><xsl:value-of select="string(//GETLATESTVALUATIONREPORT/@DATERECEIVED)"/></xsl:attribute>
					<xsl:if test="$SignatureReturned = '2' ">					
						<xsl:attribute name="SIGNATURERETURNED"><xsl:value-of select="string('Signed On Original')"/></xsl:attribute>
					</xsl:if>	
					<xsl:if test="$SignatureReturned = '1' ">					
						<xsl:attribute name="SIGNATURERETURNED"><xsl:value-of select="string('')"/></xsl:attribute>
					</xsl:if>					
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
