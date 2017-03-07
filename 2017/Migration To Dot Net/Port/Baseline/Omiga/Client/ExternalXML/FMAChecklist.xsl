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
		var titlename2 ='';		
	    function SetSalutation(Title,FName,Surname,Counter)
	    {
			if (Counter == '1') 
			{
				if (Title !='')
					titlename= titlename  + Title ;
				if (FName !='')
					titlename= titlename  + ' ' + FName ;					
				if (Surname !='')
					titlename=titlename + ' ' + Surname;
			}
		  else if (Counter == '2') 
			{
				if (Title !='')
					titlename2= titlename2 + Title ;
				if (FName !='')
					titlename2= titlename2  + ' ' + FName;					
				if (Surname !='')
					titlename2=titlename2 + ' ' + Surname ;
			}	
		return(titlename);
	    }
        
             var Name='';
		function GetSalutation()
		{
			if ((titlename2 !='') && (titlename !=''))
			    Name= titlename  + ' and ' + titlename2;
			if ((titlename2 =='') && (titlename !=''))
			    Name= titlename;
			if ((titlename2 !='') && (titlename ==''))
			    Name= titlename2;
		     	
			return(Name);
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

  
		var strKYC1Status ='';
		var strKYC2Status ='';		
	    function SetKYCDetails(KYCStatus,Counter)
	    {
			if (Counter == '1') 
			{
				if (strKYC1Status !='')
				{

				}
				else
				{
					if ((KYCStatus == 'SATISFIED') || (KYCStatus == 'Satisfied') || (KYCStatus == 'satisfied'))
					{
						strKYC1Status='SATISFIED';	
					}	
					else
					{
						strKYC1Status=KYCStatus;						
					}	
				}	
			}
		  else if (Counter == '2') 
			{
				if (strKYC2Status !='')
				{

				}
				else
				{
					if ((KYCStatus == 'SATISFIED') || (KYCStatus == 'Satisfied') || (KYCStatus == 'satisfied'))
					{
						strKYC2Status = 'SATISFIED';	
					}	
					else
					{
						strKYC2Status = KYCStatus;						
					}	
				}	
			}
		return(strKYC1Status);
	    }
        
		function GetKYC1Status()
		{
			return(strKYC1Status);
		}
		
		function GetKYC2Status()
		{
			return(strKYC2Status);
		}

		

		var strKYC1AddStatus ='';		
		var strKYC2AddStatus ='';		
	    function SetKYCAddDetails(KYCAddStatus,Counter)
	    {
			if (Counter == '1') 
			{
				if (strKYC1AddStatus !='')
				{

				}
				else
				{
					if ((KYCAddStatus == 'OUTSTANDING') || (KYCAddStatus == 'Outstanding') || (KYCAddStatus == 'outstanding'))
					{
						strKYC1AddStatus='OUTSTANDING';	
					}	
					else
					{
						strKYC1AddStatus=KYCAddStatus;						
					}	
				}	
			}
		  else if (Counter == '2') 
			{
				if (strKYC2AddStatus !='')
				{

				}
				else
				{
					if ((KYCAddStatus == 'OUTSTANDING') || (KYCAddStatus == 'Outstanding') || (KYCAddStatus == 'outstanding'))
					{
						strKYC2AddStatus='OUTSTANDING';	
					}	
					else
					{
						strKYC2AddStatus=KYCAddStatus;						
					}	
				}	
			}
		return(strKYC1AddStatus);
	    }
        
		function GetKYC1AddStatus()
		{
			return(strKYC1AddStatus);
		}

		function GetKYC2AddStatus()
		{
			return(strKYC2AddStatus);
		}





		var strKYC1IDStatus ='';		
		var strKYC2IDStatus ='';		
	    function SetKYCIDDetails(KYCIDStatus,Counter)
	    {
			if (Counter == '1') 
			{
				if (strKYC1IDStatus !='')
				{

				}
				else
				{
					if ((KYCIDStatus == 'OUTSTANDING') || (KYCIDStatus == 'Outstanding') || (KYCIDStatus == 'outstanding'))
					{
						strKYC1IDStatus='OUTSTANDING';	
					}	
					else
					{
						strKYC1IDStatus=KYCIDStatus;						
					}	
				}	
			}
		  else if (Counter == '2') 
			{
				if (strKYC2IDStatus !='')
				{

				}
				else
				{
					if ((KYCIDStatus == 'OUTSTANDING') || (KYCIDStatus == 'Outstanding') || (KYCIDStatus == 'outstanding'))
					{
						strKYC2IDStatus='OUTSTANDING';	
					}	
					else
					{
						strKYC2IDStatus=KYCIDStatus;						
					}	
				}	
			}
		return(strKYC1IDStatus);
	    }
        
		function GetKYC1IDStatus()
		{
			return(strKYC1IDStatus);
		}

		function GetKYC2IDStatus()
		{
			return(strKYC2IDStatus);
		}


		var strIndKYCStatus='';
	    function SetIndividualKYCStatus(KYCStatus)
	    {
				if (strIndKYCStatus !='')
				{

				}
				else
				{
					if ((KYCStatus == 'SATISFIED') || (KYCStatus == 'Satisfied') || (KYCStatus == 'satisfied'))
					{
						strIndKYCStatus='SATISFIED';	
					}	
					else
					{
						strIndKYCStatus=KYCStatus;						
					}	
				}
		return(strIndKYCStatus);
	    }

         function GetIndividualKYCStatus(KYCStatus)
         {
			 return(strIndKYCStatus);
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
		
		
		var strCustomerNo1 ='';	    
		var strCustomerNo2 ='';	    		
	    function SetSeperateCustomerNo(CustomerNo,Counter)
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
		return(strCustomerNo1);
	    }
        
		function GetSeperateCustomerNo1()
		{
			return(strCustomerNo1);
		}		
		
		function GetSeperateCustomerNo2()
		{
			return(strCustomerNo2);
		}				
		
		
		var strEmpStatus1='';
	    var strValueID1='';
		var strEmpStatus2='';
	    var strValueID2='';	    
	    function SetEmpStatus(EmpStatus,ValueId,Counter)		
		{
			if (Counter == '1') 
			{
					if((ValueId == '10') && (EmpStatus == 'EMP'))
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}			
					else if(ValueId == '20' && EmpStatus == 'SELF')
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}
					else if((ValueId == '30') && (EmpStatus == 'E'))
					{
						strEmpStatus1='CON';
						strValueID1=ValueId;
					}
					
					else if((ValueId == '40') && (EmpStatus == 'N'))
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}			
					
					else if((ValueId == '50') && (EmpStatus == 'HM'))
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}			
					
					else if((ValueId == '60') && (EmpStatus == 'STU'))
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}			
					
					else if((ValueId == '70') && (EmpStatus == 'R'))
					{
						strEmpStatus1=EmpStatus;
						strValueID1=ValueId;
					}			
			}
			else if (Counter == '2') 
			{
					if((ValueId == '10') && (EmpStatus == 'EMP'))
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}			
					else if(ValueId == '20' && EmpStatus == 'SELF')
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}
					else if((ValueId == '30') && (EmpStatus == 'E'))
					{
						strEmpStatus2='CON';
						strValueID2=ValueId;
					}
					
					else if((ValueId == '40') && (EmpStatus == 'N'))
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}			
					
					else if((ValueId == '50') && (EmpStatus == 'HM'))
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}			
					
					else if((ValueId == '60') && (EmpStatus == 'STU'))
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}			
					
					else if((ValueId == '70') && (EmpStatus == 'R'))
					{
						strEmpStatus2=EmpStatus;
						strValueID2=ValueId;
					}			
			}
			return(strEmpStatus1);
		}

		function GetEmpStatus1()
		{
			return(strEmpStatus1);
		}
		
		function GetValueID1()
		{
			return(strValueID1);
		}
		
		function GetEmpStatus2()
		{
			return(strEmpStatus2);
		}
		
		function GetValueID2()
		{
			return(strValueID2);
		}


		var TotIncome1=0;
		var TotIncome2=0;
		var Income1=0;
		var Income2 = 0;		
		var VarIncome1 = 0;
		var VarIncome2 = 0;		
	    function SetApplicantOtherIncome(IncomeType,Amount,Counter)		
		{
			if (Counter == '1')
			{
				TotIncome1 =TotIncome1 + Amount;
				if(IncomeType == 'OTH')
				{
				  Income1= Income1 + Amount;
				}			
				else if(IncomeType == 'OV')
				{
				  VarIncome1= VarIncome1 + Amount;
				}
				else if(IncomeType == 'BP')
				{
				  VarIncome1= VarIncome1 + Amount;
				}				
				else if(IncomeType == 'CP')
				{
				  VarIncome1= VarIncome1 + Amount;
				}
			}
			else if (Counter == '2')
				TotIncome2 =TotIncome2 + Amount;
				if(IncomeType == 'OTH')
				{
				  Income2= Income2 + Amount;
				}			
				else if(IncomeType == 'OV')
				{
				  VarIncome2= VarIncome2 + Amount;
				}
				else if(IncomeType == 'BP')
				{
				  VarIncome2= VarIncome2 + Amount;
				}				
				else if(IncomeType == 'CP')
				{
				  VarIncome2= VarIncome2 + Amount;
				}															
			return(Income1);			
		}
		
		function GetApplicant1OtherIncome()
		{
			return(Income1);
		}

		function GetApplicant2OtherIncome()
		{
			return(Income2);
		}				

		function GetApplicant1VariableIncome()
		{
			return(VarIncome1);
		}		
		
		function GetApplicant2VariableIncome()
		{
			return(VarIncome2);
		}				
		
		function GetApplicant1TotIncome()
		{
			return(TotIncome1);
		}			
		
		function GetApplicant2TotIncome()
		{
			return(TotIncome2);
		}					
		
    ]]></msxsl:script>
    <xsl:variable name="SALUTATIONONE"/> 
    <xsl:variable name="Order_No"/> 
    <xsl:variable name="FeeType"/>    
    <xsl:variable name="CustomerKYCID"/>
    <xsl:variable name="KYC"/>
    <xsl:variable name="CustomerKYCStatus"/>        
    <xsl:variable name="CustomerKYCDetails"/>    
    <xsl:variable name="CustomerNo"/>
    <xsl:variable name="App1HasIncome"/>    
    <xsl:variable name="App2HasIncome"/>        
    <xsl:variable name="RegularVariIncome"/>
    <xsl:variable name="OtherIncome"/>    
    <xsl:variable name="EmploymentStatus1"/>    
    <xsl:variable name="EmploymentStatus2"/>    
    <xsl:variable name="EmpStatus1"/>    
    <xsl:variable name="EmpStatus2"/>	
    <xsl:variable name="CustomerNo1"/>    
    <xsl:variable name="CustomerNo2"/>    
    <xsl:variable name="Income1"/>
    <xsl:variable name="Income2"/>
    <xsl:variable name="EarnedIncome1"/>
    <xsl:variable name="EarnedIncome2"/>    
    <xsl:variable name="OtherIncome1"/>    
    <xsl:variable name="OtherIncome2"/>    
    <xsl:variable name="ValueId1"/>    
    <xsl:variable name="ValueId2"/>    
    <xsl:variable name="AppOtherIncome1"/>    
    <xsl:variable name="AppOtherIncome2"/>    
    <xsl:variable name="AppVariableIncome1"/>    
    <xsl:variable name="AppVariableIncome2"/>    
	<xsl:variable name="CustomerOrderNo"/>    
	<xsl:variable name="CustNo1"/>	
	<xsl:variable name="CustNo2"/>
	<xsl:variable name="Temp"/>	
	
	<xsl:variable name="CustomerKYCAddDetails"/>
	<xsl:variable name="CustomerKYCAddStatus"/>	
	<xsl:variable name="CustomerKYCIDDetails"/>		
	<xsl:variable name="CustomerKYCIDStatus"/>	
	
	<xsl:variable name="KYC1IDStatus"/>	
	<xsl:variable name="KYC2IDStatus"/>	
	<xsl:variable name="KYC1AddStatus"/>	
	<xsl:variable name="KYC2AddStatus"/>		
    <xsl:variable name="KYC1Status"/>	
    <xsl:variable name="KYC2Status"/>	
                
    <xsl:template match="/">
		<xsl:element name="TEMPLATEDATA">

			<xsl:for-each select="//CUSTOMERROLE">
					<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
					<xsl:if test="($Order_No = '1')">
							<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),1)"/>
					</xsl:if>		

					<xsl:if test="($Order_No = '2')">
							<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),2)"/>
					</xsl:if>
			</xsl:for-each>


            <xsl:variable name="TempCounter" select="msg:ResetCounter(2)"/>
	    <xsl:variable name="Salutation"/>
	    <xsl:variable name="Applicant"/>
       	    <xsl:variable name="NumberOfApplicants" select="string(//TEMPLATEDATA/ADDRESS/@CUSTOMERSATADDRESS)"/>
            <xsl:attribute name="SALUTATION">
			<xsl:value-of select="msg:GetSalutation()"/>
      	    </xsl:attribute>
	    <xsl:attribute name="CURRENTDATE">
		<xsl:value-of select="msg:GetDate()"/>
	    </xsl:attribute>
            <xsl:attribute name="APPLICATIONNUMBER">
                <xsl:value-of select="concat('{', string(//CUSTOMERROLE/@APPLICATIONNUMBER), '}')"/>
            </xsl:attribute>
	    <xsl:attribute name="CUSTOMERNUMBER">
		<xsl:value-of select="msg:GetCustomerNo()"/>
            </xsl:attribute>									
       	    <xsl:attribute name="APPLICANT">
               	<xsl:for-each select="//CUSTOMERROLE/CUSTOMERVERSION">
                    <xsl:variable name="AppCounter" select="msg:UpdateCounter(2)"/>
       	            <xsl:variable name="Title" select="@TITLE_TEXT"/>
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


	    <xsl:if test="($NumberOfApplicants = '1')">
		    <xsl:element name="ONECUSTOMER"/>
	    </xsl:if>
	    <xsl:if test="($NumberOfApplicants = '2')">
		    <xsl:element name="TWOAPPLICANTS"/>
		    <xsl:element name="TWOCUSTOMERS"/>
	    </xsl:if>
	    <xsl:variable name="InactiveStatus" select="string(//APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/PANEL/PANELLEGALREP/@STATUS)"/>
	    <xsl:variable name="LegalPanel" select="count(//APPLICATIONLEGALREP/NAMEANDADDRESSDIRECTORY/PANEL/PANELLEGALREP)"/>
	    <xsl:variable name="ValuationInstruct" select="string(//VALUERINSTRUCTION/@DATEOFINSTRUCTION)"/>

            <xsl:if test="($LegalPanel > 0 or $InactiveStatus = '20')">
		    <xsl:element name="SOLICITOR"/>
	    </xsl:if>
	    <xsl:if test="(($LegalPanel > 0 or $InactiveStatus = '20') or ($ValuationInstruct !=''))">
		    <xsl:element name="THIS"/>
	    </xsl:if>


       	    <xsl:variable name="CustCount" select="string(//TEMPLATEDATA/ADDRESS/@CUSTOMERSATADDRESS)"/>
            <xsl:attribute name="CUSTOMERNUMBER1">
		<xsl:value-of select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[1]/@CUSTOMERNUMBER)"/>
	    </xsl:attribute>
	    <xsl:if test="(CustCount > '1')">
            	<xsl:attribute name="CUSTOMERNUMBER2">
		    <xsl:value-of select="string(//TEMPLATEDATA/ADDRESS/CUSTOMER[2]/@CUSTOMERNUMBER)"/>
	        </xsl:attribute>
	    </xsl:if>












			<xsl:variable name="CustRole" select="count(RESPONSE/FMACHECKLIST/CUSTOMERROLE)"/>
			
			<xsl:for-each select="//CUSTOMERROLE">
					<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
					<xsl:if test="($Order_No = '1')">
							<xsl:variable name="CustomerKYCDetails" select="msg:SetKYCDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCSTATUS),1)"/>									
							<xsl:variable name="CustomerKYCAddDetails" select="msg:SetKYCAddDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCADDRESSFLAG),1)"/>
							<xsl:variable name="CustomerKYCIDDetails" select="msg:SetKYCIDDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCIDFLAG),1)"/>
							<xsl:variable name="CustomerNo1" select="msg:SetSeperateCustomerNo(string(.//CUSTOMERVERSION/@CUSTOMERNUMBER),1)"/>									
							<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),1)"/>
					</xsl:if>		

					<xsl:if test="($Order_No = '2')">
							<xsl:variable name="CustomerNo2" select="msg:SetSeperateCustomerNo(string(.//CUSTOMERVERSION/@CUSTOMERNUMBER),2)"/>									
							<xsl:variable name="CustomerKYCDetails" select="msg:SetKYCDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCSTATUS),2)"/>		
							<xsl:variable name="CustomerKYCAddDetails" select="msg:SetKYCAddDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCADDRESSFLAG),1)"/>
							<xsl:variable name="CustomerKYCIDDetails" select="msg:SetKYCIDDetails(string(.//CUSTOMERVERSION/@CUSTOMERKYCIDFLAG),1)"/>						
							<xsl:variable name="SALUTATIONONE" select="msg:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TEXT),string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),2)"/>
					</xsl:if>
			</xsl:for-each>
			
					
			<xsl:variable name="CustNo1" select="msg:GetSeperateCustomerNo1()"/>
			<xsl:variable name="CustNo2" select="msg:GetSeperateCustomerNo2()"/>			
			
			<xsl:for-each select="//EMPLOYMENT/EMPLOYEDDETAILS/EARNEDINCOME">
					<xsl:if test="($CustNo1 = string(.//@CUSTOMERNUMBER))">
							<xsl:variable name="Income1" select="msg:SetApplicantOtherIncome(string(.//@EARNEDINCOMETYPE_TYPES),number(.//@EARNEDINCOMEAMOUNT),1)"/>
					</xsl:if>
					<xsl:if test="($CustNo2 = string(.//@CUSTOMERNUMBER))">
							<xsl:variable name="Income2" select="msg:SetApplicantOtherIncome(string(.//@EARNEDINCOMETYPE_TYPES),number(.//@EARNEDINCOMEAMOUNT),2)"/>
					</xsl:if>								
			</xsl:for-each>

			<xsl:for-each select="//EMPLOYMENT/COMBOVALIDATION">
					<xsl:if test="($CustNo1 = string(..//@CUSTOMERNUMBER))">
						<xsl:variable name="EmploymentStatus1" select="msg:SetEmpStatus(string(.//@VALIDATIONTYPE),string(.//@VALUEID),1)"/>
					</xsl:if>
					<xsl:if test="($CustNo2 = string(..//@CUSTOMERNUMBER))">
						<xsl:variable name="EmploymentStatus2" select="msg:SetEmpStatus(string(.//@VALIDATIONTYPE),string(.//@VALUEID),2)"/>
					</xsl:if>								
			</xsl:for-each>								

			<xsl:variable name="KYC1AddStatus" select="msg:GetKYC1AddStatus()"/>
			<xsl:variable name="KYC2AddStatus" select="msg:GetKYC2AddStatus()"/>			
			
			<xsl:variable name="KYC1IDStatus" select="msg:GetKYC1IDStatus()"/>
			<xsl:variable name="KYC2IDStatus" select="msg:GetKYC2IDStatus()"/>
			
			<xsl:variable name="KYC1Status" select="msg:GetKYC1Status()"/>			
			<xsl:variable name="KYC2Status" select="msg:GetKYC2Status()"/>		
			
			<xsl:variable name="App1HasIncome" select="msg:GetApplicant1TotIncome()"/>
			<xsl:variable name="App2HasIncome" select="msg:GetApplicant2TotIncome()"/>			
			
			<xsl:variable name="EmpStatus1" select="msg:GetEmpStatus1()"/>
			<xsl:variable name="ValueId1" select="msg:GetValueID1()"/>
			<xsl:variable name="AppOtherIncome1" select="msg:GetApplicant1OtherIncome()"/>														
			<xsl:variable name="AppVariableIncome1" select="msg:GetApplicant1VariableIncome()"/>			
			
			<xsl:variable name="EmpStatus2" select="msg:GetEmpStatus2()"/>
			<xsl:variable name="ValueId2" select="msg:GetValueID2()"/>
			<xsl:variable name="AppOtherIncome2" select="msg:GetApplicant2OtherIncome()"/>														
			<xsl:variable name="AppVariableIncome2" select="msg:GetApplicant2VariableIncome()"/>			
						
				<xsl:for-each select="//CUSTOMERROLE">
								<xsl:variable name="CustomerOrderNo" select=".//@CUSTOMERORDER"/>
								<xsl:attribute name="APPLICANTS"><xsl:value-of select="string(msg:GetSalutation())"/></xsl:attribute>								
								
								<xsl:attribute name="InactiveStatus"><xsl:value-of select="string($InactiveStatus)"/></xsl:attribute>
								<xsl:attribute name="LegalPanel"><xsl:value-of select="string($LegalPanel)"/></xsl:attribute>								

								<xsl:attribute name="ValuationInstruct"><xsl:value-of select="string($ValuationInstruct)"/></xsl:attribute>

								<xsl:attribute name="Income1"><xsl:value-of select="msg:GetApplicant1OtherIncome()"/></xsl:attribute>
								<xsl:attribute name="Income2"><xsl:value-of select="msg:GetApplicant2OtherIncome()"/></xsl:attribute>																

								<xsl:attribute name="Applicant1VariableIncome"><xsl:value-of select="msg:GetApplicant1VariableIncome()"/></xsl:attribute>
								<xsl:attribute name="Applicant2VariableIncome"><xsl:value-of select="msg:GetApplicant2VariableIncome()"/></xsl:attribute>
																	
								<xsl:attribute name="AppOtherIncome1"><xsl:value-of select="$AppOtherIncome1"/></xsl:attribute>
								<xsl:attribute name="AppOtherIncome2"><xsl:value-of select="$AppOtherIncome2"/></xsl:attribute>								
								
								<xsl:attribute name="EmploymentStatus1"><xsl:value-of select="string(msg:GetEmpStatus1())"/></xsl:attribute>
								<xsl:attribute name="EmploymentStatus2"><xsl:value-of select="string(msg:GetEmpStatus2())"/></xsl:attribute>
		
								<xsl:attribute name="ValueId1"><xsl:value-of select="msg:GetValueID1()"/></xsl:attribute>								
								<xsl:attribute name="ValueId2"><xsl:value-of select="msg:GetValueID2()"/></xsl:attribute>								
								
								<xsl:attribute name="App1HasIncome"><xsl:value-of select="msg:GetApplicant1TotIncome()"/></xsl:attribute>
								<xsl:attribute name="App2HasIncome"><xsl:value-of select="msg:GetApplicant2TotIncome()"/></xsl:attribute>	
					
								<xsl:attribute name="KYC1AddStatus"><xsl:value-of select="($KYC1AddStatus)"/></xsl:attribute>
								<xsl:attribute name="KYC2AddStatus"><xsl:value-of select="($KYC2AddStatus)"/></xsl:attribute>								

								<xsl:attribute name="KYC1IDStatus"><xsl:value-of select="($KYC1IDStatus)"/></xsl:attribute>
								<xsl:attribute name="KYC2IDStatus"><xsl:value-of select="($KYC2IDStatus)"/></xsl:attribute>
								
								<xsl:attribute name="KYC1Status"><xsl:value-of select="($KYC1Status)"/></xsl:attribute>
								<xsl:attribute name="KYC2Status"><xsl:value-of select="($KYC2Status)"/></xsl:attribute>	
															
								<xsl:attribute name="CustNo1"><xsl:value-of select="($CustNo1)"/></xsl:attribute>
								<xsl:attribute name="CustNo2"><xsl:value-of select="($CustNo2)"/></xsl:attribute>		

								<xsl:choose>
									<xsl:when test="($ValuationInstruct !='')">
											<xsl:element name="VALUATIONINSTRUCT">
													<xsl:variable name="FeeType" select="string(//APPLICATION/APPLICATIONFEETYPE/FEEPAYMENT/@FEETYPE)"/>
													<xsl:attribute name="FeeType"><xsl:value-of select="string($FeeType)"/></xsl:attribute>
													<xsl:choose>
														<xsl:when test="($FeeType = '3')">
															<xsl:attribute name="VALUATIONFEE">
																<xsl:value-of select=".//APPLICATION/APPLICATIONFEETYPE/FEEPAYMENT/@REFUNDAMOUNT"/>
															</xsl:attribute>					
														</xsl:when>
														<xsl:when test="($FeeType != '3')">
															<xsl:attribute name="VALUATIONFEE">
																<xsl:value-of select="0"/>
															</xsl:attribute>															
														</xsl:when>														
													</xsl:choose>		
											</xsl:element>
									</xsl:when>
								</xsl:choose>
								<xsl:choose>
									<xsl:when  test="($CustRole > 1)">
										    <xsl:if test="(($KYC1Status != 'SATISFIED'  or $KYC1Status = '') or ($KYC2Status != 'SATISFIED'  or $KYC2Status = ''))">
												<xsl:element name="KYCCHECK"/>
												<xsl:element name="KYCCHECKREQ">
														<xsl:if  test="($CustRole > 1)">
															<xsl:element name="TWOAPPLICANTS"/>
														</xsl:if>
												</xsl:element>
										   </xsl:if>
									</xsl:when>	   
								</xsl:choose>
								<xsl:choose>	
									<xsl:when  test="($CustRole = 1)">
										    <xsl:if test="($KYC1Status != 'SATISFIED'  or $KYC1Status = '')">
												<xsl:element name="KYCCHECK"/>
												<xsl:element name="KYCCHECKREQ">
														<xsl:if  test="($CustRole > 1)">
															<xsl:element name="TWOAPPLICANTS"/>
														</xsl:if>
												</xsl:element>
										   </xsl:if>
									</xsl:when>									
								</xsl:choose>			
								<xsl:choose>
									<xsl:when test="($CustRole > 1)">
										<xsl:element name="TWOAPPLICANTS"/>
									</xsl:when>
								</xsl:choose>		

								<xsl:choose>
									<xsl:when test="($CustRole = 1)">
													 <xsl:if test="($App1HasIncome = 0)">
															<xsl:element name="APP1INCOME1"/>
													 </xsl:if>
													 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1='0'))">
															<xsl:element name="APP1INCOME2"/>
													 </xsl:if>
													 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1 > '0'))">
															<xsl:element name="APP1INCOME3"/>
													 </xsl:if>														 
													 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1='0'))">
															<xsl:element name="APP1INCOME4"/>
													 </xsl:if>														 														 
													 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1='0'))">
															<xsl:element name="APP1INCOME5"/>
													 </xsl:if>																 
													 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1='0'))">
															<xsl:element name="APP1INCOME6"/>
													 </xsl:if>
													 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1 > '0'))">
															<xsl:element name="APP1INCOME7"/>
													 </xsl:if>
													 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1 > '0'))">
															<xsl:element name="APP1INCOME8"/>
													 </xsl:if>	
													 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1 > '0'))">
															<xsl:element name="APP1INCOME9"/>
													 </xsl:if>												
													 <xsl:if test="((($EmpStatus1 = 'N' and $ValueId1 = '40') or ($EmpStatus1 = 'HM' and $ValueId1 = '50') or ($EmpStatus1 = 'STU' and $ValueId1 = '60')) 
																 and ($App1HasIncome > '0'))">
															<xsl:element name="APP1INCOME10"/>														 				 																									 
													</xsl:if>		
													 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1IDStatus ='OUTSTANDING' or $KYC1IDStatus = ''))">
															<xsl:element name="APP1VERIFYID1"/>
													</xsl:if>		
													 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1AddStatus ='OUTSTANDING' or $KYC1AddStatus = ''))">
															<xsl:element name="APP1VERIFYADD1"/>
													</xsl:if>		
								    </xsl:when>
								    
									<xsl:when test="($CustRole > 1)">
										<xsl:element name="TWOCUSTOMERS">
	
											<xsl:attribute name="CustomerOrderNo"><xsl:value-of select="string($CustomerOrderNo)"/></xsl:attribute>												
											<xsl:choose>
												<xsl:when test="($CustomerOrderNo = '1')">
																<xsl:attribute name="EmpStatus1"><xsl:value-of select="string($EmpStatus1)"/></xsl:attribute>					
																<xsl:attribute name="ValueId1"><xsl:value-of select="string($ValueId1)"/></xsl:attribute>
																<xsl:attribute name="AppOtherIncome1"><xsl:value-of select="string($AppOtherIncome1)"/></xsl:attribute>
																<xsl:attribute name="AppVariableIncome1"><xsl:value-of select="string($AppVariableIncome1)"/></xsl:attribute>
																 <xsl:if test="($App1HasIncome = 0)">
																		<xsl:element name="APP1INCOME1"/>
																 </xsl:if>
																  <xsl:if test="($App2HasIncome = 0)">
																		<xsl:element name="APP2INCOME11"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1='0'))">
																		<xsl:element name="APP1INCOME2"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1 > '0'))">
																		<xsl:element name="APP1INCOME3"/>
																 </xsl:if>														 
																 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1='0'))">
																		<xsl:element name="APP1INCOME4"/>
																 </xsl:if>														 														 
																 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1='0'))">
																		<xsl:element name="APP1INCOME5"/>
																 </xsl:if>																 
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1='0'))">
																		<xsl:element name="APP1INCOME6"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1 > '0'))">
																		<xsl:element name="APP1INCOME7"/>
																 </xsl:if>
																 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1 > '0'))">
																		<xsl:element name="APP1INCOME8"/>
																 </xsl:if>	
																 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1 > '0'))">
																		<xsl:element name="APP1INCOME9"/>
																 </xsl:if>												
																 <xsl:if test="((($EmpStatus1 = 'N' and $ValueId1 = '40') or ($EmpStatus1 = 'HM' and $ValueId1 = '50') or ($EmpStatus1 = 'STU' and $ValueId1 = '60')) 
																			 and ($App1HasIncome > '0'))">
																		<xsl:element name="APP1INCOME10"/>
																</xsl:if>		
																 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1IDStatus ='OUTSTANDING' or $KYC1IDStatus = ''))">
																		<xsl:element name="APP1VERIFYID1"/>
																</xsl:if>		
																 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1AddStatus ='OUTSTANDING' or $KYC1AddStatus = ''))">
																		<xsl:element name="APP1VERIFYADD1"/>
																</xsl:if>		

																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2='0') and 		($AppVariableIncome2='0'))">
																		<xsl:element name="APP2INCOME12"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2='0') and 		($AppVariableIncome2 > '0'))">
																		<xsl:element name="APP2INCOME13"/>
																 </xsl:if>														 
																 <xsl:if test="(($EmpStatus2 = 'SELF' and $ValueId2 = '20') and ($AppOtherIncome2='0'))">
																		<xsl:element name="APP2INCOME14"/>
																 </xsl:if>														 														 
																 <xsl:if test="(($EmpStatus2 = 'R' and $ValueId2 = '70') and ($AppOtherIncome2='0'))">
																		<xsl:element name="APP2INCOME15"/>
																 </xsl:if>																 
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2 > '0') and 		($AppVariableIncome2='0'))">
																		<xsl:element name="APP2INCOME16"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2 > '0') and 		($AppVariableIncome2 > '0'))">
																		<xsl:element name="APP2INCOME17"/>
																 </xsl:if>
																 <xsl:if test="(($EmpStatus2 = 'SELF' and $ValueId2 = '20') and ($AppOtherIncome2 > '0'))">
																		<xsl:element name="APP2INCOME18"/>
																 </xsl:if>	
																 <xsl:if test="(($EmpStatus2 = 'R' and $ValueId2 = '70') and ($AppOtherIncome2 > '0'))">
																		<xsl:element name="APP2INCOME19"/>
																 </xsl:if>												
																 <xsl:if test="((($EmpStatus2 = 'N' and $ValueId2 = '40') or ($EmpStatus2 = 'HM' and $ValueId2 = '50') or ($EmpStatus2 = 'STU' and $ValueId2 = '60')) 
																			 and ($App2HasIncome > '0'))">
																		<xsl:element name="APP2INCOME20"/>
																</xsl:if>		
																 <xsl:if test="(($KYC2Status != 'SATISFIED') and ($KYC2IDStatus ='OUTSTANDING' or $KYC2IDStatus = ''))">
																		<xsl:element name="APP1VERIFYID1"/>
																</xsl:if>		
																 <xsl:if test="(($KYC2Status != 'SATISFIED') and ($KYC2AddStatus ='OUTSTANDING' or $KYC2AddStatus = ''))">
																		<xsl:element name="APP1VERIFYADD1"/>
																</xsl:if>														
												</xsl:when>
												<xsl:when test="($CustomerOrderNo = '2')">
																<xsl:attribute name="EmpStatus2"><xsl:value-of select="string($EmpStatus2)"/></xsl:attribute>					
																<xsl:attribute name="ValueId2"><xsl:value-of select="string($ValueId2)"/></xsl:attribute>
																<xsl:attribute name="AppOtherIncome2"><xsl:value-of select="string($AppOtherIncome2)"/></xsl:attribute>
																<xsl:attribute name="AppVariableIncome2"><xsl:value-of select="string($AppVariableIncome2)"/></xsl:attribute>
																 <xsl:if test="($App1HasIncome = 0)">
																		<xsl:element name="APP1INCOME1"/>
																 </xsl:if>
																  <xsl:if test="($App2HasIncome = 0)">
																		<xsl:element name="APP2INCOME11"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1='0'))">
																		<xsl:element name="APP1INCOME2"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1='0') and 		($AppVariableIncome1 > '0'))">
																		<xsl:element name="APP1INCOME3"/>
																 </xsl:if>														 
																 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1='0'))">
																		<xsl:element name="APP1INCOME4"/>
																 </xsl:if>														 														 
																 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1='0'))">
																		<xsl:element name="APP1INCOME5"/>
																 </xsl:if>																 
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1='0'))">
																		<xsl:element name="APP1INCOME6"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus1 = 'EMP' and $ValueId1 = '10') or ($EmpStatus1 = 'CON' and $ValueId1 = '30')) and ($AppOtherIncome1 > '0') and 		($AppVariableIncome1 > '0'))">
																		<xsl:element name="APP1INCOME7"/>
																 </xsl:if>
																 <xsl:if test="(($EmpStatus1 = 'SELF' and $ValueId1 = '20') and ($AppOtherIncome1 > '0'))">
																		<xsl:element name="APP1INCOME8"/>
																 </xsl:if>	
																 <xsl:if test="(($EmpStatus1 = 'R' and $ValueId1 = '70') and ($AppOtherIncome1 > '0'))">
																		<xsl:element name="APP1INCOME9"/>
																 </xsl:if>												
																 <xsl:if test="((($EmpStatus1 = 'N' and $ValueId1 = '40') or ($EmpStatus1 = 'HM' and $ValueId1 = '50') or ($EmpStatus1 = 'STU' and $ValueId1 = '60')) 
																			 and ($App1HasIncome > '0'))">
																		<xsl:element name="APP1INCOME10"/>
																</xsl:if>		
																 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1IDStatus ='OUTSTANDING' or $KYC1IDStatus = ''))">
																		<xsl:element name="APP1VERIFYID2"/>
																</xsl:if>		
																 <xsl:if test="(($KYC1Status != 'SATISFIED') and ($KYC1AddStatus ='OUTSTANDING' or $KYC1AddStatus = ''))">
																		<xsl:element name="APP1VERIFYADD2"/>
																</xsl:if>

																					
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2='0') and 		($AppVariableIncome2='0'))">
																		<xsl:element name="APP2INCOME12"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2='0') and 		($AppVariableIncome2 > '0'))">
																		<xsl:element name="APP2INCOME13"/>
																 </xsl:if>														 
																 <xsl:if test="(($EmpStatus2 = 'SELF' and $ValueId2 = '20') and ($AppOtherIncome2='0'))">
																		<xsl:element name="APP2INCOME14"/>
																 </xsl:if>														 														 
																 <xsl:if test="(($EmpStatus2 = 'R' and $ValueId2 = '70') and ($AppOtherIncome2='0'))">
																		<xsl:element name="APP2INCOME15"/>
																 </xsl:if>																 
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2 > '0') and 		($AppVariableIncome2='0'))">
																		<xsl:element name="APP2INCOME16"/>
																 </xsl:if>
																 <xsl:if test="((($EmpStatus2 = 'EMP' and $ValueId2 = '10') or ($EmpStatus2 = 'CON' and $ValueId2 = '30')) and ($AppOtherIncome2 > '0') and 		($AppVariableIncome2 > '0'))">
																		<xsl:element name="APP2INCOME17"/>
																 </xsl:if>
																 <xsl:if test="(($EmpStatus2 = 'SELF' and $ValueId2 = '20') and ($AppOtherIncome2 > '0'))">
																		<xsl:element name="APP2INCOME18"/>
																 </xsl:if>	
																 <xsl:if test="(($EmpStatus2 = 'R' and $ValueId2 = '70') and ($AppOtherIncome2 > '0'))">
																		<xsl:element name="APP2INCOME19"/>
																 </xsl:if>												
																 <xsl:if test="((($EmpStatus2 = 'N' and $ValueId2 = '40') or ($EmpStatus2 = 'HM' and $ValueId2 = '50') or ($EmpStatus2 = 'STU' and $ValueId2 = '60')) 
																			 and ($App2HasIncome > '0'))">
																		<xsl:element name="APP2INCOME20"/>
																</xsl:if>		
																 <xsl:if test="(($KYC2Status != 'SATISFIED') and ($KYC2IDStatus ='OUTSTANDING' or $KYC2IDStatus = ''))">
																		<xsl:element name="APP1VERIFYID2"/>
																</xsl:if>		
																 <xsl:if test="(($KYC2Status != 'SATISFIED') and ($KYC2AddStatus ='OUTSTANDING' or $KYC2AddStatus = ''))">
																		<xsl:element name="APP1VERIFYADD2"/>
																</xsl:if>														
												</xsl:when>
											</xsl:choose>
										</xsl:element>
								    </xsl:when>								    
								</xsl:choose>
				</xsl:for-each>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>
