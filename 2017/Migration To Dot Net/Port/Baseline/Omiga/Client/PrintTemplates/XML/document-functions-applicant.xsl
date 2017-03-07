<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msg="http://marlborough-stirling.com/xpath-functions" xmlns:msgint="http://marlborough-stirling.com/xpath-functions" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:ms="urn:schemas-microsoft-com:xslt">
	<!--===============================================================================================================
	EPSOM History
	===============================================================================================================
	Prg		Date			AQR			Description
	OS		11/01/2007	EP2_469		Created for use by Applicant templates
	DS		11/01/2007	EP2_739		Added OFFERISSUEDDATE attribute
	PB		12/01/2007	EP2_734		Added GetSingleOrJointSalutation and GetSingleOrJointShortSalutation - see comments for more detail
	OS		16/01/2007	EP2_800		Added call to "Applicant" template for getting "Application name"
	OS		19/01/2007	EP2_799		Added PROPERTYADDRESS attribute
	PB		21/01/2007	EP2_806		Added ApplicantOrGuarantor() function
	DS		23/01/2007	EP2_737		Added attributes at Template node and functions for REVISEDOFFERCOVERAPPLICANT
	DS		06/02/2007	EP2_737		Modified DealWithValidationText function and formatted dates
	DS		09/02/2007	EP2_739   	Updated OFFERISSUEDDATE attribute and formatted date
	OS		14/02/2007	EP2_813		Updated OFFERISSUEDDATE attribute according to new specs
	OS		15/02/2007	EP2_799		Updated PROPERTYADDRESS according to updated specs
	DS		21/02/2007	EP2_737		Modified Applicant name and added DealWithValidationText for joint applicant

	===============================================================================================================-->
	<xsl:import href="document-functions.xsl"/>
	<!--===============================================================================================================-->
	
	<msxsl:script language="JavaScript" implements-prefix="msgint"><![CDATA[
	
		// PB 21/01/2007 EP2_806
		function ApplicantOrGuarantor( Type )
		{
			var Answer = Type;
			if( Type=='A' )
			{
				Answer='Applicant'
			}
			if( Type=='G' )
			{
				Answer='Guarantor'
			}
			return( Answer );
		}
			
		// PB 12/01/2007 EP2_734
		// This function can be called with either 1 or 2 applicants to return the correct salutation WITH the first initial
		function GetSingleOrJointSalutation( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 )
		{
			var Salutation='';
			
			if( Title2 != '' )
			{
				//Joint
				Salutation = DealWithJointTitle( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 );
			}
			else
			{
				//Single
				Salutation = GetApplicantNameWithInitials(Title1, TitleOther1, Forename1, Surname1);
			}
			return( Salutation )
		}
		
		// PB 12/01/2007 EP2_734
		// This function can be called with either 1 or 2 applicants to return the correct salutation WITHOUT the first initial
		function GetSingleOrJointShortSalutation( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 )
		{
			var Salutation='';
			
			if( Title2 != '' )
			{
				//Joint
				Salutation = DealWithJointSalutation( Title1, TitleOther1, Forename1, Surname1, Title2, TitleOther2, Forename2, Surname2 );
			}
			else
			{
				//Single
				Salutation = DealWithSalutation(Title1, TitleOther1, Forename1, Surname1);
			}
			return( Salutation )
		}
	
		function DealWithJointSalutation(Title1, TitleOther1, FirstName1, Surname1, Title2, TitleOther2, FirstName2, Surname2)
    {
		var strSalutation = "";
		
		if (Title1 != '' && Title2 != '' && Surname1 == Surname2)
		{
			if ((Title1 == 'Mr' && Title2 == 'Mrs') || (Title1 == 'Mrs' && Title2 == 'Mr'))
			{
				strSalutation = DealWithSalutation('Mr and Mrs', '', '', Surname1);
			}
			else
			{
				strSalutation = DealWithSalutation(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + DealWithSalutation(Title2, TitleOther2, FirstName2, Surname2);
			}
		}
		else
		{
			strSalutation = DealWithSalutation(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + DealWithSalutation(Title2, TitleOther2, FirstName2, Surname2);
		}
			
		return (strSalutation);
    }
    
    function DealWithJointTitle(Title1, TitleOther1, FirstName1, Surname1, Title2, TitleOther2, FirstName2, Surname2)
    {
		var strAddressTitle = "";
		
		if (Title1 != '' && Title2 != '' && Surname1 == Surname2)
		{
			if (Title1 == 'Mr' && Title2 == 'Mrs')
			{
				strAddressTitle = GetApplicantNameWithInitials('Mr and Mrs', '',FirstName1, Surname1);
			}
			else if (Title1 == 'Mrs' && Title2 == 'Mr')
			{
				strAddressTitle = GetApplicantNameWithInitials('Mr and Mrs', '',FirstName2, Surname2);
			}
			else
			{
				strAddressTitle = GetApplicantNameWithInitials(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + GetApplicantNameWithInitials(Title2, TitleOther2, FirstName2, Surname2);
			}
		}
		else
		{
			strAddressTitle = GetApplicantNameWithInitials(Title1, TitleOther1, FirstName1, Surname1) + ' and ' + GetApplicantNameWithInitials(Title2, TitleOther2, FirstName2, Surname2);
		}
			
		return (strAddressTitle);
    }
   
    function DealWithApplicantAddress(FlatNumber, HouseName, HouseNumber, Street, District, Town, County, PostCode)
    	{
		var strAddress = "";
		

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

	function DealWithValidationText(typeOfApplication)
	{
		var returnTypeOfApplication = "Other"
		var typesArray = typeOfApplication.split('|')
		for( ctr=0 ; ctr < typesArray.length; ctr++)
		{
			if (typesArray[ctr]  == "FT" || typesArray[ctr]  == "HM"  || typesArray[ctr]  == "R"  || typesArray[ctr]  == "RTB" || typesArray[ctr]  == "TOE" || typesArray[ctr]  == "NP"   )
			{
				returnTypeOfApplication = "MAINADVORTOE"
				break
			}
			
			else if (typesArray[ctr]  == "ABO" )
			{
				 if (typesArray[ctr+1]  == "ABTOE" )
				{
					returnTypeOfApplication = "MAINADVORTOE"
					break
				}
				else
				{
					returnTypeOfApplication = "ADDBRWGORPRODSW1"
					break
				}
			}
			
			else if (typesArray[ctr]  == "TOE" )
			{
				returnTypeOfApplication = "MAINADVORTOE"
				break
			}
			
			else if (typesArray[ctr]  == "PSW" )
			{
				returnTypeOfApplication = "ADDBRWGORPRODSW1"
				break
			}
			
			else
			{
				returnTypeOfApplication = "Mortgage"
			}
		}
		return (returnTypeOfApplication)
	}
	
	function DealWithValidationText_old(ValidationType)
	{
		var returnText = "Other";
		
			
			if (ValidationType == "ABO" || ValidationType == "PSW")
			{
				returnText = "ADDBRWGORPRODSW1";
				
			}
			
			else if (ValidationType  == "FT" || ValidationType  == "HM"  || ValidationType  == "R"  || ValidationType  == "RTB" || ValidationType  == "ABTOE"|| ValidationType == "TOE"|| ValidationType  == "NP" )
			{
				returnText = "MAINADVORTOE";
				
			}
			
			
		return (returnText);
	}	
	
	
	function DealWithApplicationNumber(AppNumber)
	{
		var returnText = "Other";
		
			
			
				returnText = '{' + AppNumber + '}' ;
				
			
			
			
		return (returnText);
	}	
	
	function GetApplicantNames(FirstName1, Surname1,  FirstName2, Surname2)
    {
		var strApplicantNames = "";
		
		if (FirstName1 != '' && Surname1 != '' )
		{
			
			strApplicantNames = FirstName1+ ' ' + Surname1;
			
			
		}
		
		if (FirstName2 != '' && Surname2 != '' )
		{
		
			strApplicantNames = strApplicantNames + ' and ' + FirstName2+ ' ' + Surname2;
			
		}
	
			
		return (strApplicantNames);
    }
		
	
    ]]></msxsl:script>
	<!--============================================================================================================-->
	<xsl:template name="APPLICANTINFO">
		<xsl:param name="RESPONSE"/>
			<xsl:variable name="CustCount" select="count($RESPONSE/UNIQUEAPPLICATIONADDRESSES)"/>
			<xsl:variable name="CustRole" select="count($RESPONSE/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'])"/>
			<xsl:variable name="ApplicationType"><xsl:value-of select="msg:DealWithTypeOfApplication(string($RESPONSE/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))"/></xsl:variable>
			<!--DS - EP2_737-->
			<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="msg:DealWithApplicationNumber(string($RESPONSE/APPLICATIONFACTFIND/@APPLICATIONNUMBER))"/></xsl:attribute>
		   <xsl:attribute name="APPLICANTNAME1"><xsl:value-of select="concat(string(/RESPONSE/REVISEDOFFERCOVERAPPLICANT/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@FIRSTFORENAME) , ' ' , string(/RESPONSE/REVISEDOFFERCOVERAPPLICANT/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="APPLICANTNAME2"><xsl:value-of select="concat(string(/RESPONSE/REVISEDOFFERCOVERAPPLICANT/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@FIRSTFORENAME) , ' ' , string(/RESPONSE/REVISEDOFFERCOVERAPPLICANT/CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
			<xsl:attribute name="APPLICANTNAMES"><xsl:value-of select="msg:GetApplicantNames(
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@FIRSTFORENAME), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@SURNAME),
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@FIRSTFORENAME), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@SURNAME))"/></xsl:attribute>
		    <xsl:if test="//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE">
				<xsl:attribute name="OFFERISSUEDDATE"><xsl:value-of select="concat('{',msg:GetDate(string(//APPLICATIONOFFER[position()=1]/@OFFERISSUEDATE)),'}')"/></xsl:attribute>
			</xsl:if>
			<!--End of code: DS - EP2_737-->
			
			<!--DRC Epsom Double applicant-->
			<xsl:if test="($CustCount = '1')">
				<xsl:if test="($CustRole = '2')">
						<!--xsl:for-each select="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']">
							<xsl:variable name="Order_No" select=".//@CUSTOMERORDER"/>
							<xsl:if test="($Order_No = '1')">
									<xsl:variable name="SALUTATIONONE" select="msgint:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TYPES), string(.//CUSTOMERVERSION/@TITLE_TEXT), string(.//CUSTOMERVERSION/@TITLEOTHER), string(.//CUSTOMERVERSION/@FIRSTFORENAME),string(.//CUSTOMERVERSION/@SURNAME),1)"/>
									<xsl:variable name="CustomerNo" select="msgint:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),1)"/>									
									<xsl:variable name="TenureType1" select="msgint:SetTenure(.//NEWPROPERTY/@TENURETYPE)"/>									
									<xsl:variable name="AddressName" select="msgint:GetApplicantNameWithInitials(string(.//CUSTOMERVERSION/@TITLE_TYPES), string(.//CUSTOMERVERSION/@TITLE_TEXT), string(.//CUSTOMERVERSION/@TITLEOTHER), string(.//CUSTOMERVERSION/@FIRSTFORENAME), string(.//CUSTOMERVERSION/@SECONDFORENAME),  string(.//CUSTOMERVERSION/@SURNAME),1)"/>

									
							</xsl:if>		
							<xsl:if test="($Order_No = '2')">
									<xsl:variable name="SALUTATIONONE" select="msgint:SetSalutation(string(.//CUSTOMERVERSION/@TITLE_TYPES), string(.//CUSTOMERVERSION/@TITLE_TEXT), string(.//CUSTOMERVERSION/@TITLEOTHER), string(.//CUSTOMERVERSION/@FIRSTFORENAME), string(.//CUSTOMERVERSION/@SURNAME),2)"/>
									<xsl:variable name="CustomerNo" select="msgint:SetCustomerNo(concat('{',.//CUSTOMERVERSION/@CUSTOMERNUMBER,'}'),2)"/>									
									<xsl:variable name="AddressName" select="msgint:GetApplicantNameWithInitials(string(.//CUSTOMERVERSION/@TITLE_TYPES), string(.//CUSTOMERVERSION/@TITLE_TEXT), string(.//CUSTOMERVERSION/@TITLEOTHER), string(.//CUSTOMERVERSION/@FIRSTFORENAME), string(.//CUSTOMERVERSION/@SECONDFORENAME),  string(.//CUSTOMERVERSION/@SURNAME),1)"/>									
							</xsl:if>
						</xsl:for-each-->	
						<xsl:element name="INFORMATION">
								<xsl:element name="CUSTOMERDETAILS">
									<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
									<xsl:attribute name="CURRENTADDRESSTITLE">
									<xsl:value-of select="msgint:DealWithJointTitle(
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@TITLE_TEXT), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@TITLEOTHER), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@FIRSTFORENAME), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@SURNAME),
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@TITLE_TEXT), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@TITLEOTHER), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@FIRSTFORENAME), 
											string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@SURNAME))"/>
								</xsl:attribute>
									<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msgint:DealWithApplicantAddress(
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@FLATNUMBER),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENAME),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENUMBER),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@STREET),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@DISTRICT),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@TOWN),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@COUNTY),
												string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@POSTCODE))"/></xsl:attribute>
									<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
									<xsl:attribute name="APPLICATIONTYPE"><xsl:value-of select="$ApplicationType"/></xsl:attribute>
									<!--xsl:attribute name="CUSTOMERNUMBER"><xsl:value-of select="msgint:GetCustomerNo()"/></xsl:attribute-->									
									<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msgint:DealWithJointSalutation(
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@TITLE_TEXT), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@TITLEOTHER), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@FIRSTFORENAME), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=1]/CUSTOMERVERSION/@SURNAME),
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@TITLE_TEXT), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@TITLEOTHER), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@FIRSTFORENAME), 
										string(//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A'][position()=2]/CUSTOMERVERSION/@SURNAME)
									)"/>
									</xsl:attribute>
									<xsl:if test="//NEWPROPERTYADDRESS/ADDRESS">
										<xsl:element name="PROPERTYADDRESSNODE">
											<xsl:attribute name="PROPERTYADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
														string(//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
														string(//NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
														string(//NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
														string(//NEWPROPERTYADDRESS/ADDRESS/@STREET),
														string(//NEWPROPERTYADDRESS/ADDRESS/@DISTRICT),
														string(//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
														string(//NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
														string(//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
										</xsl:element>
									</xsl:if>
									<xsl:call-template name="APPLICANT">
										<xsl:with-param name="RESPONSE" select="$RESPONSE"/>
									</xsl:call-template>
									<xsl:choose>
										<xsl:when test="string(.//CUSTOMERVERSION/@TITLE_TEXT) != '' and string(.//CUSTOMERVERSION/@SURNAME) !=''">
											<xsl:element name="SINCERELY"/>
										</xsl:when>
										<xsl:otherwise>
											<xsl:element name="FAITHFULLY"/>
										</xsl:otherwise>
									</xsl:choose>	
								<!--DS EP2_737 -->
								<xsl:element name="{msg:DealWithValidationText(string($RESPONSE/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))}"></xsl:element>
								<!--End of code: DS EP2_737 -->
									<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
									<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
									<xsl:if test="($TempCounter > 1)">
										<xsl:element name="PAGEBREAK"/>
									</xsl:if>
								</xsl:element>			
						</xsl:element>
				</xsl:if>		
			</xsl:if>
			<!--Epsom Single Applicant--> 
			<xsl:if test="(($CustCount = '1' and $CustRole = '1') or ($CustCount = '2' and $CustRole = '2'))">
				<xsl:for-each select="//CUSTOMERROLE[@CUSTOMERROLETYPE_TYPES='A']">
					<xsl:element name="INFORMATION">
							<xsl:element name="CUSTOMERDETAILS">
								<xsl:attribute name="CURRENTDATE"><xsl:value-of select="concat('{',msg:GetDate(),'}')"/></xsl:attribute>
								<xsl:attribute name="CURRENTADDRESSTITLE">
									<xsl:value-of select="msg:GetApplicantNameWithInitials(
											string(.//CUSTOMERVERSION/@TITLE_TEXT), 
											string(.//CUSTOMERVERSION/@TITLEOTHER), 
											string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
											string(.//CUSTOMERVERSION/@SURNAME))"/>
								</xsl:attribute>
								<xsl:attribute name="CURRENTADDRESS"><xsl:value-of select="msgint:DealWithApplicantAddress(
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@FLATNUMBER),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENAME),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@BUILDINGORHOUSENUMBER),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@STREET),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@DISTRICT),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@TOWN),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@COUNTY),
											string(.//CUSTOMERADDRESS[@ADDRESSTYPE=1 or @ADDRESSTYPE=2][last()]/ADDRESS/@POSTCODE))"/></xsl:attribute>
								<xsl:attribute name="APPLICATIONNUMBER"><xsl:value-of select="concat('{',//CUSTOMERROLE/@APPLICATIONNUMBER,'}')"/></xsl:attribute>
								<xsl:attribute name="APPLICATIONTYPE"><xsl:value-of select="$ApplicationType"/></xsl:attribute>
								<xsl:attribute name="COAPPLICANT"><xsl:value-of select="msg:DealWithSalutation(
											string(.//CUSTOMERVERSION/@TITLE_TEXT), 
											string(.//CUSTOMERVERSION/@TITLEOTHER), 
											string(.//CUSTOMERVERSION/@FIRSTFORENAME), 
											string(.//CUSTOMERVERSION/@SURNAME))"/>
								</xsl:attribute>
								<xsl:if test="//NEWPROPERTYADDRESS/ADDRESS">
									<xsl:element name="PROPERTYADDRESSNODE">
										<xsl:attribute name="PROPERTYADDRESS"><xsl:value-of select="msg:DealWithStraightAddress(
													string(//NEWPROPERTYADDRESS/ADDRESS/@FLATNUMBER),
													string(//NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENAME),
													string(//NEWPROPERTYADDRESS/ADDRESS/@BUILDINGORHOUSENUMBER),
													string(//NEWPROPERTYADDRESS/ADDRESS/@STREET),
													string(//NEWPROPERTYADDRESS/ADDRESS/@DISTRICT),
													string(//NEWPROPERTYADDRESS/ADDRESS/@TOWN),
													string(//NEWPROPERTYADDRESS/ADDRESS/@COUNTY),
													string(//NEWPROPERTYADDRESS/ADDRESS/@POSTCODE))"/></xsl:attribute>
									</xsl:element>
								</xsl:if>
								<xsl:call-template name="APPLICANT">
									<xsl:with-param name="RESPONSE" select="$RESPONSE"/>
								</xsl:call-template>
								<xsl:choose>
									<xsl:when test="string(.//CUSTOMERVERSION/@TITLE_TEXT) != '' and string(.//CUSTOMERVERSION/@SURNAME) !=''">
										<xsl:element name="SINCERELY"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:element name="FAITHFULLY"/>
									</xsl:otherwise>
								</xsl:choose>
								<!--DS EP2_737 -->
								<xsl:element name="{msg:DealWithValidationText(string($RESPONSE/APPLICATIONFACTFIND/@TYPEOFAPPLICATION_TYPES))}"></xsl:element>
								<!--End of code: DS EP2_737 -->
								<xsl:variable name="Temp" select="msg:UpdateCounter(1)"/>
								<xsl:variable name="TempCounter" select="msg:GetCounter(1)"/>
								<xsl:if test="($TempCounter > 1)">
									<xsl:element name="PAGEBREAK"/>
								</xsl:if>
							</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:if>
	</xsl:template>
	
	<!--============================================================================================================-->
	
</xsl:stylesheet>
 