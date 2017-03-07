<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>

<% 
	Response.Buffer = true; 
%>


<%		
	var  sRequest ;
	var intRandNumber ;
	var sCustomerNumber ;
	var sCustomerVersionNumber;
	var sApplNumber ;
		
	intRandNumber = Math.floor((Math.random() * 10 )) 
	switch (intRandNumber)
	{
		case 0:
			sCustomerNumber = '00101451' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003743' ;
			break ;
		case 1:
			sCustomerNumber = '00101435' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003697' ;
			break ;
		case 2:
			sCustomerNumber = '00101419' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003689' ;
			break ;
		case 3:
			sCustomerNumber = '00101397' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = '00003670' ;
			break ;
		case 4:
			sCustomerNumber = '00101370' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = '00003662' ;
			break ;
		case 5:
			sCustomerNumber = '00101354' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = '00003654' ;
			break ;
		case 6:
			sCustomerNumber = '00101311' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003646' ;
			break ;
		case 7:
			sCustomerNumber = '00101281' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003638' ;
			break ;
		case 8:
			sCustomerNumber = '00101265' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003611' ;
			break ;
		case 9:
			sCustomerNumber = '00101249' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = '00003603' ;
			break ;
		case 10:
			sCustomerNumber = '00101222' ;
			sCustomerVersionNumber = '1' ;
			sApplNumber = 'B00003581' ;
			break ;
	}

	sRequest = "<REQUEST USERID=\"PRU\" MACHINEID=\"MSG001840\" UNITID=\"MSG001840\">"
				+ "<APPLICATION>"
				+ "<CUSTOMERNUMBER>" + sCustomerNumber + "</CUSTOMERNUMBER>"
				+ "<CUSTOMERVERSIONNUMBER>" + sCustomerVersionNumber + "</CUSTOMERVERSIONNUMBER>"
				+ "<CUSTOMERROLETYPE></CUSTOMERROLETYPE>"
				+ "<CUSTOMERORDER></CUSTOMERORDER>"
				+ "<APPLICATIONNUMBER>" + sApplNumber + "</APPLICATIONNUMBER>"
				+ "<APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
				+ "<TYPEOFAPPLICATION>1</TYPEOFAPPLICATION>"
				+ "</APPLICATION>"
				+ "</REQUEST>"
		
	var thisApplMgr = new ActiveXObject("omApp.ApplicationManagerBO");		
	sResponse = thisApplMgr.AddCustomerToApplication(sRequest);				
	sResponse = thisApplMgr.DeleteCustomerFromApplication(sRequest);	
	thisApplMgr = null ;
			
%>
</BODY>
</HTML>
