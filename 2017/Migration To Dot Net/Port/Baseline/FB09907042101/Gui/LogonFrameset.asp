<html>
<TITLE>System Logon</TITLE>
<comment>
Workfile:      LogonFrameset.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Frameset for Omiga 4 logon screens
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date     Description
RF		10/11/99 Created
GD		20/12/00 SYS1742 : file back in use. LHS frame will have width = 0,
				 so that it is apparently not there,
				 as far as user is concerned.
				 
				 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
fixme -file not required - file is now in use - GD 28/12/00 (sys1742)
</comment>

<%@  Language=JavaScript %>

<frameset cols="0,*" rows="*" frameborder="YES" bordercolor="#000066" border="0" framespacing="0"> 
	<frame src="OmigaLogon.asp" id="fraDetails" name="fraDetails" bordercolor="#000066" frameborder="NO" scrolling="YES" marginwidth="0" noresize>
	<frame src="sc010.asp" id="fraMain" name="fraMain" bordercolor="#000066" frameborder="NO" scrolling="YES">
</frameset>

</html>
