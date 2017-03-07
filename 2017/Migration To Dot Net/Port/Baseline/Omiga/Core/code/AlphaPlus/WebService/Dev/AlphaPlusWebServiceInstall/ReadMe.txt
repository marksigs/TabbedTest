Installation notes for Alpha Plus Web Service installation
==========================================================

If you are installing via Terminal Services onto a Microsoft Windows 2000 machine 
which has Microsoft Windows Installer version 1.2 or below, you will need to copy 
the installation files to a local drive before installing (see Microsoft Knowledge 
Base Article 255582, “BUG: Error Running Windows Installer Installation on Terminal 
Server” at http://support.microsoft.com/default.aspx?scid=kb;EN-US;255582 for more 
information).

In versions 1.1 and earlier, compilation debug is set to "true" in the installed 
web.config file (by default, C:\Inetpub\wwwroot\AlphaPlusWebService\Web.config). For 
faster execution this can be set to "false". The intended setting for future releases 
is "false".

To test an installation, you can view the default test page (assuming default 
installation options are chosen) at 

http://localhost/alphapluswebservice/alphaplussoaprpc.asmx 

and at 

http://localhost/alphapluswebservice/alphaplussoapdoc.asmx.




