@ECHO OFF
REM Unregister ODIConverter service
REM
REM Added for EP2_842 
REM
REM See OMIGA 4 DOCUMENTATION $\OMIGA4 PRODUCT\Documentation\Business Logic\ODI\Supporting Documentation\ODI Setup Guide.doc
REM for more information
REM
REM Please note this file assumes D drive installation by default


@ECHO ON
CD "D:\Program Files\Marlborough Stirling\Omiga 4\DLL"
odiconverter -unregserver

pause
