@echo off

REM Filename: 	UnregisterOmiga4.bat
REM
REM Description: 
REM	Batch file that unregisters Omiga 4 dlls.
REM ------------------------------------------------------------------------------------------
REM History:
REM
REM Prog  Date     Description
REM RF    08/12/99 Created
REM AS    07/02/00 Added MortgageRepayment.dll, PolarisInterface and BankDetailsInterface
REM ------------------------------------------------------------------------------------------

echo Starting UnregisterOmiga4.bat version 07/02/2000 ...

@echo on

regsvr32 /u /c "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\MortgageRepayment.dll"
regsvr32 /u /c "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\PolarisInterface.dll"
regsvr32 /u /c "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\BankDetailsInterface.dll"

REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\Audit\omAU\omAU.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\Base\omBase\omBase.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\CostModelling\omCM\omCM.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\custreg\middletier\omiga4CustReg\omCR.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\MortgageProduct\omMP\omMP.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\QuickQuote\omiga4QuickQuote\omQQ.dll"
REM regsvr32 /u /c "C:\PROJECTS\Omiga4\dev\Organisation\omOrg\omOrg.dll"

REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omAU.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omBase.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omCM.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omCR.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omMP.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omQQ.dll"
REM regsvr32 /u /c "C:\Program Files\omiga4dlls\omOrg.dll"
