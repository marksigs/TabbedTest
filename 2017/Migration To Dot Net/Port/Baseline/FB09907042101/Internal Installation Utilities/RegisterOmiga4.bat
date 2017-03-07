@echo off

REM Filename: 	RegisterOmiga4.bat
REM
REM Description: 
REM	Batch file that registers Omiga 4 external components.
REM ------------------------------------------------------------------------------------------
REM History:
REM
REM Prog  Date     Description
REM RF    09/12/99 Created
REM AS    07/02/00 Added PolarisInterface and BankDetailsInterface
REM ------------------------------------------------------------------------------------------

echo Starting RegisterOmiga4.bat version 07/02/2000 ...

@echo on

regsvr32 "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\MortgageRepayment.dll"
regsvr32 "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\PolarisInterface.dll"
regsvr32 "C:\Program Files\Marlborough Stirling\Omiga 4\Dll\BankDetailsInterface.dll"
