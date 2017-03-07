@ECHO ON
CD D:\Program Files\Marlborough Stirling\Omiga 4\DLLDotNet

SET DOTNETFRAMEWORK=%SystemRoot%\Microsoft.NET\Framework\v1.1.4322
SET REGASM=%DOTNETFRAMEWORK%\regasm.exe
SET REGSVCS=%DOTNETFRAMEWORK%\regsvcs.exe

%REGASM% omAdminCS.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omCardPayment.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omExperianCC.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omExperianHI.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omFDM.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omHomeTrack.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omKnowYourCustomer.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omDC.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omLAU.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omPinMailer.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omDG.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omESurv.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omFirstTitle.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGSVCS% omEsurv.dll
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGSVCS% omFirstTitle.dll
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omGemini.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omVex.dll /codebase
IF %ERRORLEVEL% NEQ 0 PAUSE
