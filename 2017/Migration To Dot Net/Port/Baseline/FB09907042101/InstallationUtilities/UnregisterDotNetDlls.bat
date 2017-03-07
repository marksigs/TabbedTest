@ECHO ON
CD D:\Program Files\Marlborough Stirling\Omiga 4\DLLDotNet

SET DOTNETFRAMEWORK=%SystemRoot%\Microsoft.NET\Framework\v1.1.4322
SET REGASM=%DOTNETFRAMEWORK%\regasm.exe
SET REGSVCS=%DOTNETFRAMEWORK%\regsvcs.exe

%REGASM% omAdminCS.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omCardPayment.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omExperianCC.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omExperianHI.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omFDM.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omHomeTrack.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omKnowYourCustomer.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omDC.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omLAU.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omPinMailer.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omDG.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omESurv.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omFirstTitle.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGSVCS% /u omEsurv.dll
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGSVCS% /u omFirstTitle.dll
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omGemini.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE
%REGASM% omVex.dll /unregister
IF %ERRORLEVEL% NEQ 0 PAUSE