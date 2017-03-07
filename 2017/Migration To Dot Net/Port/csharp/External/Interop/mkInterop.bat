@echo off
SET InputDll=%1
SET InteropDll=%2
SET Namespace=%3
SET BinDir=C:\Program Files\Microsoft Visual Studio 8\SDK\v2.0\Bin
SET KeyFile=C:\projects\MigrationToDotNet\Port\csharp\bin\core.snk
SET InteropDir=C:\projects\MigrationToDotNet\Port\csharp\Interop
SET References=/reference:%InteropDir%\Interop.comsvcs.dll /reference:%InteropDir%\Interop.msado27.dll /reference:%InteropDir%\Interop.MSVBVM60.dll /reference:%InteropDir%\Interop.msxml4.dll /reference:%InteropDir%\Interop.OmigaToMessageQueueInterface.dll /reference:%InteropDir%\Interop.MessageQueueComponentVC.dll
REM SET References=/reference:%InteropDir%\Interop.comsvcs.dll /reference:%InteropDir%\Interop.msado27.dll /reference:%InteropDir%\Interop.MSVBVM60.dll /reference:%InteropDir%\Interop.msxml2.dll /reference:%InteropDir%\Interop.OmigaToMessageQueueInterface.dll /reference:%InteropDir%\Interop.MessageQueueComponentVC.dll

echo tlbimp %InputDll%
"%BinDir%\tlbimp" %InputDll% /out:%InteropDll% /namespace:%Namespace% /keyfile:%KeyFile% %References%

:End