@REM test
@REM this batch process receives X params

@REM lets log what we are doing
echo on

set i=0
for /R %%f IN (*.vbp) DO call typeFileContent.bat %%f
pause
