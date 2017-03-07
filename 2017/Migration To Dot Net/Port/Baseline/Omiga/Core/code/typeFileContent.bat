@echo on
@echo "The file name is:" %1
set /A i = %i% + 1 
"C:\Program Files\Artinsoft\Visual Basic Companion\vbcompanion.exe" %1 /out %1-%i% /verbose /stats /generateCS
echo %1		%i%	>> c:\projectlist.txt